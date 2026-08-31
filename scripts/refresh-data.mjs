#!/usr/bin/env node
/**
 * 各営業グループ台帳から実績・積上げ・計画を集計し、JSONに書き出す。
 *
 * 守っていること（過去の教訓：列の埋まり方から実績/見込みを自己流で推測しない）:
 *  - 実績+積上げは、各担当者ブロックの「合　計（売上＋積上げ）」小計行をそのまま使う
 *    （個別受注明細の再集計はしない）。
 *  - 計画値・当月実績・積上げは、各グループ台帳シート上部のヘッダー行
 *    （「計画売上」「実績売上」「積上売上」等、row4-6 col3/col5）からも取得し、
 *    担当者ブロック小計の合算値と突き合わせる（一致しない場合は警告）。
 *  - 前年実績（当月分）は、2025年度の同一グループ台帳・同月シートの同じヘッダー行
 *    から直接取得する。
 *  - ロールアップ台帳「各グループ管理台帳合計」は使わない（数式キャッシュが古いタイミングの
 *    まま更新されないことがあり、当月の実績・計画の値として信頼できないため。2026-08時点で
 *    ユーザーからも「今後は参照不要」と明示された）。
 *  - 久田氏・小串氏の担当グループ変更に伴う前年実績の按分調整額は、台帳から自動導出
 *    せず、下記 MANUAL_CONFIG に明示的な設定値として保持する（要・人手更新）。
 *  - 4〜7月（既に終了している月）の確定実績・計画・前年実績は(20260731)販管システム出力を正とし、
 *    このスクリプトの出力ではなく既存のindex.html記載値をそのまま維持する。このスクリプトが
 *    実質使われるのは当月（進行中の月）の実績・積上げ・計画・前年同月実績のみ。
 *
 * 使い方: node scripts/refresh-data.mjs [--asof=YYYY-MM-DD]
 */
import XLSX from "xlsx";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const LEDGER_DIR_CURRENT = "O:/全社共有/宣伝/★営業週報2026/各グループ会議資料/";
const LEDGER_DIR_PREV = "O:/全社共有/宣伝/★営業週報2025/各グループ会議資料/";

const asOfArg = process.argv.find((a) => a.startsWith("--asof="));
const asOf = asOfArg ? asOfArg.split("=")[1] : new Date().toISOString().slice(0, 10);
const asOfDate = new Date(asOf + "T00:00:00+09:00");
const currentMonth = asOfDate.getMonth() + 1; // JSTベース想定
const YTD_MONTHS = [4, 5, 6, 7, 8]; // 通期の4〜8月累計（2026.8.31に8月が確定したため追加。確定値は販管システム出力を正とし更新しない）

// --- 手動設定（台帳から自動導出しない値） -------------------------------
const MANUAL_CONFIG = {
  note:
    "久田氏（BOATRACE振興会案件, 第1G→第2G）・小串氏（退職。南都/沖縄長正薬草本社/イーアス沖縄 第1G→第3G、佐久本工機 第1G→第2G）の" +
    "担当グループ変更に伴う前年実績の付け替え額。台帳の生数値からは自動で読み取れないため、現行ダッシュボード掲載の内訳をそのまま転記した設定値。" +
    "台帳側で新たな担当変更が発生した場合はここを手動更新すること。",
  staffTransferAdjustments: {
    // 単位：円（千円ではない）。sign: g1からの控除はg1側マイナス、移管先はプラス。
    g1: { salesDelta: -(30885000 + 2305000 + 2841000), profitDelta: -(6425000 + 763000 + 862000) },
    g2: { salesDelta: +(30885000 + 2841000), profitDelta: +(6425000 + 862000) },
    g3: { salesDelta: +2305000, profitDelta: +763000 }
  }
};

// --- ユーティリティ --------------------------------------------------------
function normalizeDigits(s) {
  return String(s).replace(/[０-９]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0xfee0));
}

function findMonthSheet(wb, month) {
  const target = `${month}月`;
  return wb.SheetNames.find((n) => {
    const norm = normalizeDigits(n).trim();
    return (
      norm.endsWith(target) &&
      !norm.includes("計画") &&
      !norm.includes("クオーター") &&
      !norm.includes("個別") &&
      !norm.includes("クライアント")
    );
  });
}

function cell(ws, r, c) {
  const v = ws[XLSX.utils.encode_cell({ r, c })];
  return v ? v.v : undefined;
}

// 担当者ブロックの「合　計（売上＋積上げ）」小計行をすべて合算する
// （個別受注明細を再集計しない＝過去の教訓に従う）
function sumStaffTotals(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws["!ref"]);
  let sales = 0,
    profit = 0,
    blocks = 0;
  for (let r = range.s.r; r <= range.e.r; r++) {
    const label = cell(ws, r, 1);
    if (typeof label === "string" && label.includes("合　計（売上＋積上げ）")) {
      sales += Number(cell(ws, r, 6)) || 0;
      profit += Number(cell(ws, r, 7)) || 0;
      blocks++;
    }
  }
  return { sales, profit, blocks };
}

// シート上部のヘッダー行（計画売上/実績売上/計画利益/実績利益/積上売上/積上利益）を読む。
// row4: 計画売上(col3) / 実績売上(col5)、row5: 計画利益(col3) / 実績利益(col5)、
// row6: 積上売上(col3) / 積上利益(col5)。2025年度・2026年度いずれの台帳も同一レイアウト。
function readHeader(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  return {
    planSales: Number(cell(ws, 4, 3)) || 0,
    actualSales: Number(cell(ws, 4, 5)) || 0,
    planProfit: Number(cell(ws, 5, 3)) || 0,
    actualProfit: Number(cell(ws, 5, 5)) || 0,
    pipelineSales: Number(cell(ws, 6, 3)) || 0,
    pipelineProfit: Number(cell(ws, 6, 5)) || 0
  };
}

function ledgerMonthly(file, months) {
  const wb = XLSX.readFile(LEDGER_DIR_CURRENT + file);
  const out = {};
  for (const m of months) {
    const sheetName = findMonthSheet(wb, m);
    if (!sheetName) {
      out[m] = { ok: false, error: `${m}月のシートが見つかりません` };
      continue;
    }
    try {
      const { sales, profit, blocks } = sumStaffTotals(wb, sheetName);
      const header = readHeader(wb, sheetName);
      out[m] = { ok: true, sheetName, sales, profit, blocks, header };
    } catch (e) {
      out[m] = { ok: false, error: String(e && e.message) };
    }
  }
  return out;
}

function prevYearActual(file, month) {
  const wb = XLSX.readFile(LEDGER_DIR_PREV + file);
  const sheetName = findMonthSheet(wb, month);
  if (!sheetName) return { ok: false, error: `${month}月のシートが見つかりません（前年度台帳）` };
  const header = readHeader(wb, sheetName);
  return { ok: true, sheetName, sales: header.actualSales, profit: header.actualProfit };
}

// --- 実行 --------------------------------------------------------------
const GROUP_FILES_CURRENT = {
  g1: "第1営業グループ管理台帳・会議資料【2026年度】.xlsx",
  g2: "第2営業グループ管理台帳・会議資料【2026年度】.xlsx",
  g3: "第3営業グループ管理台帳・会議資料【2026年度】.xlsx"
};
const GROUP_FILES_PREV = {
  g1: "第1営業グループ管理台帳・会議資料【2025年度】.xlsx",
  g2: "第2営業グループ管理台帳・会議資料【2025年度】.xlsx",
  g3: "第3営業グループ管理台帳・会議資料【2025年度】.xlsx"
};

const allMonths = Array.from(new Set([...YTD_MONTHS, currentMonth]));

console.log(`\n=== refresh-data.mjs 実行（本日時点: ${asOf}, 当月: ${currentMonth}月） ===\n`);
console.log("※4〜7月（既に終了している月）の確定実績・計画・前年実績は(20260731)販管システム出力を正とし、この出力では更新しません。実質使うのは当月分のみです。\n");

const ledgerByGroup = {};
const prevYearByGroup = {};
for (const [g, file] of Object.entries(GROUP_FILES_CURRENT)) {
  console.log(`-- ${g} 台帳を読み込み中 (${file})`);
  ledgerByGroup[g] = ledgerMonthly(file, allMonths);
  for (const m of allMonths) {
    const r = ledgerByGroup[g][m];
    if (!r.ok) { console.log(`   ${m}月: !! ${r.error}`); continue; }
    console.log(`   ${m}月: シート"${r.sheetName}" 担当者ブロック${r.blocks}件 売上=${r.sales.toLocaleString()} 利益=${r.profit.toLocaleString()}`);
    const h = r.header;
    // r.sales/r.profit は「合計（売上＋積上げ）」小計行の合算＝着地見込み値なので、
    // ヘッダー側も 実績＋積上げ で比較する（実績のみの列と比べると必ず不一致になるため）。
    const subtotalMatchesHeader =
      h.actualSales + h.pipelineSales === r.sales && h.actualProfit + h.pipelineProfit === r.profit;
    console.log(
      `        [ヘッダー] 計画売上=${h.planSales.toLocaleString()} 計画利益=${h.planProfit.toLocaleString()} ` +
      `積上売上=${h.pipelineSales.toLocaleString()} 積上利益=${h.pipelineProfit.toLocaleString()} ` +
      (subtotalMatchesHeader ? "(担当者ブロック集計と一致)" : "!! 担当者ブロック集計と不一致・要確認")
    );
  }
  const prevFile = GROUP_FILES_PREV[g];
  console.log(`-- ${g} 前年度台帳を読み込み中 (${prevFile})、当月(${currentMonth}月)の前年実績のみ取得`);
  const pv = prevYearActual(prevFile, currentMonth);
  prevYearByGroup[g] = pv;
  if (pv.ok) {
    console.log(`   前年${currentMonth}月: シート"${pv.sheetName}" 売上=${pv.sales.toLocaleString()} 利益=${pv.profit.toLocaleString()} （担当者異動の按分調整は未反映の生値）`);
  } else {
    console.log(`   !! ${pv.error}`);
  }
}

// --- 出力 ------------------------------------------------------------------
const output = {
  meta: {
    asOf,
    generatedAt: new Date().toISOString(),
    currentMonth,
    ytdMonths: YTD_MONTHS,
    note: "ロールアップ台帳（各グループ管理台帳合計）は参照しない方針（2026-08時点でユーザー指示）。計画・当月実績・積上げは各グループ台帳シート上部のヘッダー行、前年実績は前年度台帳の同じヘッダー行から直接取得。"
  },
  ledgerComputed: ledgerByGroup, // G1/G2/G3のみ。担当者ブロック小計行の直接集計＋シートヘッダー行（計画/実績/積上げ）
  prevYearActual: prevYearByGroup, // G1/G2/G3のみ。前年度台帳の当月ヘッダーから取得した生値（按分調整は未反映）
  manualConfig: MANUAL_CONFIG
};

const outDir = path.join(__dirname, "output");
fs.mkdirSync(outDir, { recursive: true });
const outPath = path.join(outDir, `data-${asOf}.json`);
fs.writeFileSync(outPath, JSON.stringify(output, null, 2), "utf8");
console.log(`\n出力: ${outPath}`);
