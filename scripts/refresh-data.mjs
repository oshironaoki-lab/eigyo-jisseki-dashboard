#!/usr/bin/env node
/**
 * 各営業グループ台帳から実績・積上げ・計画を集計し、JSONに書き出す。
 *
 * 守っていること（過去の教訓：列の埋まり方から実績/見込みを自己流で推測しない）:
 *  - 実績+積上げは、各担当者ブロックの「合　計（売上＋積上げ）」小計行をそのまま使う
 *    （個別受注明細の再集計はしない）。
 *  - 上記の台帳集計値は、「各グループ管理台帳合計【2026年度】」ロールアップの値と
 *    必ず突合する。差異が閾値を超える場合は自動で解決せず、両方の値をJSONに残して
 *    reconciliation に警告として記録する（人が判断する）。
 *  - 計画数値は、ロールアップの「2026目標」ブロックを正とする（各グループ自身の
 *    計画実績サマリーシートの値と一致することを確認済み）。
 *  - 前年実績は、ロールアップ内の「2025コンテンツ含む計」シートの
 *    「2025実数(積上込)」ブロックを使う。
 *  - 久田氏・小串氏の担当グループ変更に伴う前年実績の按分調整額は、台帳から自動導出
 *    せず、下記 MANUAL_CONFIG に明示的な設定値として保持する（要・人手更新）。
 *  - 全社／コンテンツの「計画」はロールアップに存在するのでそこから取得できるが、
 *    ロールアップ自体が販管システムの正式出力と完全一致する保証はないため、
 *    ダッシュボードに反映する前に必ずこのスクリプトの突合レポートを確認すること。
 *
 * 使い方: node scripts/refresh-data.mjs [--asof=YYYY-MM-DD]
 */
import XLSX from "xlsx";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const LEDGER_DIR = "O:/全社共有/宣伝/★営業週報2026/各グループ会議資料/";

const asOfArg = process.argv.find((a) => a.startsWith("--asof="));
const asOf = asOfArg ? asOfArg.split("=")[1] : new Date().toISOString().slice(0, 10);
const asOfDate = new Date(asOf + "T00:00:00+09:00");
const currentMonth = asOfDate.getMonth() + 1; // JSTベース想定
const YTD_MONTHS = [4, 5, 6, 7]; // 通期の4〜7月累計（本ダッシュボードの既定レンジ）

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

function ledgerMonthly(groupKey, file, months) {
  const wb = XLSX.readFile(LEDGER_DIR + file);
  const out = {};
  for (const m of months) {
    const sheetName = findMonthSheet(wb, m);
    if (!sheetName) {
      out[m] = { ok: false, error: `${m}月のシートが見つかりません` };
      continue;
    }
    try {
      const { sales, profit, blocks } = sumStaffTotals(wb, sheetName);
      out[m] = { ok: true, sheetName, sales, profit, blocks };
    } catch (e) {
      out[m] = { ok: false, error: String(e && e.message) };
    }
  }
  return out;
}

// ロールアップワークブック（各グループ管理台帳合計）の「2026目標」「2026実数(積上込)」
// と、同ファイル内の「2025コンテンツ含む計」シートの「2025実数(積上込)」を読む。
// レイアウトは事前調査で確認済み（列は四半期小計・年計を挟むため月→列の対応表が必要）。
const MONTH_COL = { 4: 2, 5: 3, 6: 4, 7: 6, 8: 7, 9: 8, 10: 10, 11: 11, 12: 12, 1: 14, 2: 15, 3: 16 };
// 「目標」ブロックはグループごとに2行（売上・粗利のみ）。
const PLAN_GROUP_ROW = { g1: 1, g2: 3, g3: 5, content: 7, all: 9 };
// 「実数(積上込)」ブロックはグループごとに4行（売上・粗利・売上達成率・粗利達成率）を含むため、
// PLAN_GROUP_ROW に一定オフセットを足すのではなく実際の行番号を明示する（事前調査で確認済み）。
const ACTUAL_GROUP_ROW = { g1: 13, g2: 17, g3: 21, content: 25, all: 29 };

function readRollupBlock(ws, groupRowBase, months) {
  const out = {};
  for (const m of months) {
    const col = MONTH_COL[m];
    out[m] = {
      sales: Number(cell(ws, groupRowBase, col)) || 0,
      profit: Number(cell(ws, groupRowBase + 1, col)) || 0
    };
  }
  return out;
}

function readRollup(months) {
  const wb = XLSX.readFile(LEDGER_DIR + "各グループ管理台帳合計【2026年度】.xlsx");
  const ws = wb.Sheets["2026"];
  const result = {};
  for (const g of Object.keys(PLAN_GROUP_ROW)) {
    result[g] = {
      plan: readRollupBlock(ws, PLAN_GROUP_ROW[g], months),
      actualWithPipeline: readRollupBlock(ws, ACTUAL_GROUP_ROW[g], months)
    };
  }
  return result;
}

function readPrevYear(months) {
  const wb = XLSX.readFile(LEDGER_DIR + "各グループ管理台帳合計【2026年度】.xlsx");
  const ws = wb.Sheets["2025コンテンツ含む計"];
  const result = {};
  for (const g of Object.keys(ACTUAL_GROUP_ROW)) {
    result[g] = readRollupBlock(ws, ACTUAL_GROUP_ROW[g], months);
  }
  return result;
}

// --- 実行 --------------------------------------------------------------
const GROUP_FILES = {
  g1: "第1営業グループ管理台帳・会議資料【2026年度】.xlsx",
  g2: "第2営業グループ管理台帳・会議資料【2026年度】.xlsx",
  g3: "第3営業グループ管理台帳・会議資料【2026年度】.xlsx"
};

const allMonths = Array.from(new Set([...YTD_MONTHS, currentMonth]));

console.log(`\n=== refresh-data.mjs 実行（本日時点: ${asOf}, 当月: ${currentMonth}月） ===\n`);

const ledgerByGroup = {};
for (const [g, file] of Object.entries(GROUP_FILES)) {
  console.log(`-- ${g} 台帳を読み込み中 (${file})`);
  ledgerByGroup[g] = ledgerMonthly(g, file, allMonths);
  for (const m of allMonths) {
    const r = ledgerByGroup[g][m];
    if (r.ok) console.log(`   ${m}月: シート"${r.sheetName}" 担当者ブロック${r.blocks}件 売上=${r.sales.toLocaleString()} 利益=${r.profit.toLocaleString()}`);
    else console.log(`   ${m}月: !! ${r.error}`);
  }
}

console.log(`\n-- ロールアップ（各グループ管理台帳合計）を読み込み中`);
const rollup = readRollup(allMonths);
console.log(`-- 前年実績（2025コンテンツ含む計シート）を読み込み中`);
const prevYear = readPrevYear(allMonths);

// --- 突合 ----------------------------------------------------------------
const TOLERANCE_PCT = 0.5; // これを超える差異は要確認として明示する
const reconciliation = [];
for (const g of Object.keys(GROUP_FILES)) {
  for (const m of allMonths) {
    const led = ledgerByGroup[g][m];
    if (!led.ok) continue;
    const rb = rollup[g].actualWithPipeline[m];
    for (const metric of ["sales", "profit"]) {
      const a = led[metric];
      const b = rb[metric];
      const diff = a - b;
      const pct = b !== 0 ? (Math.abs(diff) / Math.abs(b)) * 100 : a === 0 ? 0 : 100;
      const entry = { group: g, month: m, metric, ledgerValue: a, rollupValue: b, diff, diffPct: Number(pct.toFixed(2)) };
      if (pct > TOLERANCE_PCT) {
        entry.flag = "MISMATCH_NEEDS_REVIEW";
        reconciliation.push(entry);
      }
    }
  }
}

console.log(`\n=== 突合レポート（台帳直接集計 vs ロールアップ, 差異${TOLERANCE_PCT}%超のみ表示） ===`);
if (reconciliation.length === 0) {
  console.log("差異なし（すべて許容範囲内で一致）。");
} else {
  for (const r of reconciliation) {
    console.log(
      `  [要確認] ${r.group} ${r.month}月 ${r.metric}: 台帳集計=${r.ledgerValue.toLocaleString()} / ロールアップ=${r.rollupValue.toLocaleString()} (差 ${r.diff.toLocaleString()}, ${r.diffPct}%)`
    );
  }
  console.log(
    "\n  → 上記は自動では解決しません。どちらの数値源（担当者ブロック台帳集計 or ロールアップの数式値）を採用するか、" +
      "人が確認のうえ index.html への反映時に選んでください。"
  );
}

// --- 出力 ------------------------------------------------------------------
const output = {
  meta: {
    asOf,
    generatedAt: new Date().toISOString(),
    currentMonth,
    ytdMonths: YTD_MONTHS,
    toleranceUsedForReconciliation: `${TOLERANCE_PCT}%`
  },
  ledgerComputed: ledgerByGroup, // G1/G2/G3のみ。担当者ブロック小計行を直接集計した値（最も追跡しやすい一次値）
  rollup, // 各グループ+全社+コンテンツ。「2026目標」（計画）と「2026実数(積上込)」（実績+積上げ）
  prevYear, // 各グループ+全社+コンテンツ。前年（2025）実数(積上込)、按分調整は未反映の生値
  reconciliation,
  manualConfig: MANUAL_CONFIG
};

const outDir = path.join(__dirname, "output");
fs.mkdirSync(outDir, { recursive: true });
const outPath = path.join(outDir, `data-${asOf}.json`);
fs.writeFileSync(outPath, JSON.stringify(output, null, 2), "utf8");
console.log(`\n出力: ${outPath}`);
