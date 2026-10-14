/**
 * 月次売上報告ツール — Excel の読み書き（exceljs）
 * 読み: 先頭のシートをセルの表（Grid）にする。書き: Python 版と同じ 6 シート構成の xlsx を作る。
 * exceljs は大きいので、使うときに読み込む（dynamic import）。
 *
 * Copyright (c) 2026 OFFREL Inc.
 * MIT License（https://opensource.org/licenses/MIT）
 * このコードは動作の保証をしていません。結果はご自身でお確かめください。
 */
import type ExcelJSNS from "exceljs";
import { ymdToDate } from "./normalize";
import type { CellValue, Grid, ReportModel } from "./types";

type ExcelJS = typeof ExcelJSNS;

async function loadExcelJS(): Promise<ExcelJS> {
  const mod = (await import("exceljs")) as unknown as { default?: ExcelJS } & ExcelJS;
  return (mod.default ?? mod) as ExcelJS;
}

/** exceljs のセル値を素の値に落とす（数式は結果、リッチテキストは文字、ハイパーリンクは文字） */
export function toCellValue(v: ExcelJSNS.CellValue): CellValue {
  if (v === null || v === undefined) return null;
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return v;
  if (v instanceof Date) return v;
  if (typeof v === "object") {
    if ("richText" in v && Array.isArray(v.richText)) return v.richText.map((t) => t.text).join("");
    if ("result" in v) {
      const r = (v as ExcelJSNS.CellFormulaValue).result;
      if (r === undefined || r === null) return null;
      if (typeof r === "object" && "error" in r) return null;
      return r as CellValue;
    }
    if ("text" in v) {
      const t = (v as ExcelJSNS.CellHyperlinkValue).text;
      return typeof t === "string" ? t : toCellValue(t as ExcelJSNS.CellValue);
    }
    if ("error" in v) return null;
  }
  return String(v);
}

/** xlsx（ArrayBuffer）の先頭のシートを Grid にする。grid[r][c] は r+1 行目・c+1 列目 */
export async function readGrid(data: ArrayBuffer): Promise<Grid> {
  const ExcelJS = await loadExcelJS();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(data);
  const ws = wb.worksheets[0];
  if (!ws) return [];
  const grid: Grid = [];
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const values = row.values as ExcelJSNS.CellValue[]; // 1 始まりの配列（[0] は空）
    const cells: CellValue[] = [];
    for (let c = 1; c < values.length; c++) cells.push(toCellValue(values[c]));
    while (grid.length < rowNumber - 1) grid.push([]);
    grid[rowNumber - 1] = cells;
  });
  return grid;
}

export async function readGridFromFile(file: File): Promise<Grid> {
  return readGrid(await file.arrayBuffer());
}

/* ---------------------------------------------------------------- 書き出し */

const HEAD_FILL: ExcelJSNS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF7" } };
const WARN_FILL: ExcelJSNS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };

type Row = Array<string | number | Date | null>;

function putTable(
  ws: ExcelJSNS.Worksheet,
  headers: string[],
  rows: Row[],
  opts: { widths?: number[]; numberCols?: number[]; pctCols?: number[]; totalRow?: Row },
) {
  const h = ws.addRow(headers);
  h.eachCell((c) => {
    c.font = { bold: true };
    c.fill = HEAD_FILL;
  });
  for (const r of rows) ws.addRow(r);
  if (opts.totalRow) {
    const t = ws.addRow(opts.totalRow);
    t.eachCell((c) => {
      c.font = { bold: true };
    });
  }
  const numberCols = new Set(opts.numberCols ?? []);
  const pctCols = new Set(opts.pctCols ?? []);
  ws.eachRow((row, rowNumber) => {
    if (rowNumber < 2) return;
    row.eachCell({ includeEmpty: true }, (c, colNumber) => {
      if (numberCols.has(colNumber)) c.numFmt = "#,##0";
      else if (pctCols.has(colNumber)) c.numFmt = "0.0%";
    });
  });
  (opts.widths ?? []).forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });
  ws.views = [{ state: "frozen", ySplit: 1 }];
}

function cmpRows(rows: ReportModel["customers"] | ReportModel["staff"]): Row[] {
  return rows.map((r) => [r.key, r.name, r.amount, r.prevAmount, r.diff, r.ratio]);
}

function nowText(): string {
  const d = new Date();
  const p = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

/** 報告を xlsx（バイナリ）にする。シート: サマリー / 得意先別 / 商品分類別 / 担当者別 / 要確認 / 統合明細 */
export async function writeReportWorkbook(m: ReportModel): Promise<ArrayBuffer> {
  const ExcelJS = await loadExcelJS();
  const wb = new ExcelJS.Workbook();
  const tc = m.total.amount;
  const tp = m.total.prevAmount;
  const tDiff = m.total.diff;
  const tRatio = m.total.ratio;

  // --- サマリー
  {
    const ws = wb.addWorksheet("サマリー");
    ws.addRow([`月次売上報告 ${m.ym.slice(0, 4)}年${Number(m.ym.slice(5, 7))}月`]).getCell(1).font = { bold: true, size: 14 };
    ws.addRow([`作成: ${nowText()}   前月: ${m.prevYm}${m.hasPrev ? "" : "（ファイルなし・前月比は空欄）"}`]);
    ws.addRow([]);
    const h = ws.addRow(["営業所", "ファイル", "明細の行数", "集計した行", "要確認の行", "読み飛ばし", "当月金額", "前月金額", "増減額", "前月比"]);
    h.eachCell((c) => {
      c.font = { bold: true };
      c.fill = HEAD_FILL;
    });
    for (const s of m.summary) {
      ws.addRow([s.office, s.file, s.rows, s.included, s.flagged, s.skipped, s.amount, s.prevAmount, s.diff, s.ratio]);
    }
    const t = ws.addRow(["合計", "", m.total.rows, m.total.included, m.total.flagged, "", tc, tp, tDiff, tRatio]);
    t.eachCell((c) => {
      c.font = { bold: true };
    });
    ws.addRow([]);
    ws.addRow(["要確認の内訳（理由の先頭の語で数えた件数）"]).getCell(1).font = { bold: true };
    for (const r of m.reasonCounts) ws.addRow([r.reason, r.count]);
    ws.addRow([]);
    ws.addRow(["注: 要確認の行のうち「同じ内容の行」だけは集計に含めている（本当に2回売れた可能性があるため）。"]);
    ws.addRow(["注: それ以外の要確認の行は集計に含めていない。直したら、ファイルを置き直して再実行する。"]);
    for (let r = 5; r <= 5 + m.summary.length; r++) {
      const row = ws.getRow(r);
      for (const c of [3, 4, 5, 7, 8, 9]) row.getCell(c).numFmt = "#,##0";
      row.getCell(10).numFmt = "0.0%";
    }
    [14, 26, 11, 11, 11, 24, 14, 14, 12, 9].forEach((w, i) => {
      ws.getColumn(i + 1).width = w;
    });
  }

  // --- 得意先別
  putTable(wb.addWorksheet("得意先別"), ["得意先コード", "得意先名", "当月金額", "前月金額", "増減額", "前月比"], cmpRows(m.customers), {
    widths: [12, 28, 14, 14, 12, 9],
    numberCols: [3, 4, 5],
    pctCols: [6],
    totalRow: ["合計", "", tc, tp, tDiff, tRatio],
  });

  // --- 商品分類別
  putTable(
    wb.addWorksheet("商品分類別"),
    ["商品分類", "当月数量", "当月金額", "前月金額", "増減額", "前月比"],
    m.categories.map((r) => [r.key, r.qty ?? 0, r.amount, r.prevAmount, r.diff, r.ratio]),
    {
      widths: [12, 10, 14, 14, 12, 9],
      numberCols: [2, 3, 4, 5],
      pctCols: [6],
      totalRow: ["合計", m.totalQty, tc, tp, tDiff, tRatio],
    },
  );

  // --- 担当者別
  putTable(wb.addWorksheet("担当者別"), ["営業所", "担当者", "当月金額", "前月金額", "増減額", "前月比"], cmpRows(m.staff), {
    widths: [12, 12, 14, 14, 12, 9],
    numberCols: [3, 4, 5],
    pctCols: [6],
    totalRow: ["合計", "", tc, tp, tDiff, tRatio],
  });

  // --- 要確認
  {
    const ws = wb.addWorksheet("要確認");
    putTable(
      ws,
      ["営業所", "ファイル", "行", "理由", "日付", "得意先コード", "得意先名（ファイルの表記）", "商品コード", "商品名", "数量", "単価", "金額", "数量×単価", "集計に"],
      m.checks.map((r) => [
        r.office,
        r.file,
        r.row,
        r.reasons,
        r.date,
        r.custCodeRaw,
        r.custNameRaw,
        r.productCode,
        cellForExcel(r.productName),
        r.qty,
        r.unit,
        r.amount,
        r.calc,
        r.included ? "含む" : "含まない",
      ]),
      { widths: [12, 24, 6, 60, 11, 12, 26, 10, 24, 8, 9, 11, 11, 8], numberCols: [10, 11, 12, 13] },
    );
    ws.eachRow((row, rowNumber) => {
      if (rowNumber < 2) return;
      const c = row.getCell(4);
      c.fill = WARN_FILL;
      c.alignment = { wrapText: true, vertical: "top" };
    });
  }

  // --- 統合明細
  {
    const ws = wb.addWorksheet("統合明細");
    putTable(
      ws,
      ["営業所", "日付", "得意先コード", "得意先名（マスタ）", "商品コード", "商品名", "商品分類", "数量", "単価", "金額", "担当者", "元ファイル", "元の行"],
      m.details.map((r) => [
        r.office,
        r.date ? ymdToDate(r.date) : null,
        r.custCode,
        r.custName,
        r.productCode,
        cellForExcel(r.productName),
        r.category,
        r.qty,
        r.unit,
        r.amount,
        r.staff,
        r.file,
        r.row,
      ]),
      { widths: [12, 11, 12, 26, 10, 26, 10, 8, 9, 11, 10, 24, 8], numberCols: [8, 9, 10] },
    );
    ws.eachRow((row, rowNumber) => {
      if (rowNumber < 2) return;
      row.getCell(2).numFmt = "yyyy-mm-dd";
    });
  }

  return (await wb.xlsx.writeBuffer()) as unknown as ArrayBuffer;
}

function cellForExcel(v: CellValue): string | number | Date | null {
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  return v;
}
