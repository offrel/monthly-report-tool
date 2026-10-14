/**
 * 月次売上報告ツール — 見出し行の検出と列の推定
 *
 * Copyright (c) 2026 OFFREL Inc.
 * MIT License（https://opensource.org/licenses/MIT）
 * このコードは動作の保証をしていません。結果はご自身でお確かめください。
 */
import { normHeader, officeNameFromFile } from "./normalize";
import { FIELDS, REQUIRED_FIELDS, type ColumnMapping, type Field, type FileSetup, type Grid, type SheetData } from "./types";

/**
 * 見出しの文字 → 内部の項目名。営業所ごとに見出しが違うので、ここで吸収する。
 * Python 版の COLUMN_ALIASES と同じ。新しい書き方が出てきたら、この表に足す。
 */
export const COLUMN_ALIASES: Record<Field, string[]> = {
  日付: ["日付", "売上日", "伝票日付", "年月日"],
  得意先コード: ["得意先コード", "得意先cd", "得意先ｺｰﾄﾞ", "取引先コード", "顧客コード"],
  得意先名: ["得意先名", "得意先", "取引先名", "取引先", "顧客名"],
  商品コード: ["商品コード", "商品cd", "商品ｺｰﾄﾞ", "品番"],
  商品名: ["商品名", "品名", "商品"],
  数量: ["数量", "数", "個数"],
  単価: ["単価", "売上単価"],
  金額: ["金額", "金額(税抜)", "売上金額", "売上", "税抜金額"],
  担当者: ["担当者", "担当", "担当者名", "営業担当"],
};

const NORMALIZED_ALIASES: Array<{ field: Field; aliases: string[] }> = FIELDS.map((field) => ({
  field,
  aliases: COLUMN_ALIASES[field].map(normHeader),
}));

/** 見出し行を探す範囲（先頭から何行） */
export const HEADER_SCAN_ROWS = 10;

/** 1 行が見出しらしい度合い（項目の語に一致するセルの数） */
export function scoreHeaderRow(row: readonly unknown[]): number {
  let score = 0;
  for (const cell of row) {
    const h = normHeader(cell);
    if (!h) continue;
    const hit = NORMALIZED_ALIASES.some(({ aliases }) => aliases.some((a) => h === a || h.includes(a)));
    if (hit) score += 1;
  }
  return score;
}

/**
 * 見出し行（1 始まり）を自動で見つける。
 * 先頭 HEADER_SCAN_ROWS 行のうち、「日付」「金額」などの語を含むセルが最も多い行。
 * 一つも無ければ 1 行目とみなす。
 */
export function detectHeaderRow(grid: Grid): number {
  let best = 1;
  let bestScore = 0;
  const n = Math.min(HEADER_SCAN_ROWS, grid.length);
  for (let i = 0; i < n; i++) {
    const s = scoreHeaderRow(grid[i] ?? []);
    if (s > bestScore) {
      best = i + 1;
      bestScore = s;
    }
  }
  return best;
}

/**
 * 見出しのセルから、9 項目それぞれの列を推定する。
 * まず語がそのまま一致する列、無ければ見出しに語を含む列（長い語から優先）。
 * 一度使った列は他の項目に使わない。
 */
export function guessMapping(headerCells: readonly unknown[]): ColumnMapping {
  const headers = headerCells.map(normHeader);
  const used = new Set<number>();
  const mapping: ColumnMapping = {};

  // 1 周目: そのまま一致
  for (const { field, aliases } of NORMALIZED_ALIASES) {
    const idx = headers.findIndex((h, i) => h !== "" && !used.has(i) && aliases.includes(h));
    if (idx >= 0) {
      mapping[field] = idx;
      used.add(idx);
    }
  }
  // 2 周目: 部分一致（長い語から）
  for (const { field, aliases } of NORMALIZED_ALIASES) {
    if (mapping[field] !== undefined) continue;
    const sorted = [...aliases].sort((a, b) => b.length - a.length);
    let found = -1;
    for (const a of sorted) {
      found = headers.findIndex((h, i) => h !== "" && !used.has(i) && h.includes(a));
      if (found >= 0) break;
    }
    if (found >= 0) {
      mapping[field] = found;
      used.add(found);
    }
  }
  return mapping;
}

/** 足りない必須の項目 */
export function missingFields(mapping: ColumnMapping): Field[] {
  return REQUIRED_FIELDS.filter((f) => mapping[f] === undefined);
}

/** 同じ列を 2 つ以上の項目に使っていないか */
export function duplicatedColumns(mapping: ColumnMapping): number[] {
  const seen = new Map<number, number>();
  for (const f of FIELDS) {
    const c = mapping[f];
    if (c === undefined) continue;
    seen.set(c, (seen.get(c) ?? 0) + 1);
  }
  return [...seen.entries()].filter(([, n]) => n > 1).map(([c]) => c);
}

/** 見出し行のセルの表示用の文字（空なら「列 C」のような列名） */
export function headerLabel(grid: Grid, headerRow: number, col: number): string {
  const v = grid[headerRow - 1]?.[col];
  const text = v === null || v === undefined ? "" : String(v).trim();
  return text || `（空）${columnLetter(col)} 列`;
}

export function columnLetter(col: number): string {
  let s = "";
  let n = col + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** シートの列数（見出し行と先頭 50 行の最大） */
export function columnCount(grid: Grid, headerRow: number): number {
  let n = grid[headerRow - 1]?.length ?? 0;
  const end = Math.min(grid.length, headerRow + 50);
  for (let i = headerRow; i < end; i++) n = Math.max(n, grid[i]?.length ?? 0);
  return n;
}

/** ファイルを読んだ直後の設定（見出し行と列を自動推定） */
export function makeFileSetup(fileName: string, grid: Grid): FileSetup {
  const sheet: SheetData = { fileName, office: officeNameFromFile(fileName), grid };
  const headerRow = detectHeaderRow(grid);
  return { sheet, headerRow, mapping: guessMapping(grid[headerRow - 1] ?? []) };
}
