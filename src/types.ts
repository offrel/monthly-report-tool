/**
 * 月次売上報告ツール — 型の定義
 *
 * Copyright (c) 2026 OFFREL Inc.
 * MIT License（https://opensource.org/licenses/MIT）
 * このコードは動作の保証をしていません。結果はご自身でお確かめください。
 */

/** 読み取る 9 項目（内部の項目名）。Python 版 monthly_report.py の COLUMN_ALIASES と同じ並び */
export const FIELDS = [
  "日付",
  "得意先コード",
  "得意先名",
  "商品コード",
  "商品名",
  "数量",
  "単価",
  "金額",
  "担当者",
] as const;
export type Field = (typeof FIELDS)[number];

/** 必須の項目（商品名だけは無くてもよい。Python 版の REQUIRED と同じ） */
export const REQUIRED_FIELDS: readonly Field[] = [
  "日付",
  "得意先コード",
  "得意先名",
  "商品コード",
  "数量",
  "単価",
  "金額",
  "担当者",
];

/** セルの値（Excel から読んだそのまま。数式は計算結果、リッチテキストは文字に落とす） */
export type CellValue = string | number | boolean | Date | null;

/** シート 1 枚の中身。grid[r][c] は r+1 行目・c+1 列目のセル */
export type Grid = CellValue[][];

export interface SheetData {
  /** 元のファイル名 */
  fileName: string;
  /** 営業所名（ファイル名の「売上明細_」以降。無ければファイル名） */
  office: string;
  grid: Grid;
}

/** 項目名 → 列番号（0 始まり） */
export type ColumnMapping = Partial<Record<Field, number>>;

/** 1 ファイルぶんの読み取り設定（見出し行と列の対応） */
export interface FileSetup {
  sheet: SheetData;
  /** 見出し行（1 始まり） */
  headerRow: number;
  mapping: ColumnMapping;
}

/** 年月日（タイムゾーンの影響を受けないよう、Date ではなく数字で持つ） */
export interface YMD {
  y: number;
  m: number;
  d: number;
}

export interface ProductInfo {
  name: string;
  category: string;
}

export interface Masters {
  /** 得意先コード → 正式名称 */
  customers: Map<string, string>;
  /** 正規化した名前 → 得意先コード */
  customersByName: Map<string, string>;
  /** 商品コード → 商品名・分類 */
  products: Map<string, ProductInfo>;
  /** 商品分類（マスタに出てきた順） */
  categories: string[];
  hasCustomers: boolean;
  hasProducts: boolean;
  /** 読み込み時の注意（正規化後の名前が重複など） */
  notes: string[];
}

/** 明細 1 行の読み取り結果 */
export interface DetailRow {
  office: string;
  file: string;
  /** 元ファイルの行番号（1 始まり） */
  row: number;
  date: YMD | null;
  rawDate: CellValue;
  custCodeRaw: string | null;
  custNameRaw: string | null;
  /** 名寄せ後の得意先コード（マスタに寄らなければ null） */
  custCode: string | null;
  /** 名寄せ後の得意先名（マスタの正式名称。マスタ無しなら正規化した名前） */
  custName: string | null;
  productCode: string | null;
  productNameRaw: CellValue;
  productName: CellValue;
  category: string | null;
  qty: number | null;
  unit: number | null;
  amount: number | null;
  staff: string;
  /** 要確認の理由（空なら問題なし） */
  reasons: string[];
  /** 集計に含めるか */
  included: boolean;
}

export interface OfficeResult {
  office: string;
  file: string;
  rows: DetailRow[];
  /** 読み飛ばした行の種類と件数（空行・合計行） */
  skipped: Record<string, number>;
  /** 見出しが見つからない・列が足りない などのエラー */
  error?: string;
}

export interface Aggregates {
  /** key: `${code}\t${name}` */
  customers: Map<string, { code: string | null; name: string | null; amount: number }>;
  categories: Map<string, { amount: number; qty: number }>;
  /** key: `${office}\t${staff}` */
  staff: Map<string, { office: string; staff: string; amount: number }>;
  offices: Map<string, number>;
}

export interface SummaryRow {
  office: string;
  file: string;
  rows: number;
  included: number;
  flagged: number;
  skipped: string;
  amount: number;
  prevAmount: number | null;
  diff: number | null;
  ratio: number | null;
}

export interface ComparisonRow {
  /** 得意先コード／商品分類／営業所 */
  key: string | null;
  /** 得意先名／担当者 */
  name: string | null;
  qty?: number;
  amount: number;
  prevAmount: number | null;
  diff: number | null;
  ratio: number | null;
}

export interface CheckRow {
  office: string;
  file: string;
  row: number;
  reasons: string;
  date: string | null;
  custCodeRaw: string | null;
  custNameRaw: string | null;
  productCode: string | null;
  productName: CellValue;
  qty: number | null;
  unit: number | null;
  amount: number | null;
  calc: number | null;
  included: boolean;
}

/** 画面表示と Excel 書き出しの両方で使う、報告の全体 */
export interface ReportModel {
  ym: string;
  prevYm: string;
  hasPrev: boolean;
  summary: SummaryRow[];
  total: SummaryRow;
  reasonCounts: Array<{ reason: string; count: number }>;
  customers: ComparisonRow[];
  categories: ComparisonRow[];
  staff: ComparisonRow[];
  checks: CheckRow[];
  details: DetailRow[];
  totalQty: number;
}
