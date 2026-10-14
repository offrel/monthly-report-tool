/**
 * 月次売上報告ツール — 文字・数値・日付の読み方
 * Python 版（monthly_report.py）の norm_header / normalize_name / parse_number / parse_date と同じ判定。
 *
 * Copyright (c) 2026 OFFREL Inc.
 * MIT License（https://opensource.org/licenses/MIT）
 * このコードは動作の保証をしていません。結果はご自身でお確かめください。
 */
import type { CellValue, YMD } from "./types";

/** 得意先名から取り除く法人格。NFKC で ㈱→(株)、㈲→(有) になるので、カッコ書きも含める */
export const CORP_WORDS = ["株式会社", "有限会社", "合同会社", "(株)", "(有)", "(同)"];

/** 見出しの比較用。全角半角をそろえ、小文字にし、空白を取る */
export function normHeader(s: unknown): string {
  return String(s ?? "")
    .normalize("NFKC")
    .trim()
    .toLowerCase()
    .replace(/[ 　]/g, "");
}

/** 得意先名の名寄せ用。全角半角をそろえ、空白と法人格を取る */
export function normalizeName(s: unknown): string {
  let t = String(s ?? "").normalize("NFKC"); // ㈱→(株)、半角カナ→全角、全角英数→半角
  t = t.replace(/\s+/g, "");
  for (const w of CORP_WORDS) t = t.split(w).join("");
  return t;
}

/** Python の round()（偶数丸め）と同じにする */
function roundHalfEven(x: number): number {
  const f = Math.floor(x);
  const diff = x - f;
  if (diff > 0.5) return f + 1;
  if (diff < 0.5) return f;
  return f % 2 === 0 ? f : f + 1;
}

const NUMBER_RE = /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i;

/** 数値のセルを整数にする。'12,000' や全角数字も読む。読めなければ null */
export function parseNumber(v: CellValue | undefined): number | null {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "boolean") return null;
  if (typeof v === "number") return Number.isFinite(v) ? roundHalfEven(v) : null;
  if (v instanceof Date) return null;
  let s = String(v).normalize("NFKC").trim();
  s = s.replace(/[,¥円 ]/g, "");
  if (s.startsWith("△") || s.startsWith("▲")) s = "-" + s.slice(1);
  if (!NUMBER_RE.test(s)) return null;
  const n = Number(s);
  return Number.isFinite(n) ? roundHalfEven(n) : null;
}

function validYmd(y: number, m: number, d: number): YMD | null {
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  const t = new Date(Date.UTC(y, m - 1, d));
  if (t.getUTCFullYear() !== y || t.getUTCMonth() !== m - 1 || t.getUTCDate() !== d) return null;
  return { y, m, d };
}

/**
 * 日付のセルを年月日にする。
 * 読める書き方: 日付型 / 2026/9/1 / 2026-09-01 / 2026.9.1 / 2026年9月1日 / 9月1日・9/1（対象月の年とみなす）/ 20260901 / Excel のシリアル値
 */
export function parseDate(v: CellValue | undefined, year: number): YMD | null {
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date) {
    // Excel から読んだ日付は UTC の 0 時で入ってくる
    if (Number.isNaN(v.getTime())) return null;
    return { y: v.getUTCFullYear(), m: v.getUTCMonth() + 1, d: v.getUTCDate() };
  }
  if (typeof v === "boolean") return null;
  if (typeof v === "number") {
    const n = Math.trunc(v);
    if (n >= 19000101 && n <= 21001231) {
      return validYmd(Math.floor(n / 10000), Math.floor(n / 100) % 100, n % 100);
    }
    if (n >= 20000 && n <= 80000) {
      // Excel のシリアル値（1899-12-30 からの日数）
      const t = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
      return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
    }
    return null;
  }
  const s = String(v).normalize("NFKC").trim();
  let m = s.match(/^(\d{4})[/\-.年](\d{1,2})[/\-.月](\d{1,2})日?$/);
  if (m) return validYmd(Number(m[1]), Number(m[2]), Number(m[3]));
  m = s.match(/^(\d{1,2})[/月](\d{1,2})日?$/);
  if (m) return validYmd(year, Number(m[1]), Number(m[2]));
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return validYmd(Number(m[1]), Number(m[2]), Number(m[3]));
  return null;
}

export function formatYmd(d: YMD | null): string | null {
  if (!d) return null;
  return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
}

/** Excel に書くための Date（UTC の 0 時） */
export function ymdToDate(d: YMD): Date {
  return new Date(Date.UTC(d.y, d.m - 1, d.d));
}

/** セルの文字表現（空なら ""） */
export function cellText(v: CellValue | undefined): string {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return formatYmd(parseDate(v, 0)) ?? "";
  return String(v);
}

/** 対象月 "2026-09" → 前月 "2026-08" */
export function previousMonth(ym: string): string {
  const y = Number(ym.slice(0, 4));
  const m = Number(ym.slice(5, 7));
  return m === 1 ? `${y - 1}-12` : `${y}-${String(m - 1).padStart(2, "0")}`;
}

export function isValidYm(ym: string): boolean {
  return /^\d{4}-(0[1-9]|1[0-2])$/.test(ym);
}

/**
 * ファイル名から営業所名を取る。「売上明細_」の後ろ。無ければ拡張子を除いたファイル名。
 * 先頭の「2026-08_」のような年月の接頭辞は外す（前月分のファイルを同じ場所に置けるように）
 */
export function officeNameFromFile(fileName: string): string {
  const base = fileName.replace(/\.[^.]+$/, "");
  const stripped = base.replace(/^\d{4}-?\d{2}[_＿]/, "").replace(/^売上明細[_＿]?/, "");
  return stripped || base;
}
