/**
 * 月次売上報告ツール — 判定規則と集計
 * Python 版（monthly_report.py）の load_masters / read_office_file / aggregate / write_report の判定と同じ。
 *
 * 判定の決まり（ヒアリングで決めたこと）:
 * - 得意先はコードでマスタに寄せる。コードが空欄なら、名前を正規化してマスタと照合。寄らなければ要確認・集計外
 * - コードと名前の食い違いは要確認・集計外
 * - 商品分類はマスタから。マスタに無い商品コードは要確認・集計外
 * - 数量×単価≠金額は要確認・集計外
 * - 対象月外の日付は要確認・集計外
 * - 合計行・空行・題名行は読み飛ばす（件数を数える）
 * - 同じ内容の行は要確認に出すが、集計には両方入れる
 * - 返品（負の数）はそのまま集計
 * - マスタ無しのときは、得意先名を正規化した名前で集計し、商品分類は「未分類」
 *
 * Copyright (c) 2026 OFFREL Inc.
 * MIT License（https://opensource.org/licenses/MIT）
 * このコードは動作の保証をしていません。結果はご自身でお確かめください。
 */
import { cellText, formatYmd, normalizeName, parseDate, parseNumber, previousMonth } from "./normalize";
import { missingFields } from "./columns";
import type {
  Aggregates,
  CellValue,
  CheckRow,
  ComparisonRow,
  DetailRow,
  Field,
  FileSetup,
  Grid,
  Masters,
  OfficeResult,
  ReportModel,
  SummaryRow,
} from "./types";

export const NO_STAFF = "（担当者なし）";
export const NO_CATEGORY = "未分類";

function text(v: CellValue | undefined): string | null {
  const s = cellText(v).trim();
  return s || null;
}

/* ---------------------------------------------------------------- マスタ */

/**
 * 得意先マスタ（A: コード, B: 名前）と商品マスタ（A: コード, B: 名前, C: 分類）。1 行目は見出しとして読み飛ばす。
 * どちらも無ければ undefined を渡す。
 */
export function buildMasters(customerGrid?: Grid, productGrid?: Grid): Masters {
  const customers = new Map<string, string>();
  const customersByName = new Map<string, string>();
  const products = new Map<string, { name: string; category: string }>();
  const categories: string[] = [];
  const notes: string[] = [];

  if (customerGrid) {
    for (const row of customerGrid.slice(1)) {
      const code = text(row?.[0]);
      if (!code) continue;
      customers.set(code, cellText(row[1]).trim());
    }
    for (const [code, name] of customers) {
      const key = normalizeName(name);
      const prev = customersByName.get(key);
      if (prev !== undefined) notes.push(`得意先マスタで正規化後の名前が重複「${name}」（${code} と ${prev}）`);
      customersByName.set(key, code);
    }
  }
  if (productGrid) {
    for (const row of productGrid.slice(1)) {
      const code = text(row?.[0]);
      if (!code) continue;
      const category = cellText(row[2]).trim();
      products.set(code, { name: cellText(row[1]).trim(), category });
      if (!categories.includes(category)) categories.push(category);
    }
  }
  return {
    customers,
    customersByName,
    products,
    categories: productGrid ? categories : [NO_CATEGORY],
    hasCustomers: Boolean(customerGrid),
    hasProducts: Boolean(productGrid),
    notes,
  };
}

/* ---------------------------------------------------------------- 営業所ファイル */

function isBlankRow(row: readonly CellValue[]): boolean {
  return row.every((v) => v === null || v === undefined || String(v).trim() === "");
}

function firstText(row: readonly CellValue[]): string {
  for (const v of row) {
    if (v !== null && v !== undefined && String(v).trim()) return String(v);
  }
  return "";
}

/** 1 ファイルを読んで、行ごとの判定をつける */
export function readOffice(setup: FileSetup, year: number, month: number, masters: Masters): OfficeResult {
  const { sheet, headerRow, mapping } = setup;
  const office = sheet.office;
  const file = sheet.fileName;
  const missing = missingFields(mapping);
  if (missing.length) {
    return { office, file, rows: [], skipped: {}, error: `列が足りない: ${missing.join("、")}` };
  }

  const rows: DetailRow[] = [];
  const skipped: Record<string, number> = {};
  const seen = new Map<string, DetailRow[]>();

  const cell = (row: readonly CellValue[], field: Field): CellValue | undefined => {
    const i = mapping[field];
    if (i === undefined || i >= row.length) return undefined;
    const v = row[i];
    return typeof v === "string" ? v.trim() : v;
  };

  for (let r = headerRow; r < sheet.grid.length; r++) {
    const row = sheet.grid[r] ?? [];
    const rowNo = r + 1;
    if (isBlankRow(row)) {
      skipped["空行"] = (skipped["空行"] ?? 0) + 1;
      continue;
    }
    const head = firstText(row);
    if (head.startsWith("合計") || head.startsWith("小計") || head.startsWith("総計")) {
      skipped["合計行"] = (skipped["合計行"] ?? 0) + 1;
      continue;
    }

    const rawDate = cell(row, "日付") ?? null;
    const rec: DetailRow = {
      office,
      file,
      row: rowNo,
      date: parseDate(rawDate, year),
      rawDate,
      custCodeRaw: text(cell(row, "得意先コード")),
      custNameRaw: text(cell(row, "得意先名")),
      custCode: null,
      custName: null,
      productCode: text(cell(row, "商品コード")),
      productNameRaw: cell(row, "商品名") ?? null,
      productName: null,
      category: null,
      qty: parseNumber(cell(row, "数量")),
      unit: parseNumber(cell(row, "単価")),
      amount: parseNumber(cell(row, "金額")),
      staff: text(cell(row, "担当者")) ?? NO_STAFF,
      reasons: [],
      included: true,
    };

    // 日付
    if (!rec.date) {
      rec.reasons.push(`日付が読めない（${JSON.stringify(cellText(rawDate))}）`);
    } else if (rec.date.y !== year || rec.date.m !== month) {
      rec.reasons.push(`対象月外の日付（${formatYmd(rec.date)}）`);
    }

    // 得意先: コード優先。空欄なら名前で寄せる
    const code = rec.custCodeRaw;
    const name = rec.custNameRaw;
    if (masters.hasCustomers) {
      if (code) {
        const masterName = masters.customers.get(code);
        if (masterName !== undefined) {
          rec.custCode = code;
          rec.custName = masterName;
          if (name && normalizeName(name) !== normalizeName(masterName)) {
            rec.reasons.push(`コードと名前が食い違い（${code} はマスタでは「${masterName}」）`);
          }
        } else {
          rec.reasons.push(`得意先コードがマスタにない（${code}）`);
        }
      } else if (name) {
        const hit = masters.customersByName.get(normalizeName(name));
        if (hit !== undefined) {
          rec.custCode = hit;
          rec.custName = masters.customers.get(hit) ?? null;
        } else {
          rec.reasons.push("得意先コード空欄で、名前もマスタに寄らない");
        }
      } else {
        rec.reasons.push("得意先コードも名前も空欄");
      }
    } else if (code || name) {
      // マスタ無し: 正規化した名前で集計（名前が無ければコード）
      rec.custCode = code;
      rec.custName = name ? normalizeName(name) : code;
    } else {
      rec.reasons.push("得意先コードも名前も空欄");
    }

    // 商品
    if (masters.hasProducts) {
      const p = rec.productCode ? masters.products.get(rec.productCode) : undefined;
      if (p) {
        rec.productName = p.name;
        rec.category = p.category;
      } else {
        rec.productName = rec.productNameRaw;
        rec.reasons.push(`商品コードがマスタにない（${rec.productCode ?? "空欄"}）`);
      }
    } else {
      rec.productName = rec.productNameRaw;
      rec.category = NO_CATEGORY;
    }

    // 数値
    for (const f of ["数量", "単価", "金額"] as const) {
      const key = f === "数量" ? "qty" : f === "単価" ? "unit" : "amount";
      if (rec[key] === null) rec.reasons.push(`${f}が読めない（${JSON.stringify(cellText(cell(row, f)))}）`);
    }
    if (rec.qty !== null && rec.unit !== null && rec.amount !== null && rec.qty * rec.unit !== rec.amount) {
      rec.reasons.push(
        `数量×単価≠金額（${rec.qty}×${rec.unit}=${(rec.qty * rec.unit).toLocaleString("en-US")}、金額は ${rec.amount.toLocaleString("en-US")}）`,
      );
    }

    rec.included = rec.reasons.length === 0;
    const key = [
      formatYmd(rec.date) ?? "",
      rec.custCode ?? rec.custNameRaw ?? "",
      rec.productCode ?? "",
      rec.qty ?? "",
      rec.unit ?? "",
      rec.amount ?? "",
      rec.staff,
    ].join("\u0001");
    const list = seen.get(key);
    if (list) list.push(rec);
    else seen.set(key, [rec]);
    rows.push(rec);
  }

  // 同じ内容の行（名寄せ後に全項目が一致）: 要確認に出すが、集計には両方入れる
  for (const recs of seen.values()) {
    if (recs.length > 1) {
      const nums = recs.map((x) => x.row).join(", ");
      for (const rec of recs) {
        rec.reasons.push(
          `同じ内容の行が ${recs.length} 回（行 ${nums}。得意先を寄せた後の日付・得意先・商品・数量・単価・金額・担当者が同じ）`,
        );
      }
    }
  }
  return { office, file, rows, skipped };
}

/* ---------------------------------------------------------------- 集計 */

export function aggregate(results: readonly OfficeResult[]): Aggregates {
  const agg: Aggregates = { customers: new Map(), categories: new Map(), staff: new Map(), offices: new Map() };
  for (const { office, rows } of results) {
    if (!agg.offices.has(office)) agg.offices.set(office, 0);
    for (const r of rows) {
      if (!r.included || r.amount === null) continue;
      const ck = `${r.custCode ?? ""}\t${r.custName ?? ""}`;
      const c = agg.customers.get(ck) ?? { code: r.custCode, name: r.custName, amount: 0 };
      c.amount += r.amount;
      agg.customers.set(ck, c);

      const cat = r.category ?? NO_CATEGORY;
      const g = agg.categories.get(cat) ?? { amount: 0, qty: 0 };
      g.amount += r.amount;
      g.qty += r.qty ?? 0;
      agg.categories.set(cat, g);

      const sk = `${office}\t${r.staff}`;
      const s = agg.staff.get(sk) ?? { office, staff: r.staff, amount: 0 };
      s.amount += r.amount;
      agg.staff.set(sk, s);

      agg.offices.set(office, (agg.offices.get(office) ?? 0) + r.amount);
    }
  }
  return agg;
}

export function ratio(cur: number, prev: number | null): number | null {
  if (prev === null || prev === 0) return null;
  return cur / prev - 1;
}

/** 要確認の理由の先頭の語（カッコの前まで） */
export function reasonHead(reason: string): string {
  return reason.split(/[（(]/)[0];
}

/* ---------------------------------------------------------------- 報告の組み立て */

export interface BuildReportInput {
  ym: string;
  current: OfficeResult[];
  /** 前月のファイルが無ければ null */
  previous: OfficeResult[] | null;
  /** 商品分類の並び（マスタの順） */
  categories: string[];
}

function cmpRow(key: string | null, name: string | null, cur: number, prev: number | null, qty?: number): ComparisonRow {
  return {
    key,
    name,
    ...(qty !== undefined ? { qty } : {}),
    amount: cur,
    prevAmount: prev,
    diff: prev === null ? null : cur - prev,
    ratio: ratio(cur, prev),
  };
}

/** 画面と Excel の両方で使う報告の全体を作る */
export function buildReport(input: BuildReportInput): ReportModel {
  const { ym, current, previous } = input;
  const hasPrev = previous !== null;
  const aggC = aggregate(current);
  const aggP = aggregate(previous ?? []);

  const summary: SummaryRow[] = [];
  let tRows = 0;
  let tInc = 0;
  let tChk = 0;
  for (const o of current) {
    const inc = o.rows.filter((r) => r.included).length;
    const chk = o.rows.filter((r) => r.reasons.length > 0).length;
    const cur = aggC.offices.get(o.office) ?? 0;
    const prev = hasPrev ? (aggP.offices.get(o.office) ?? 0) : null;
    summary.push({
      office: o.office,
      file: o.file,
      rows: o.rows.length,
      included: inc,
      flagged: chk,
      skipped:
        Object.entries(o.skipped)
          .map(([k, v]) => `${k} ${v}`)
          .join("、") || "なし",
      amount: cur,
      prevAmount: prev,
      diff: prev === null ? null : cur - prev,
      ratio: ratio(cur, prev),
    });
    tRows += o.rows.length;
    tInc += inc;
    tChk += chk;
  }
  const tc = [...aggC.offices.values()].reduce((a, b) => a + b, 0);
  const tp = hasPrev ? [...aggP.offices.values()].reduce((a, b) => a + b, 0) : null;
  const total: SummaryRow = {
    office: "合計",
    file: "",
    rows: tRows,
    included: tInc,
    flagged: tChk,
    skipped: "",
    amount: tc,
    prevAmount: tp,
    diff: tp === null ? null : tc - tp,
    ratio: ratio(tc, tp),
  };

  const reasonMap = new Map<string, number>();
  for (const o of current) {
    for (const r of o.rows) {
      for (const why of r.reasons) {
        const h = reasonHead(why);
        reasonMap.set(h, (reasonMap.get(h) ?? 0) + 1);
      }
    }
  }
  const reasonCounts = [...reasonMap.entries()]
    .map(([reason, count]) => ({ reason, count }))
    .sort((a, b) => b.count - a.count);

  // 得意先別（当月の金額の大きい順。前月だけに出てくる得意先も載せる）
  const custKeys = new Set([...aggC.customers.keys(), ...(hasPrev ? aggP.customers.keys() : [])]);
  const customers = [...custKeys]
    .map((k) => {
      const c = aggC.customers.get(k);
      const p = aggP.customers.get(k);
      const info = c ?? p!;
      return cmpRow(info.code, info.name, c?.amount ?? 0, hasPrev ? (p?.amount ?? 0) : null);
    })
    .sort((a, b) => b.amount - a.amount || String(a.key ?? "").localeCompare(String(b.key ?? ""), "ja"));

  // 商品分類別（マスタの順。マスタに無い分類は後ろ）
  const catOrder = [...input.categories];
  for (const k of [...aggC.categories.keys(), ...aggP.categories.keys()]) if (!catOrder.includes(k)) catOrder.push(k);
  const categories = catOrder.map((cat) =>
    cmpRow(cat, null, aggC.categories.get(cat)?.amount ?? 0, hasPrev ? (aggP.categories.get(cat)?.amount ?? 0) : null, aggC.categories.get(cat)?.qty ?? 0),
  );
  const totalQty = [...aggC.categories.values()].reduce((a, b) => a + b.qty, 0);

  // 担当者別（営業所の順 → 当月の金額の大きい順）
  const officeIndex = (name: string) => {
    const i = current.findIndex((o) => o.office === name);
    return i < 0 ? 99 : i;
  };
  const staffKeys = new Set([...aggC.staff.keys(), ...(hasPrev ? aggP.staff.keys() : [])]);
  const staff = [...staffKeys]
    .map((k) => {
      const c = aggC.staff.get(k);
      const p = aggP.staff.get(k);
      const info = c ?? p!;
      return cmpRow(info.office, info.staff, c?.amount ?? 0, hasPrev ? (p?.amount ?? 0) : null);
    })
    .sort((a, b) => officeIndex(a.key ?? "") - officeIndex(b.key ?? "") || b.amount - a.amount);

  // 要確認
  const checks: CheckRow[] = [];
  for (const o of current) {
    for (const r of o.rows) {
      if (!r.reasons.length) continue;
      checks.push({
        office: o.office,
        file: r.file,
        row: r.row,
        reasons: r.reasons.join(" / "),
        date: formatYmd(r.date),
        custCodeRaw: r.custCodeRaw,
        custNameRaw: r.custNameRaw,
        productCode: r.productCode,
        productName: r.productNameRaw,
        qty: r.qty,
        unit: r.unit,
        amount: r.amount,
        calc: r.qty !== null && r.unit !== null ? r.qty * r.unit : null,
        included: r.included,
      });
    }
  }

  const details = current.flatMap((o) => o.rows.filter((r) => r.included));

  return {
    ym,
    prevYm: previousMonth(ym),
    hasPrev,
    summary,
    total,
    reasonCounts,
    customers,
    categories,
    staff,
    checks,
    details,
    totalQty,
  };
}

/** 対象月の各ファイルと（あれば）前月のファイルを読んで、報告を組み立てる */
export function runReport(params: {
  ym: string;
  current: FileSetup[];
  previous: FileSetup[];
  masters: Masters;
}): { report: ReportModel; current: OfficeResult[]; previous: OfficeResult[] | null } {
  const { ym, masters } = params;
  const year = Number(ym.slice(0, 4));
  const month = Number(ym.slice(5, 7));
  const prevYm = previousMonth(ym);
  const py = Number(prevYm.slice(0, 4));
  const pm = Number(prevYm.slice(5, 7));

  const current = params.current.map((s) => readOffice(s, year, month, masters));
  const previous = params.previous.length ? params.previous.map((s) => readOffice(s, py, pm, masters)) : null;
  const report = buildReport({ ym, current, previous, categories: masters.categories });
  return { report, current, previous };
}
