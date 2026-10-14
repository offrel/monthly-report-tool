import { describe, expect, it } from "vitest";
import { detectHeaderRow, guessMapping, makeFileSetup, missingFields } from "../columns";
import { NO_CATEGORY, buildMasters, buildReport, readOffice, runReport } from "../process";
import type { Grid } from "../types";

const customerGrid: Grid = [
  ["得意先コード", "得意先名"],
  ["C001", "株式会社青葉商店"],
  ["C002", "白樺物産"],
];
const productGrid: Grid = [
  ["商品コード", "商品名", "商品分類", "標準単価"],
  ["P001", "ノート", "文具", 300],
  ["P002", "牛乳", "飲料", 200],
];
const masters = buildMasters(customerGrid, productGrid);

const HEADER = ["日付", "得意先コード", "得意先名", "商品コード", "商品名", "数量", "単価", "金額", "担当者"];

function office(rows: Grid, name = "売上明細_北営業所.xlsx") {
  const setup = makeFileSetup(name, [HEADER, ...rows]);
  return readOffice(setup, 2026, 9, masters);
}

describe("見出し行の検出と列の推定", () => {
  it("題名行があっても見出し行を見つけ、短い見出しも対応づける", () => {
    const grid: Grid = [
      ["北営業所 売上明細 2026年9月", null, null],
      [null, null, null],
      ["担当", "売上日", "得意先CD", "得意先", "商品CD", "品名", "数", "単価", "金額（税抜）", "備考"],
      ["担当A", "2026/9/1", "C001", "青葉商店", "P001", "ノート", 1, 300, 300, null],
    ];
    expect(detectHeaderRow(grid)).toBe(3);
    const m = guessMapping(grid[2]);
    expect(m).toEqual({ 担当者: 0, 日付: 1, 得意先コード: 2, 得意先名: 3, 商品コード: 4, 商品名: 5, 数量: 6, 単価: 7, 金額: 8 });
    expect(missingFields(m)).toEqual([]);
  });
  it("見出しの語を含む列も推定する（部分一致）", () => {
    const m = guessMapping(["売上日付", "お得意先コード", "お得意先名", "商品コード", "商品名称", "数量(個)", "単価(円)", "金額(円)", "営業担当者"]);
    expect(m.日付).toBe(0);
    expect(m.得意先コード).toBe(1);
    expect(m.得意先名).toBe(2);
    expect(m.商品名).toBe(4);
    expect(m.数量).toBe(5);
    expect(m.金額).toBe(7);
    expect(m.担当者).toBe(8);
  });
  it("見出しが見つからなければ 1 行目", () => {
    expect(detectHeaderRow([["a", "b"], ["c", "d"]])).toBe(1);
  });
});

describe("読み飛ばし", () => {
  it("合計行・空行は読み飛ばして件数を数える（要確認には出さない）", () => {
    const r = office([
      ["2026/9/1", "C001", "青葉商店", "P001", "ノート", 2, 300, 600, "担当A"],
      [null, null, null, null, null, null, null, null, null],
      ["", "", "", "", "", "", "", "", ""],
      ["合計", null, null, null, null, null, null, 600, null],
      ["小計", null, null, null, null, null, null, 600, null],
    ]);
    expect(r.rows).toHaveLength(1);
    expect(r.skipped).toEqual({ 空行: 2, 合計行: 2 });
    expect(r.rows[0].included).toBe(true);
  });
});

describe("要確認の各条件", () => {
  it("名寄せ: コードが空欄でも名前の表記ゆれでマスタに寄る", () => {
    const r = office([
      ["2026/9/1", "", "㈱青葉商店", "P001", "ノート", 1, 300, 300, "担当A"],
      ["2026/9/1", null, "青葉商店 株式会社", "P001", "ノート", 1, 300, 300, "担当B"],
    ]);
    expect(r.rows.map((x) => x.custCode)).toEqual(["C001", "C001"]);
    expect(r.rows.map((x) => x.custName)).toEqual(["株式会社青葉商店", "株式会社青葉商店"]);
    expect(r.rows.every((x) => x.included)).toBe(true);
  });
  it("コード空欄で名前も寄らない → 要確認・集計外", () => {
    const r = office([["2026/9/1", "", "新規商店", "P001", "ノート", 1, 300, 300, "担当A"]]);
    expect(r.rows[0].reasons).toEqual(["得意先コード空欄で、名前もマスタに寄らない"]);
    expect(r.rows[0].included).toBe(false);
  });
  it("コードがマスタにない → 要確認・集計外", () => {
    const r = office([["2026/9/1", "C900", "新規商店", "P001", "ノート", 1, 300, 300, "担当A"]]);
    expect(r.rows[0].reasons).toEqual(["得意先コードがマスタにない（C900）"]);
    expect(r.rows[0].included).toBe(false);
  });
  it("コードと名前の食い違い → 要確認・集計外", () => {
    const r = office([["2026/9/1", "C001", "白樺物産", "P001", "ノート", 1, 300, 300, "担当A"]]);
    expect(r.rows[0].reasons).toEqual(["コードと名前が食い違い（C001 はマスタでは「株式会社青葉商店」）"]);
    expect(r.rows[0].included).toBe(false);
  });
  it("コードも名前も空欄 → 要確認", () => {
    const r = office([["2026/9/1", "", "", "P001", "ノート", 1, 300, 300, "担当A"]]);
    expect(r.rows[0].reasons).toEqual(["得意先コードも名前も空欄"]);
  });
  it("商品コードがマスタにない → 要確認・集計外", () => {
    const r = office([["2026/9/1", "C001", "青葉商店", "P999", "謎の品", 1, 300, 300, "担当A"]]);
    expect(r.rows[0].reasons).toEqual(["商品コードがマスタにない（P999）"]);
    expect(r.rows[0].category).toBeNull();
    expect(r.rows[0].included).toBe(false);
  });
  it("数量×単価≠金額 → 要確認・集計外", () => {
    const r = office([["2026/9/1", "C001", "青葉商店", "P001", "ノート", 2, 865, 1710, "担当A"]]);
    expect(r.rows[0].reasons).toEqual(["数量×単価≠金額（2×865=1,730、金額は 1,710）"]);
    expect(r.rows[0].included).toBe(false);
  });
  it("対象月外の日付 → 要確認・集計外", () => {
    const r = office([["2026/8/31", "C001", "青葉商店", "P001", "ノート", 1, 300, 300, "担当A"]]);
    expect(r.rows[0].reasons).toEqual(["対象月外の日付（2026-08-31）"]);
    expect(r.rows[0].included).toBe(false);
  });
  it("日付・数量が読めない → 要確認", () => {
    const r = office([["いつか", "C001", "青葉商店", "P001", "ノート", "たくさん", 300, 300, "担当A"]]);
    expect(r.rows[0].reasons[0]).toMatch(/^日付が読めない/);
    expect(r.rows[0].reasons[1]).toMatch(/^数量が読めない/);
    expect(r.rows[0].included).toBe(false);
  });
  it("同じ内容の行 → 要確認だが集計には両方含む", () => {
    const r = office([
      ["2026/9/1", "C001", "青葉商店", "P001", "ノート", 1, 300, 300, "担当A"],
      ["2026/9/1", "", "㈱青葉商店", "P001", "ノート", 1, 300, 300, "担当A"],
    ]);
    expect(r.rows[0].reasons).toEqual(["同じ内容の行が 2 回（行 2, 3。得意先を寄せた後の日付・得意先・商品・数量・単価・金額・担当者が同じ）"]);
    expect(r.rows.every((x) => x.included)).toBe(true);
  });
  it("返品（負の数）はそのまま集計", () => {
    const r = office([["2026/9/1", "C001", "青葉商店", "P001", "ノート", -2, 300, -600, "担当A"]]);
    expect(r.rows[0].included).toBe(true);
    expect(r.rows[0].amount).toBe(-600);
  });
  it("担当者が空欄なら（担当者なし）", () => {
    const r = office([["2026/9/1", "C001", "青葉商店", "P001", "ノート", 1, 300, 300, null]]);
    expect(r.rows[0].staff).toBe("（担当者なし）");
  });
  it("必須の列が無ければエラー", () => {
    const setup = makeFileSetup("x.xlsx", [["日付", "金額"], ["2026/9/1", 100]]);
    const r = readOffice(setup, 2026, 9, masters);
    expect(r.error).toMatch(/列が足りない/);
    expect(r.rows).toHaveLength(0);
  });
});

describe("マスタ無し", () => {
  it("正規化した名前で集計し、商品分類は未分類", () => {
    const none = buildMasters();
    const setup = makeFileSetup("売上明細_南営業所.xlsx", [
      HEADER,
      ["2026/9/1", "C001", "㈱青葉商店", "P001", "ノート", 1, 300, 300, "担当A"],
      ["2026/9/2", "C001", "青葉商店 株式会社", "P002", "牛乳", 2, 200, 400, "担当A"],
    ]);
    const r = readOffice(setup, 2026, 9, none);
    expect(r.rows.every((x) => x.included)).toBe(true);
    expect(r.rows.map((x) => x.custName)).toEqual(["青葉商店", "青葉商店"]);
    expect(r.rows[0].category).toBe(NO_CATEGORY);
    const report = buildReport({ ym: "2026-09", current: [r], previous: null, categories: none.categories });
    expect(report.customers).toHaveLength(1);
    expect(report.customers[0].amount).toBe(700);
    expect(report.categories.map((c) => c.key)).toEqual([NO_CATEGORY]);
  });
});

describe("集計と前月比", () => {
  it("営業所別・得意先別・分類別・担当者別の合計が一致し、前月比が出る", () => {
    const cur = makeFileSetup("売上明細_北営業所.xlsx", [
      HEADER,
      ["2026/9/1", "C001", "青葉商店", "P001", "ノート", 2, 300, 600, "担当A"],
      ["2026/9/2", "C002", "白樺物産", "P002", "牛乳", 3, 200, 600, "担当B"],
      ["2026/9/3", "C002", "白樺物産", "P002", "牛乳", 1, 200, 999, "担当B"], // 要確認・集計外
    ]);
    const prev = makeFileSetup("売上明細_北営業所.xlsx", [
      HEADER,
      ["2026/8/1", "C001", "青葉商店", "P001", "ノート", 1, 300, 300, "担当A"],
      ["2026/8/2", "C002", "白樺物産", "P002", "牛乳", 1, 200, 200, "担当B"],
      ["2026/8/2", "C002", "白樺物産", "P002", "牛乳", 1, 200, 200, "担当C"],
    ]);
    const { report } = runReport({ ym: "2026-09", current: [cur], previous: [prev], masters });
    expect(report.total.amount).toBe(1200);
    expect(report.total.prevAmount).toBe(700);
    expect(report.total.rows).toBe(3);
    expect(report.total.included).toBe(2);
    expect(report.total.flagged).toBe(1);
    expect(report.summary[0].ratio).toBeCloseTo(1200 / 700 - 1);
    expect(report.customers.map((c) => [c.key, c.amount, c.prevAmount])).toEqual([
      ["C001", 600, 300],
      ["C002", 600, 400],
    ]);
    expect(report.categories.map((c) => [c.key, c.qty, c.amount])).toEqual([
      ["文具", 2, 600],
      ["飲料", 3, 600],
    ]);
    expect(report.staff.map((s) => [s.name, s.amount, s.prevAmount])).toEqual([
      ["担当A", 600, 300],
      ["担当B", 600, 200],
      ["担当C", 0, 200],
    ]);
    expect(report.checks).toHaveLength(1);
    expect(report.reasonCounts).toEqual([{ reason: "数量×単価≠金額", count: 1 }]);
    expect(report.details).toHaveLength(2);
  });
  it("前月が無ければ前月比は空欄", () => {
    const cur = makeFileSetup("売上明細_北営業所.xlsx", [HEADER, ["2026/9/1", "C001", "青葉商店", "P001", "ノート", 2, 300, 600, "担当A"]]);
    const { report } = runReport({ ym: "2026-09", current: [cur], previous: [], masters });
    expect(report.hasPrev).toBe(false);
    expect(report.total.prevAmount).toBeNull();
    expect(report.total.ratio).toBeNull();
  });
});
