import { describe, expect, it } from "vitest";
import { normHeader, normalizeName, officeNameFromFile, parseDate, parseNumber, previousMonth } from "../normalize";

describe("normalizeName（得意先名の名寄せ）", () => {
  it("株式会社・(株)・㈱ を取り、全角半角と空白をそろえる", () => {
    expect(normalizeName("株式会社青葉商店")).toBe("青葉商店");
    expect(normalizeName("(株)青葉商店")).toBe("青葉商店");
    expect(normalizeName("㈱青葉商店")).toBe("青葉商店");
    expect(normalizeName("青葉商店 株式会社")).toBe("青葉商店");
    expect(normalizeName("株式会社　青葉商店")).toBe("青葉商店");
    expect(normalizeName("（株）青葉商店")).toBe("青葉商店");
    expect(normalizeName("㈲青葉商店")).toBe("青葉商店");
    expect(normalizeName("ｱｵﾊﾞ商店０１")).toBe("アオバ商店01");
  });
});

describe("normHeader（見出しの比較）", () => {
  it("全角英字を半角小文字にし、空白を取る", () => {
    expect(normHeader("得意先ＣＤ")).toBe("得意先cd");
    expect(normHeader(" 金額（税抜） ")).toBe("金額(税抜)");
    expect(normHeader(null)).toBe("");
  });
});

describe("parseNumber", () => {
  it("数字・文字の数字・全角・カンマ・円・△ を読む", () => {
    expect(parseNumber(12000)).toBe(12000);
    expect(parseNumber("12,000")).toBe(12000);
    expect(parseNumber("１２０００")).toBe(12000);
    expect(parseNumber("¥1,200円")).toBe(1200);
    expect(parseNumber("△500")).toBe(-500);
    expect(parseNumber(-2)).toBe(-2);
    expect(parseNumber(2.5)).toBe(2); // Python の round と同じ偶数丸め
    expect(parseNumber(3.5)).toBe(4);
  });
  it("読めないものは null", () => {
    expect(parseNumber("")).toBeNull();
    expect(parseNumber(null)).toBeNull();
    expect(parseNumber("abc")).toBeNull();
    expect(parseNumber(true)).toBeNull();
    expect(parseNumber("1,2,3円x")).toBeNull();
  });
});

describe("parseDate（5 書式＋日付型）", () => {
  it("2026/9/1", () => expect(parseDate("2026/9/1", 2026)).toEqual({ y: 2026, m: 9, d: 1 }));
  it("2026-09-01", () => expect(parseDate("2026-09-01", 2026)).toEqual({ y: 2026, m: 9, d: 1 }));
  it("9月1日 は対象月の年", () => expect(parseDate("9月1日", 2026)).toEqual({ y: 2026, m: 9, d: 1 }));
  it("9/1 は対象月の年", () => expect(parseDate("9/1", 2026)).toEqual({ y: 2026, m: 9, d: 1 }));
  it("20260901（数字）", () => expect(parseDate(20260901, 2026)).toEqual({ y: 2026, m: 9, d: 1 }));
  it("20260901（文字）", () => expect(parseDate("20260901", 2026)).toEqual({ y: 2026, m: 9, d: 1 }));
  it("2026年9月1日", () => expect(parseDate("2026年9月1日", 2026)).toEqual({ y: 2026, m: 9, d: 1 }));
  it("日付型（UTC 0 時）", () => expect(parseDate(new Date(Date.UTC(2026, 8, 1)), 2026)).toEqual({ y: 2026, m: 9, d: 1 }));
  it("Excel のシリアル値", () => expect(parseDate(46266, 2026)).toEqual({ y: 2026, m: 9, d: 1 }));
  it("読めないもの・存在しない日は null", () => {
    expect(parseDate("昨日", 2026)).toBeNull();
    expect(parseDate("2026/13/1", 2026)).toBeNull();
    expect(parseDate("2026/2/30", 2026)).toBeNull();
    expect(parseDate("", 2026)).toBeNull();
    expect(parseDate(null, 2026)).toBeNull();
    expect(parseDate(123, 2026)).toBeNull();
  });
});

describe("previousMonth / officeNameFromFile", () => {
  it("前月", () => {
    expect(previousMonth("2026-09")).toBe("2026-08");
    expect(previousMonth("2026-01")).toBe("2025-12");
  });
  it("ファイル名から営業所名", () => {
    expect(officeNameFromFile("売上明細_北営業所.xlsx")).toBe("北営業所");
    expect(officeNameFromFile("売上明細＿南営業所.xlsx")).toBe("南営業所");
    expect(officeNameFromFile("2026-08_売上明細_北営業所.xlsx")).toBe("北営業所");
    expect(officeNameFromFile("9月分.xlsx")).toBe("9月分");
    expect(officeNameFromFile("売上明細.xlsx")).toBe("売上明細");
  });
});
