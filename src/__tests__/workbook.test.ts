import { describe, expect, it } from "vitest";
import { makeFileSetup } from "../columns";
import { buildMasters, runReport } from "../process";
import { readGrid, toCellValue, writeReportWorkbook } from "../workbook";

const HEADER = ["日付", "得意先コード", "得意先名", "商品コード", "商品名", "数量", "単価", "金額", "担当者"];

describe("Excel の読み書き", () => {
  it("6 シートの xlsx を書き、読み戻せる", async () => {
    const masters = buildMasters(
      [
        ["得意先コード", "得意先名"],
        ["C001", "株式会社青葉商店"],
      ],
      [
        ["商品コード", "商品名", "商品分類"],
        ["P001", "ノート", "文具"],
      ],
    );
    const cur = makeFileSetup("売上明細_北営業所.xlsx", [
      HEADER,
      ["2026/9/1", "C001", "青葉商店", "P001", "ノート", 2, 300, 600, "担当A"],
      ["2026/9/2", "C001", "青葉商店", "P001", "ノート", 1, 300, 999, "担当A"],
    ]);
    const { report } = runReport({ ym: "2026-09", current: [cur], previous: [], masters });
    const buf = await writeReportWorkbook(report);
    expect(buf.byteLength).toBeGreaterThan(1000);

    const ExcelJS = await import("exceljs");
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    expect(wb.worksheets.map((w) => w.name)).toEqual(["サマリー", "得意先別", "商品分類別", "担当者別", "要確認", "統合明細"]);
    const summary = wb.getWorksheet("サマリー")!;
    expect(summary.getRow(5).getCell(1).value).toBe("北営業所");
    expect(summary.getRow(5).getCell(7).value).toBe(600);
    expect(summary.getRow(6).getCell(1).value).toBe("合計");
    expect(summary.getRow(6).getCell(7).value).toBe(600);
    expect(wb.getWorksheet("要確認")!.rowCount).toBe(2);
    expect(wb.getWorksheet("統合明細")!.rowCount).toBe(2);

    // 読み戻し（先頭シート = サマリー）
    const grid = await readGrid(buf);
    expect(grid[3][0]).toBe("営業所");
    expect(grid[4][6]).toBe(600);
  });

  it("日付型・数式・リッチテキストのセルを素の値にする", async () => {
    const ExcelJS = await import("exceljs");
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("s");
    ws.addRow(["日付", "金額", "名前"]);
    const r = ws.addRow([new Date(Date.UTC(2026, 8, 1)), null, null]);
    r.getCell(2).value = { formula: "1+1", result: 2 };
    r.getCell(3).value = { richText: [{ text: "青葉" }, { text: "商店" }] };
    const buf = (await wb.xlsx.writeBuffer()) as unknown as ArrayBuffer;
    const grid = await readGrid(buf);
    expect(grid[1][0]).toBeInstanceOf(Date);
    expect(grid[1][1]).toBe(2);
    expect(grid[1][2]).toBe("青葉商店");
    expect(toCellValue({ text: "x", hyperlink: "http://example.com" })).toBe("x");
  });
});
