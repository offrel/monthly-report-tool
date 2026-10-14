/**
 * 同梱のサンプル（samples/）で、全体を通した結果を確かめる。
 * 期待値は Python 版（monthly_report.py）の出力と突き合わせた値。
 */
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { makeFileSetup } from "../columns";
import { buildMasters, runReport } from "../process";
import { readGrid, writeReportWorkbook } from "../workbook";

const DIR = fileURLToPath(new URL("../../samples/", import.meta.url));
const OFFICES = ["北営業所", "中央営業所", "南営業所"];

function toArrayBuffer(path: string): ArrayBuffer {
  const b = readFileSync(path);
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
}

async function setups(prefix: string) {
  const out = [];
  for (const o of OFFICES) {
    const name = `${prefix}売上明細_${o}.xlsx`;
    out.push(makeFileSetup(name, await readGrid(toArrayBuffer(join(DIR, name)))));
  }
  return out;
}

describe("サンプルデータ（2026-09・前月 2026-08・マスタあり）", () => {
  it("総合計・行数・要確認・分類別が期待値と一致し、xlsx を書き出せる", async () => {
    const masters = buildMasters(
      await readGrid(toArrayBuffer(join(DIR, "得意先マスタ.xlsx"))),
      await readGrid(toArrayBuffer(join(DIR, "商品マスタ.xlsx"))),
    );
    const { report } = runReport({ ym: "2026-09", current: await setups(""), previous: await setups("2026-08_"), masters });

    expect(report.summary.map((s) => s.office)).toEqual(OFFICES);
    expect(report.summary.map((s) => [s.rows, s.included, s.flagged, s.amount, s.prevAmount])).toEqual([
      [1802, 1794, 14, 19443624, 18353207],
      [2400, 2387, 19, 25149129, 24313634],
      [1600, 1592, 10, 17953004, 15791873],
    ]);
    expect(report.summary[2].skipped).toBe("空行 1、合計行 1");
    expect([report.total.rows, report.total.included, report.total.flagged, report.total.amount, report.total.prevAmount]).toEqual([
      5802, 5773, 43, 62545757, 58458714,
    ]);
    expect(report.categories.map((c) => [c.key, c.qty, c.amount, c.prevAmount])).toEqual([
      ["食品", 20155, 15539994, 12771244],
      ["飲料", 16284, 10921204, 9571899],
      ["日用品", 16392, 16065107, 16490347],
      ["文具", 11761, 10483503, 11001075],
      ["雑貨", 13161, 9535949, 8624149],
    ]);
    expect(report.customers[0]).toMatchObject({ key: "C011", amount: 2318341, prevAmount: 1471476 });
    expect(report.checks).toHaveLength(43);
    expect(report.reasonCounts).toEqual([
      { reason: "同じ内容の行が 2 回", count: 14 },
      { reason: "得意先コード空欄で、名前もマスタに寄らない", count: 9 },
      { reason: "数量×単価≠金額", count: 8 },
      { reason: "得意先コードがマスタにない", count: 4 },
      { reason: "コードと名前が食い違い", count: 3 },
      { reason: "対象月外の日付", count: 3 },
      { reason: "商品コードがマスタにない", count: 2 },
    ]);

    const buf = await writeReportWorkbook(report);
    expect(buf.byteLength).toBeGreaterThan(100_000);
  });
});
