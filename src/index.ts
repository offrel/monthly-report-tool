/**
 * 月次売上報告ツール — ロジックの入口
 * 営業所ごとの売上明細 Excel を読み、名寄せ・要確認の判定・集計をして、6 シートの xlsx にする。
 * UI（React）には依存しない。ブラウザでも Node でも動く。
 *
 * Copyright (c) 2026 OFFREL Inc.
 * MIT License（https://opensource.org/licenses/MIT）
 * このコードは動作の保証をしていません。結果はご自身でお確かめください。
 */
export * from "./types";
export * from "./normalize";
export * from "./columns";
export * from "./process";
export * from "./workbook";
