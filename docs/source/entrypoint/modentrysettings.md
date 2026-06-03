---
title: modEntrySettings.bas
---

# modEntrySettings.bas

| 項目 | 内容 |
|---|---|
| 層 | エントリポイント層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 管理.xlsm |
| 役割 | 格納先設定（M-10）・設定（M-11）画面の入口。設定値のシート反映・再読込・破棄 |
| 行数 | 384 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modEntrySettings.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modEntrySettings"
' ============================================================
' modEntrySettings (Entry layer, 管理.xlsm only / v2.1)
' Role:
'   管理.xlsm の設定系 2 画面に対する entry point Sub 群。
'   設計 SSOT (architecture §6 / §7.2 #47-48 / screen_design_v2 §2.9 /
'   placement_v2 §6) が規定する 2 画面分割 (SSOT-Q22 確定):
'     M-10 格納先設定 (clsStorageScreen)        = config.txt の
'         パス系 4 dir (data_dir / format_dir / ui_dir / backup_dir) を編集
'     M-11 設定       (clsSystemSettingsScreen)   = config.txt の
'         debugLevel を編集
'   本 module は M-10 / M-11 両画面の entry (placement_v2 §6)。
'   modConfigHolder (in-memory) の値をシートに反映 (Open) / シート編集値を
'   再ロード (Apply) / 編集破棄 (Discard) する。
'   config.txt 本体への永続化は modConfigLoader 経由の責務 (本 module は
'   memory holder に対する 反映 only、external I/O 直接呼び出しは禁止
'   = ADR-0053 §2.9 I/O 独立原則)。
'   modConfigHolder.SetConfig は dict 全置換 (merge しない) のため、
'   各 