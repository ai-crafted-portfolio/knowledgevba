---
title: modFormatColumns.bas
---

# modFormatColumns.bas

| 項目 | 値 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | フォーマット一覧シート列番号定数の Public 公開 |
| 行数 | 32 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modFormatColumns` に変更してから、コードペインに貼り付けます。

> v12 で新規追加。v11 までは `clsFormatManager.cls` 内で `Public Const FL_COL_*` を定義していたが、VBA 仕様上クラスモジュールでは `Public Const` を使えず、本物の Excel が compile-time に reject する。標準モジュールへ切り出して合法化した。値は v11 と同一で後方互換。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modFormatColumns"
Option Explicit

' ================================================================
' モジュール: modFormatColumns (release_v2 / v12 新規追加)
' 概要:       フォーマット一覧シートの列番号定数を Public 公開する。
'             v11 までは clsFormatManager.cls 内で Public Const 定義していたが、
'             VBA 仕様上「クラスモジュールで Public Const は不可」のため
'             real Excel が compile-time に reject。
'             本モジュール (.bas = 標準モジュール) に移動して合法化。
' 依存先:     なし
' 備考:       clsFormatManager 内部、modDemoSeeder 外部の両方から参照可能。
'             値は v11 までと同一で互換性維持。
' ================================================================

' --- フォーマット一覧シート 列番号 ---
Public Const FL_COL_NO As Long = 1
Public Const FL_COL_FMT_ID As Long = 2
Public Const FL_COL_FMT_NAME As Long = 3
Public Const FL_COL_ID_PATTERN As Long = 4
Public Const FL_COL_CURRENT_NUM As Long = 5
Public Const FL_COL_NEXT_NUM As Long = 6
Public Const FL_COL_VERSION As Long = 7
Public Const FL_COL_FIELD_COUNT As Long = 8
Public Const FL_COL_KNW_COUNT As Long = 9
Public Const FL_COL_CREATED As Long = 10
Public Const FL_COL_UPDATED As Long = 11

' --- フォーマット一覧シート データ開始行 (1=ヘッダー領域) ---
Public Const FL_START_ROW As Long = 3
```

## 関連

- 呼び出される: `clsFormatManager`, `modDemoSeeder`
- 呼び出す: なし
- 関連コミット: v12 (Public Const 切出し)
