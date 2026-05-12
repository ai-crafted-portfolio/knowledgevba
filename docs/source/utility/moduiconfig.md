---
title: modUIConfig.bas
---

# modUIConfig.bas

| 項目 | 値 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | 全画面の外観 (ボタンサイズ・フォント・行高) を集中管理する定数モジュール |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modUIConfig` に変更してから、コードペインに貼り付けます。

カスタマイズの手順は [外観のカスタマイズ](../../customize.md) を参照してください。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modUIConfig"
Option Explicit

' ================================================================
' モジュール: modUIConfig (ユーティリティ層 - UI 設定 SSOT)
' ----------------------------------------------------------------
' 概要:   全画面 (M-01〜M-14) の見た目に関わるパラメータを集約。
'         clsSheetRenderer / modScreenRender / modScreenSpecRegistry は
'         本モジュールの Public Const を参照することで、ハードコードを
'         一掃する。
'
' カスタマイズしたい場合:
'   本モジュールの値を書き換えて Excel を再起動するだけで、
'   全画面が新しい設定で再描画される。シート毎に手作業する必要なし。
'   各 Const のコメントに「何を制御するか」「目安値」を記載。
'
' 依存先: なし (純粋な定数モジュール)
' ================================================================

' --- ボタンの大きさ・形 ---
Public Const UI_BTN_WIDTH_PT  As Double  = 110#
Public Const UI_BTN_HEIGHT_PT As Double  = 26#
Public Const UI_BTN_CELL_SPAN As Long    = 2
Public Const UI_BTN_FONT_SIZE As Double  = 10#
Public Const UI_BTN_CAPTION_PREFIX As String = ""
Public Const UI_BTN_ROW_HEIGHT_PT  As Double = 28#
Public Const UI_HINT_ROW_HEIGHT_PT As Double = 32#

' --- 文字サイズ ---
Public Const UI_TITLE_FONT_SIZE   As Double = 14#
Public Const UI_TITLE_ROW_HEIGHT  As Double = 28#
Public Const UI_SECTION_FONT_SIZE As Double = 11#
Public Const UI_SECTION_ROW_HEIGHT As Double = 20#
Public Const UI_LABEL_FONT_SIZE   As Double = 10#
Public Const UI_HINT_FONT_SIZE    As Double = 9#
Public Const UI_HINT_CELL_WRAP    As Boolean = True
```

完全版・コメント付きは [外観のカスタマイズ](../../customize.md) でも閲覧できます。
