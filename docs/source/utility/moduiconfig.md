---
title: modUIConfig.bas
---

# modUIConfig.bas

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 3 ブック共通 |
| 役割 | 画面の見た目に関する既定値（フォールバック用） |
| 行数 | 145 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modUIConfig.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modUIConfig"
Option Explicit

' ================================================================
' モジュール: modUIConfig (ユーティリティ層 - UI 設定 SSOT)
' ----------------------------------------------------------------
' 概要:   全画面 (M-01?M-14) の見た目に関わるパラメータを集約。
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

' ----------------------------------------------------------------
' ボタンの大きさ・形
' ----------------------------------------------------------------

' UI_BTN_WIDTH_PT
'   ボタンの最小幅 (ポイント)。
'   セル幅 × UI_BTN_CELL_SPAN がこの値より小さい時は本値が採用される。
'   キャプション "ナレッジ登録" (6 文字) を 10pt フォントで表示するには
'   100pt 以上推奨。短いボタン ("検索") なら 70pt でも収まる。
'   既定: 110
Public Const UI_BTN_WIDTH_PT As Double = 110#

' UI_BTN_HEIGHT_PT
'   ボタンの最小高さ (ポイント)。フォント 10pt + 上下 padding で 26 程度必要。
'   既定: 26
Public Const UI_BTN_HEIGHT_PT As Double = 26#

' UI_BTN_CELL_SPAN
'   ボタンが横方向に占めるセル数。1 だと単独セル幅のみ参照。
'   2 にすると spec の CellAddr とその右隣セルの幅を合算する。
'   メイン画面は B/E/H/K の 3 列間隔配置なので span=2 でも干渉しない。
'   既定: 2
Public Const UI_BTN_CELL_SPAN As Long = 2

' UI_BTN_FONT_SIZE
'   フォームコントロールボタン本体のキャプションフォントサイズ (ポイント)。
'   Excel フォームコントロールはボタン側 Font 制御が効きにくいため、
'   参考値として保持。フォーム描画時に shp.TextFrame.Characters.Font.Size に
'   反映を試みる。
'   既定: 10
Public Const UI_BTN_FONT_SIZE As Double = 10#

' UI_BTN_CAPTION_PREFIX
'   全ボタンキャプションの先頭に挟む文字列。
'   例: "? " (ChrW(&H25B6) + " ") で 三角マーク付き。
'        ""        でマーク無し。
'   旧版は "?" を inline hardcoded していたが、本モジュールで上書き可能。
'   キャプションが長くなりすぎてはみ出す場合は "" にして UI_BTN_WIDTH_PT を
'   そのまま使う運用が安全。
'   既定: "" (マークなし。短い視認性より、はみ出し回避を優先)
Public Const UI_BTN_CAPTION_PREFIX As String = ""

' UI_BTN_ROW_HEIGHT_PT
'   ボタンが配置される行の高さ。UI_BTN_HEIGHT_PT より大きくないと、
'   ボタンが下の行 (ヒント文の行) に食い込んで重なる原因になる。
'   既定: 28 (= UI_BTN_HEIGHT_PT 26 + padding 2)
Public Const UI_BTN_ROW_HEIGHT_PT As Double = 28#

' UI_HINT_ROW_HEIGHT_PT
'   ボタン直下のヒント文を表示する行の高さ。
'   2 行折り返し (例: "新規ナレッジを" & vbLf & "登録") が見えるには 30 以上推奨。
'   既定: 32
Public Const UI_HINT_ROW_HEIGHT_PT As Double = 32#

' ----------------------------------------------------------------
' 文字サイズ
' ----------------------------------------------------------------

' UI_TITLE_FONT_SIZE
'   画面タイトル帯 (A1) のフォントサイズ。
'   既定: 14
Public Const UI_TITLE_FONT_SIZE As Double = 14#

' UI_TITLE_ROW_HEIGHT
'   画面タイトル帯の行高さ。
'   既定: 28
Public Const UI_TITLE_ROW_HEIGHT As Double = 28#

' UI_SECTION_FONT_SIZE
'   セクション帯 (■ ラベル) のフォントサイズ。
'   既定: 11
Public Const UI_SECTION_FONT_SIZE As Double = 11#

' UI_SECTION_ROW_HEIGHT
'   セクション帯の行高さ。
'   既定: 20
Public Const UI_SECTION_ROW_HEIGHT As Double = 20#

' UI_LABEL_FONT_SIZE
'   フィールドラベル (項目名) のフォントサイズ。
'   既定: 10
Public Const UI_LABEL_FONT_SIZE As Double = 10#

' UI_HINT_FONT_SIZE
'   ボタン下のヒント文 / フィールド型表示のフォントサイズ。
'   既定: 9
Public Const UI_HINT_FONT_SIZE As Double = 9#

' UI_HINT_CELL_WRAP
'   ヒント文セルで折り返しを有効にするか。
'   False にすると 1 行表示 (はみ出し時は隣セルに見切れる)。
'   既定: True
Public Const UI_HINT_CELL_WRAP As Boolean = True

' ----------------------------------------------------------------
' タイトル列幅 (A:L) - 既存挙動と同じ
' ----------------------------------------------------------------
' タイトル帯の塗り範囲は A～L 列固定。v2.1 では v1 の RenderTitle メソッドは
' 廃止され、画面タイトル帯の描画は modUILoader 経由の外部 UI スタンザ駆動へ
' 移譲済 (IScreenRenderer は 8 メソッドの薄い shim、ADR-0053 §1.4/§6)。

' ----------------------------------------------------------------
' フィールド領域 vs ボタン領域 境界 (UI-CLEANUP-V1 / Fix 1, 2026-05-16)
' ----------------------------------------------------------------
' 背景:
'   M-03 フォーマット設計 (SPEC_DRIFT SD-05?08 / TC-U-024/025/029/032) で
'   フィールドテーブル (A6 ヘッダ + 7-14 標準 8 フィールド) が
'   下方のアクションボタン領域 (B20/D20/F20/H20) に視覚的に被る事象が
'   YY3/YY4 実機 E2E で報告された。本 Const は spec 構築時の領域境界を
'   SSOT 化し、新規画面追加時の境界違反を防ぐ。

' UI_FIELD_AREA_MAX_COL
'   フィールドテーブル (HeaderLabels + 値) が占有可能な最右列。
'   M-03 標準は 6 列ヘッダ ("並び/フィールド名/型/必須/行数/選択肢/既定値")
'   なので A?F まで使用、G 列はマージン。"G" を超える列にフィールドを
'   描画すると、 ボタン領域 (H 列以降) と被る。
'   既定: "G"
Public Const UI_FIELD_AREA_MAX_COL As String = "G"

' UI_BTN_AREA_ROW_GAP
'   フィールドテーブル最終行とアクションボタン領域開始行の最小余白行数。
'   M-03 では最終フィールド=row 14, [フィールド追加] section=A16, アクション
'   section=A19, ボタン row=20 で、間に 2 行の空白がある (row 15 + 18)。
'   現状は 1 行余白を必須とし、 spec で row 詰めしすぎを防ぐ目安。
'   既定: 1
Public Const UI_BTN_AREA_ROW_GAP As Long = 1
```
