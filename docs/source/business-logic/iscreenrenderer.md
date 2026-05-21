---
title: IScreenRenderer.cls
---

# IScreenRenderer.cls

| 項目 | 内容 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | インターフェースクラス (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | 画面描画の抽象インターフェース（8 メソッド） |
| 行数 | 88 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`IScreenRenderer.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。（このクラスは画面描画クラスが `Implements` で実装するインターフェースです。）詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "IScreenRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' インターフェース: IScreenRenderer v2（ビジネスロジック層）
' 概要:   v2 で simplification、画面構築は modUILoader.ApplyUIToSheet
'         に委譲、本 interface は thin shim 用の 8 メソッドのみ
' 関連:   clsSheetRenderer.cls（v2 実装）, modUILoader.bas（描画本体）
' 関連 ADR: ADR-0053 §1.4 / §6, ADR-0056（mockup → UI スタンザ）
' 関連 schema: v2_ui_stanza_schema.md
'
' v1 → v2 削除済 method:
'   RenderTitle / RenderSubTitle / SetColumnWidths /
'   RenderSection / RenderButton / RenderLabel /
'   RenderInputField / RenderRequiredMark / RenderHint /
'   RenderHeaderRow / RenderEmptyState
'   （全 11 メソッドを modUILoader 内 Apply* に移譲）
' ================================================================
Option Explicit

' ================================================================
' BindSheet
' 概要: 描画対象シートをこの renderer に bind
' 引数: sheetName - ThisWorkbook 内のシート名
' ================================================================
Public Sub BindSheet(ByVal sheetName As String)
End Sub

' ================================================================
' ClearScreen
' 概要: 現 bind sheet の全 cell をクリア（modUILoader が再構築する前提）
' ================================================================
Public Sub ClearScreen()
End Sub

' ================================================================
' ApplyFromStanza
' 概要: <ui_dir>/<xlsmName>/<screenId>.txt を modUILoader 経由で適用
' 引数: xlsmName  - "登録修正" / "検索" / "管理"
'       screenId  - "M-05" 等
' ================================================================
Public Sub ApplyFromStanza(ByVal xlsmName As String, ByVal screenId As String)
End Sub

' ================================================================
' ShowSheet
' 概要: 現 bind sheet を Visible に
' ================================================================
Public Sub ShowSheet()
End Sub

' ================================================================
' HideSheet
' 概要: 現 bind sheet を Hidden に
'       v2 systemSheetVisibility = Hidden 既定（ADR-0053 §7.1 Q5 D2 確定）
' ================================================================
Public Sub HideSheet()
End Sub

' ================================================================
' ActivateSheet
' 概要: 現 bind sheet を active 状態に（シートタブ切替相当）
' ================================================================
Public Sub ActivateSheet()
End Sub

' ================================================================
' ProtectSheet
' 概要: シート保護レベルを適用
' 引数: level - "none" / "light" / "strict"
'       light: パスワード無し、誤入力防止用（書式変更は禁止、内容入力は可）
'       strict: 全保護（DrawingObjects / Contents / Scenarios すべて TRUE）
' ================================================================
Public Sub ProtectSheet(ByVal level As String)
End Sub

' ================================================================
' UnprotectSheet
' 概要: シート保護解除
' ================================================================
Public Sub UnprotectSheet()
End Sub
```
