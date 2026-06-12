---
title: IScreenRenderer.cls
description: IScreenRenderer.cls のソースコード（コピペ用）
---

# IScreenRenderer.cls

**配置先**: 共通モジュール（3 ブック共通）
**種類**: クラスモジュール
**更新日**: 2026-06-03 23:22

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `IScreenRenderer.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0370] IScreenRenderer.BindSheet ENTER"  ' [ADR-0100]
End Sub

' ================================================================
' ClearScreen
' 概要: 現 bind sheet の全 cell をクリア（modUILoader が再構築する前提）
' ================================================================
Public Sub ClearScreen()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0371] IScreenRenderer.ClearScreen ENTER"  ' [ADR-0100]
End Sub

' ================================================================
' ApplyFromStanza
' 概要: <ui_dir>/<xlsmName>/<screenId>.txt を modUILoader 経由で適用
' 引数: xlsmName  - "登録修正" / "検索" / "管理"
'       screenId  - "M-05" 等
' ================================================================
Public Sub ApplyFromStanza(ByVal xlsmName As String, ByVal screenId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0372] IScreenRenderer.ApplyFromStanza ENTER"  ' [ADR-0100]
End Sub

' ================================================================
' ShowSheet
' 概要: 現 bind sheet を Visible に
' ================================================================
Public Sub ShowSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0373] IScreenRenderer.ShowSheet ENTER"  ' [ADR-0100]
End Sub

' ================================================================
' HideSheet
' 概要: 現 bind sheet を Hidden に
'       v2 systemSheetVisibility = Hidden 既定（ADR-0053 §7.1 Q5 D2 確定）
' ================================================================
Public Sub HideSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0374] IScreenRenderer.HideSheet ENTER"  ' [ADR-0100]
End Sub

' ================================================================
' ActivateSheet
' 概要: 現 bind sheet を active 状態に（シートタブ切替相当）
' ================================================================
Public Sub ActivateSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0375] IScreenRenderer.ActivateSheet ENTER"  ' [ADR-0100]
End Sub

' ================================================================
' ProtectSheet
' 概要: シート保護レベルを適用
' 引数: level - "none" / "light" / "strict"
'       light: パスワード無し、誤入力防止用（書式変更は禁止、内容入力は可）
'       strict: 全保護（DrawingObjects / Contents / Scenarios すべて TRUE）
' ================================================================
Public Sub ProtectSheet(ByVal level As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0376] IScreenRenderer.ProtectSheet ENTER"  ' [ADR-0100]
End Sub

' ================================================================
' UnprotectSheet
' 概要: シート保護解除
' ================================================================
Public Sub UnprotectSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0377] IScreenRenderer.UnprotectSheet ENTER"  ' [ADR-0100]
End Sub
```
