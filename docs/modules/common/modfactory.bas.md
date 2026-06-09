---
title: modFactory.bas
description: modFactory.bas のソースコード（コピペ用）
---

# modFactory.bas

**配置先**: 共通モジュール（3 ブック共通）
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modFactory.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modFactory"
Option Explicit

' ================================================================
' modFactory (v2.3 Phase R-1-j stub, 2026-05-28)
' Original CreateScreen path is legacy + references modules that
' no longer exist in current package (modScreenSpecRegistry,
' clsFormatListScreen, clsMainScreen, etc.) causing project-wide
' compile errors. Stubbed to keep API surface (CreateRenderer +
' CreateScreen) without dead references.
' ================================================================

Public Const RENDERER_KIND_SHEET As String = "sheet"
Public Const RENDERER_KIND_FORM As String = "form"

Public Function CreateRenderer(ByVal kind As String) As IScreenRenderer
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1167] modFactory.CreateRenderer ENTER"  ' [ADR-0100]
    Select Case LCase(kind)
        Case RENDERER_KIND_SHEET
            Set CreateRenderer = New clsSheetRenderer
        Case RENDERER_KIND_FORM
            Set CreateRenderer = New clsUserFormRenderer
        Case Else
            Set CreateRenderer = New clsSheetRenderer
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1168] modFactory.CreateRenderer EXIT-OK"  ' [ADR-0100]
End Function

' Phase R-1-j: stubbed. Returns Nothing for all screenIds.
Public Function CreateScreen(ByVal screenId As String, _
                              ByVal renderer As IScreenRenderer) As Object
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1169] modFactory.CreateScreen ENTER"  ' [ADR-0100]
    Set CreateScreen = Nothing
End Function
```
