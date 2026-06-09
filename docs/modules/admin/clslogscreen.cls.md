---
title: clsLogScreen.cls
description: clsLogScreen.cls のソースコード（コピペ用）
---

# clsLogScreen.cls

**配置先**: `管理.xlsm` 用の VBA モジュール
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\admin\`
- ファイル名: `clsLogScreen.cls`
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
END
Attribute VB_Name = "clsLogScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsLogScreen (画面層 - M-14 操作ログ)
' 概要:   操作ログ表示画面 (M-14 frame)。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   v2.1 (2026-05-20) で publish 投入。buildtest archive 旧版 (M-14)
'         を踏襲しつつ logging 規約 v1 準拠の新規 emit を追加。
'         既存 LOG-M14-SCREENCLS-SETUP-* 系は §3.3 後方互換維持で残置。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0068] clsLogScreen.Init ENTER"  ' [ADR-0100]
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0069] clsLogScreen.Setup ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsLogScreen", "Setup", "ENTER sid=" & m_spec.ScreenId, "LOG-M14-SCREENCLS-SETUP-ENTRY")

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsLogScreen", "Setup", "EXIT ok", "LOG-M14-SCREENCLS-SETUP-EXIT-OK")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0070] clsLogScreen.Setup EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0071] clsLogScreen.Setup EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Call modScreenRender.LogScreenError("clsLogScreen", "Setup", stepName, Err.Number, Err.Description, "LOG-M14-SCREENCLS-SETUP-ERR")
End Sub

Public Sub Render()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0072] clsLogScreen.Render ENTER"  ' [ADR-0100]
    Setup
End Sub

Public Function ValidateInput() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0073] clsLogScreen.ValidateInput ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' Screen-level input precondition check (LOG sheet view).
    ' The LOG screen is read-only and defines no Required input fields, so
    ' a correctly initialised spec validates True; an uninitialised screen
    ' (m_spec Is Nothing) fails the precondition.
    If m_spec Is Nothing Then
        If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0078] clsLogScreen.ValidateInput STEP-1 pre modScreenRender.LogScreenTrace"  ' [ADR-0100]
        Call modScreenRender.LogScreenTrace("clsLogScreen", "ValidateInput", "spec not initialised", "VALIDATE-WARN-WW-250")
        ValidateInput = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0074] clsLogScreen.ValidateInput EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim fld As clsFieldSpec
    For Each fld In m_spec.Fields
        If fld.Required And Len(fld.InputAddr) > 0 Then
            If Len(Trim(GetInputValue(fld.InputAddr))) = 0 Then
                If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0079] clsLogScreen.ValidateInput STEP-2 pre modScreenRender.LogScreenTrace"  ' [ADR-0100]
                Call modScreenRender.LogScreenTrace("clsLogScreen", "ValidateInput", "required field empty: " & fld.Label, "VALIDATE-WARN-WW-250")
                ValidateInput = False
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0075] clsLogScreen.ValidateInput EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
        End If
    Next fld

    ValidateInput = True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0076] clsLogScreen.ValidateInput EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0080] clsLogScreen.ValidateInput STEP-3 pre modScreenRender.LogScreenError"  ' [ADR-0100]
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0077] clsLogScreen.ValidateInput EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Call modScreenRender.LogScreenError("clsLogScreen", "ValidateInput", "validate", Err.Number, Err.Description, "VALIDATE-WARN-WW-250")
    ValidateInput = False
End Function

Private Function GetInputValue(ByVal cellAddr As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0081] clsLogScreen.GetInputValue ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(m_spec.SheetName)
    If ws Is Nothing Then
        GetInputValue = ""
    Else
        GetInputValue = CStr(ws.Range(cellAddr).Value)
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0082] clsLogScreen.GetInputValue EXIT-OK"  ' [ADR-0100]
End Function
```
