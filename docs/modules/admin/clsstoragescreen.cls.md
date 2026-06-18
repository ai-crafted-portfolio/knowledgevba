---
title: clsStorageScreen.cls
description: clsStorageScreen.cls のソースコード（コピペ用）
---

# clsStorageScreen.cls

**配置先**: `管理.xlsm` 用の VBA モジュール
**種類**: クラスモジュール
**更新日**: 2026-06-03 23:22 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\admin\`
- ファイル名: `clsStorageScreen.cls`
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
Attribute VB_Name = "clsStorageScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' Class: clsStorageScreen (Screen layer - M-10 storage management screen)
' Overview: M-10 storage management frame. spec is delegated to modScreenRender.
' Dependencies: IScreenRenderer, clsScreenSpec, modScreenRender
' Notes: v2.1 (2026-05-20) publish promotion. Carries the buildtest archive
'        template while adding logging convention v1 emit points.
'        Legacy LOG-M10-SCREENCLS-SETUP-* IDs retained for backward compatibility.
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0098] clsStorageScreen.Init ENTER"  ' [ADR-0100]
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0099] clsStorageScreen.Setup ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsStorageScreen", "Setup", "ENTER sid=" & m_spec.ScreenId, "LOG-M10-SCREENCLS-SETUP-ENTRY")

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsStorageScreen", "Setup", "EXIT ok", "LOG-M10-SCREENCLS-SETUP-EXIT-OK")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0100] clsStorageScreen.Setup EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0101] clsStorageScreen.Setup EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Call modScreenRender.LogScreenError("clsStorageScreen", "Setup", stepName, Err.Number, Err.Description, "LOG-M10-SCREENCLS-SETUP-ERR")
End Sub

Public Sub Render()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0102] clsStorageScreen.Render ENTER"  ' [ADR-0100]
    Setup
End Sub

Public Function ValidateInput() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0103] clsStorageScreen.ValidateInput ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' Screen-level input precondition check (M-10 storage).
    ' (a) spec must be initialised; (b) every Required field carrying an
    ' input cell address must hold a non-empty value on the bound sheet.
    ' Detailed per-row storage validation stays in the entry layer.
    If m_spec Is Nothing Then
        If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0108] clsStorageScreen.ValidateInput STEP-1 pre modScreenRender.LogScreenTrace"  ' [ADR-0100]
        Call modScreenRender.LogScreenTrace("clsStorageScreen", "ValidateInput", "spec not initialised", "VALIDATE-WARN-WW-210")
        ValidateInput = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0104] clsStorageScreen.ValidateInput EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim fld As clsFieldSpec
    For Each fld In m_spec.Fields
        If fld.Required And Len(fld.InputAddr) > 0 Then
            If Len(Trim$(GetInputValue(fld.InputAddr))) = 0 Then
                If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0109] clsStorageScreen.ValidateInput STEP-2 pre modScreenRender.LogScreenTrace"  ' [ADR-0100]
                Call modScreenRender.LogScreenTrace("clsStorageScreen", "ValidateInput", "required field empty: " & fld.Label, "VALIDATE-WARN-WW-210")
                ValidateInput = False
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0105] clsStorageScreen.ValidateInput EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
        End If
    Next fld

    ValidateInput = True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0106] clsStorageScreen.ValidateInput EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0110] clsStorageScreen.ValidateInput STEP-3 pre modScreenRender.LogScreenError"  ' [ADR-0100]
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0107] clsStorageScreen.ValidateInput EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Call modScreenRender.LogScreenError("clsStorageScreen", "ValidateInput", "validate", Err.Number, Err.Description, "VALIDATE-WARN-WW-210")
    ValidateInput = False
End Function

Private Function GetInputValue(ByVal cellAddr As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0111] clsStorageScreen.GetInputValue ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(m_spec.SheetName)
    If ws Is Nothing Then
        GetInputValue = ""
    Else
        GetInputValue = CStr(ws.Range(cellAddr).Value)
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0112] clsStorageScreen.GetInputValue EXIT-OK"  ' [ADR-0100]
End Function
```
