---
title: clsSystemSettingsScreen.cls
description: clsSystemSettingsScreen.cls のソースコード（コピペ用）
---

# clsSystemSettingsScreen.cls

**配置先**: `管理.xlsm` 用の VBA モジュール
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\admin\`
- ファイル名: `clsSystemSettingsScreen.cls`
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
Attribute VB_Name = "clsSystemSettingsScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsSystemSettingsScreen (画面層 - M-11 システム設定画面)
' 概要:   システム設定画面 (M-11 frame)。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   v2.1 (2026-05-20) で publish 投入。buildtest archive 旧版を踏襲
'         しつつ logging 規約 v1 準拠の新規 emit を追加。
'         既存 LOG-M11-SCREENCLS-SETUP-* 系は §3.3 後方互換維持で残置。
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0113] clsSystemSettingsScreen.Init ENTER"  ' [ADR-0100]
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0114] clsSystemSettingsScreen.Setup ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsSystemSettingsScreen", "Setup", "ENTER sid=" & m_spec.ScreenId, "LOG-M11-SCREENCLS-SETUP-ENTRY")

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsSystemSettingsScreen", "Setup", "EXIT ok", "LOG-M11-SCREENCLS-SETUP-EXIT-OK")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0115] clsSystemSettingsScreen.Setup EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0116] clsSystemSettingsScreen.Setup EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Call modScreenRender.LogScreenError("clsSystemSettingsScreen", "Setup", stepName, Err.Number, Err.Description, "LOG-M11-SCREENCLS-SETUP-ERR")
End Sub

Public Sub Render()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0117] clsSystemSettingsScreen.Render ENTER"  ' [ADR-0100]
    Setup
End Sub

Public Function ValidateInput() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0118] clsSystemSettingsScreen.ValidateInput ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' Screen-level input precondition check (M-11 system settings).
    ' (a) spec must be initialised; (b) every Required field carrying an
    ' input cell address must hold a non-empty value on the bound sheet.
    ' Detailed settings validation stays in the entry layer.
    If m_spec Is Nothing Then
        If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0123] clsSystemSettingsScreen.ValidateInput STEP-1 pre modScreenRender.LogScreenTrace"  ' [ADR-0100]
        Call modScreenRender.LogScreenTrace("clsSystemSettingsScreen", "ValidateInput", "spec not initialised", "VALIDATE-WARN-WW-211")
        ValidateInput = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0119] clsSystemSettingsScreen.ValidateInput EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim fld As clsFieldSpec
    For Each fld In m_spec.Fields
        If fld.Required And Len(fld.InputAddr) > 0 Then
            If Len(Trim(GetInputValue(fld.InputAddr))) = 0 Then
                If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0124] clsSystemSettingsScreen.ValidateInput STEP-2 pre modScreenRender.LogScreenTrace"  ' [ADR-0100]
                Call modScreenRender.LogScreenTrace("clsSystemSettingsScreen", "ValidateInput", "required field empty: " & fld.Label, "VALIDATE-WARN-WW-211")
                ValidateInput = False
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0120] clsSystemSettingsScreen.ValidateInput EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
        End If
    Next fld

    ValidateInput = True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0121] clsSystemSettingsScreen.ValidateInput EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0125] clsSystemSettingsScreen.ValidateInput STEP-3 pre modScreenRender.LogScreenError"  ' [ADR-0100]
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0122] clsSystemSettingsScreen.ValidateInput EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Call modScreenRender.LogScreenError("clsSystemSettingsScreen", "ValidateInput", "validate", Err.Number, Err.Description, "VALIDATE-WARN-WW-211")
    ValidateInput = False
End Function

Private Function GetInputValue(ByVal cellAddr As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0126] clsSystemSettingsScreen.GetInputValue ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(m_spec.SheetName)
    If ws Is Nothing Then
        GetInputValue = ""
    Else
        GetInputValue = CStr(ws.Range(cellAddr).Value)
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0127] clsSystemSettingsScreen.GetInputValue EXIT-OK"  ' [ADR-0100]
End Function
```
