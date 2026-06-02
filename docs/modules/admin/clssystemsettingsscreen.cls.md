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
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

Public Sub Setup()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Call modScreenRender.LogScreenTrace("clsSystemSettingsScreen", "Setup", "ENTER sid=" & m_spec.ScreenId, "LOG-M11-SCREENCLS-SETUP-ENTRY")

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsSystemSettingsScreen", "Setup", "EXIT ok", "LOG-M11-SCREENCLS-SETUP-EXIT-OK")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsSystemSettingsScreen", "Setup", stepName, Err.Number, Err.Description, "LOG-M11-SCREENCLS-SETUP-ERR")
End Sub

Public Sub Render()
    Setup
End Sub

Public Function ValidateInput() As Boolean
    On Error GoTo ErrHandler
    ' Screen-level input precondition check (M-11 system settings).
    ' (a) spec must be initialised; (b) every Required field carrying an
    ' input cell address must hold a non-empty value on the bound sheet.
    ' Detailed settings validation stays in the entry layer.
    If m_spec Is Nothing Then
        Call modScreenRender.LogScreenTrace("clsSystemSettingsScreen", "ValidateInput", "spec not initialised", "VALIDATE-WARN-WW-211")
        ValidateInput = False
        Exit Function
    End If

    Dim fld As clsFieldSpec
    For Each fld In m_spec.Fields
        If fld.Required And Len(fld.InputAddr) > 0 Then
            If Len(Trim$(GetInputValue(fld.InputAddr))) = 0 Then
                Call modScreenRender.LogScreenTrace("clsSystemSettingsScreen", "ValidateInput", "required field empty: " & fld.Label, "VALIDATE-WARN-WW-211")
                ValidateInput = False
                Exit Function
            End If
        End If
    Next fld

    ValidateInput = True
    Exit Function

ErrHandler:
    Call modScreenRender.LogScreenError("clsSystemSettingsScreen", "ValidateInput", "validate", Err.Number, Err.Description, "VALIDATE-WARN-WW-211")
    ValidateInput = False
End Function

Private Function GetInputValue(ByVal cellAddr As String) As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(m_spec.SheetName)
    If ws Is Nothing Then
        GetInputValue = ""
    Else
        GetInputValue = CStr(ws.Range(cellAddr).Value)
    End If
End Function
```
