---
title: clsMigrationScreen.cls
---

# clsMigrationScreen.cls

| 項目 | 内容 |
|---|---|
| 層 | 画面層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 管理.xlsm |
| 役割 | M-12 フィールド反映画面の構築・再描画 |
| 行数 | 88 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsMigrationScreen.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMigrationScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' Class: clsMigrationScreen (Screen layer - M-12 migration screen)
' Overview: M-12 migration frame. spec is delegated to modScreenRender.
' Dependencies: IScreenRenderer, clsScreenSpec, modScreenRender
' Notes: v2.1 (2026-05-20) publish promotion. Carries the buildtest archive
'        template while adding logging convention v1 emit points.
'        Legacy LOG-M12-SCREENCLS-SETUP-* IDs retained for backward compatibility.
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
    Call modScreenRender.LogScreenTrace("clsMigrationScreen", "Setup", "ENTER sid=" & m_spec.ScreenId, "LOG-M12-SCREENCLS-SETUP-ENTRY")

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsMigrationScreen", "Setup", "EXIT ok", "LOG-M12-SCREENCLS-SETUP-EXIT-OK")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsMigrationScreen", "Setup", stepName, Err.Number, Err.Description, "LOG-M12-SCREENCLS-SETUP-ERR")
End Sub

Public Sub Render()
    Setup
End Sub

Public Function ValidateInput() As Boolean
    On Error GoTo ErrHandler
    ' Screen-level input precondition check (M-12 migration).
    ' (a) spec must be initialised; (b) every Required field carrying an
    ' input cell address must hold a non-empty value on the bound sheet.
    ' Detailed migration pre-run validation stays in the entry layer.
    If m_spec Is Nothing Then
        Call modScreenRender.LogScreenTrace("clsMigrationScreen", "ValidateInput", "spec not initialised", "VALIDATE-WARN-WW-212")
        ValidateInput = False
        Exit Function
    End If

    Dim fld As clsFieldSpec
    For Each fld In m_spec.Fields
        If fld.Required And Len(fld.InputAddr) > 0 Then
            If Len(Trim$(GetInputValue(fld.InputAddr))) = 0 Then
                Call modScreenRender.LogScreenTrace("clsMigrationScreen", "ValidateInput", "required field empty: " & fld.Label, "VALIDATE-WARN-WW-212")
                ValidateInput = False
                Exit Function
            End If
        End If
    Next fld

    ValidateInput = True
    Exit Function

ErrHandler:
    Call modScreenRender.LogScreenError("clsMigrationScreen", "ValidateInput", "validate", Err.Number, Err.Description, "VALIDATE-WARN-WW-212")
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
