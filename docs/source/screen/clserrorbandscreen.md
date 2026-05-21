---
title: clsErrorBandScreen.cls
---

# clsErrorBandScreen.cls

| 項目 | 内容 |
|---|---|
| 層 | 画面層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 管理.xlsm |
| 役割 | M-13 のエラーバンド（4 列）表示を担う連動画面クラス |
| 行数 | 88 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsErrorBandScreen.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorBandScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' Class: clsErrorBandScreen (Screen layer - error band screen)
' Overview: M-13 file format settings error band screen (4-column schema).
'           spec is delegated to modScreenRender / RenderStandardScreen.
' Dependencies: IScreenRenderer, clsScreenSpec, modScreenRender
' Notes: v2.1 (2026-05-20) new screen (placement_v2 Q27, 4-column schema).
'        Uses logging convention v1 logId only (no legacy LOG-* prefix).
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
    Call modScreenRender.LogScreenTrace("clsErrorBandScreen", "Setup", "ENTER sid=" & m_spec.ScreenId, "SAVE-EXIT-OK-II-057")

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsErrorBandScreen", "Setup", "EXIT ok", "SAVE-EXIT-OK-II-058")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsErrorBandScreen", "Setup", stepName, Err.Number, Err.Description, "BACKTOMAIN-ERR-EE-054")
End Sub

Public Sub Render()
    Setup
End Sub

Public Function ValidateInput() As Boolean
    On Error GoTo ErrHandler
    ' Screen-level input precondition check (error band, 4-column schema).
    ' (a) spec must be initialised; (b) every Required field carrying an
    ' input cell address must hold a non-empty value on the bound sheet.
    ' Detailed 4-column schema validation stays in the entry layer.
    If m_spec Is Nothing Then
        Call modScreenRender.LogScreenTrace("clsErrorBandScreen", "ValidateInput", "spec not initialised", "VALIDATE-WARN-WW-215")
        ValidateInput = False
        Exit Function
    End If

    Dim fld As clsFieldSpec
    For Each fld In m_spec.Fields
        If fld.Required And Len(fld.InputAddr) > 0 Then
            If Len(Trim$(GetInputValue(fld.InputAddr))) = 0 Then
                Call modScreenRender.LogScreenTrace("clsErrorBandScreen", "ValidateInput", "required field empty: " & fld.Label, "VALIDATE-WARN-WW-215")
                ValidateInput = False
                Exit Function
            End If
        End If
    Next fld

    ValidateInput = True
    Exit Function

ErrHandler:
    Call modScreenRender.LogScreenError("clsErrorBandScreen", "ValidateInput", "validate", Err.Number, Err.Description, "VALIDATE-WARN-WW-215")
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
