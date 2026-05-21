---
title: clsFormatPreviewScreen.cls
---

# clsFormatPreviewScreen.cls

| 項目 | 内容 |
|---|---|
| 層 | 画面層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 管理.xlsm |
| 役割 | M-04 フォーマットプレビュー画面の構築・再描画 |
| 行数 | 88 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsFormatPreviewScreen.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormatPreviewScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsFormatPreviewScreen (画面層 - M-04 フォーマットプレビュー)
' 概要:   M-04 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   v2.1 (2026-05-20) で publish 投入。buildtest archive 旧版を
'         踏襲しつつ logging 規約 v1 準拠の新規 emit を追加。
'         既存 LOG-M04-SCREENCLS-SETUP-* 系は §3.3 後方互換維持で残置。
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
    Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "Setup", "ENTER sid=" & m_spec.ScreenId, "LOG-M04-SCREENCLS-SETUP-ENTRY")

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "Setup", "EXIT ok", "LOG-M04-SCREENCLS-SETUP-EXIT-OK")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsFormatPreviewScreen", "Setup", stepName, Err.Number, Err.Description, "LOG-M04-SCREENCLS-SETUP-ERR")
End Sub

Public Sub Render()
    Setup
End Sub

Public Function ValidateInput() As Boolean
    On Error GoTo ErrHandler
    ' Screen-level input precondition check (M-04 format preview).
    ' The preview screen is display-only and defines no Required input
    ' fields, so an initialised spec validates True; an uninitialised
    ' screen (m_spec Is Nothing) fails the precondition.
    If m_spec Is Nothing Then
        Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "ValidateInput", "spec not initialised", "VALIDATE-WARN-WW-204")
        ValidateInput = False
        Exit Function
    End If

    Dim fld As clsFieldSpec
    For Each fld In m_spec.Fields
        If fld.Required And Len(fld.InputAddr) > 0 Then
            If Len(Trim$(GetInputValue(fld.InputAddr))) = 0 Then
                Call modScreenRender.LogScreenTrace("clsFormatPreviewScreen", "ValidateInput", "required field empty: " & fld.Label, "VALIDATE-WARN-WW-204")
                ValidateInput = False
                Exit Function
            End If
        End If
    Next fld

    ValidateInput = True
    Exit Function

ErrHandler:
    Call modScreenRender.LogScreenError("clsFormatPreviewScreen", "ValidateInput", "validate", Err.Number, Err.Description, "VALIDATE-WARN-WW-204")
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
