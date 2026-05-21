---
title: clsFormatDesignScreen.cls
---

# clsFormatDesignScreen.cls

| 項目 | 内容 |
|---|---|
| 層 | 画面層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 管理.xlsm |
| 役割 | M-03 フォーマット設計画面の構築・再描画 |
| 行数 | 89 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsFormatDesignScreen.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormatDesignScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsFormatDesignScreen (画面層 - M-03 フォーマット設計)
' 概要:   M-03 画面の構築・再描画。spec を modScreenRender に委譲。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 備考:   v2.1 (2026-05-20) で publish 投入。buildtest archive 旧版を
'         踏襲しつつ logging 規約 v1 準拠の新規 emit を追加。
'         既存 LOG-M03-SCREENCLS-SETUP-* 系は §3.3 後方互換維持で残置。
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
    Call modScreenRender.LogScreenTrace("clsFormatDesignScreen", "Setup", "ENTER sid=" & m_spec.ScreenId, "LOG-M03-SCREENCLS-SETUP-ENTRY")

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsFormatDesignScreen", "Setup", "EXIT ok", "LOG-M03-SCREENCLS-SETUP-EXIT-OK")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsFormatDesignScreen", "Setup", stepName, Err.Number, Err.Description, "LOG-M03-SCREENCLS-SETUP-ERR")
End Sub

Public Sub Render()
    Setup
End Sub

Public Function ValidateInput() As Boolean
    On Error GoTo ErrHandler
    ' Screen-level input precondition check (M-03 format design).
    ' (a) spec must be initialised; (b) every Required field carrying an
    ' input cell address must hold a non-empty value on the bound sheet.
    ' Detailed FormatID / FormatName validation stays in the entry layer
    ' (modEntryFormat.SaveFormat_Workflow).
    If m_spec Is Nothing Then
        Call modScreenRender.LogScreenTrace("clsFormatDesignScreen", "ValidateInput", "spec not initialised", "VALIDATE-WARN-WW-203")
        ValidateInput = False
        Exit Function
    End If

    Dim fld As clsFieldSpec
    For Each fld In m_spec.Fields
        If fld.Required And Len(fld.InputAddr) > 0 Then
            If Len(Trim$(GetInputValue(fld.InputAddr))) = 0 Then
                Call modScreenRender.LogScreenTrace("clsFormatDesignScreen", "ValidateInput", "required field empty: " & fld.Label, "VALIDATE-WARN-WW-203")
                ValidateInput = False
                Exit Function
            End If
        End If
    Next fld

    ValidateInput = True
    Exit Function

ErrHandler:
    Call modScreenRender.LogScreenError("clsFormatDesignScreen", "ValidateInput", "validate", Err.Number, Err.Description, "VALIDATE-WARN-WW-203")
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
