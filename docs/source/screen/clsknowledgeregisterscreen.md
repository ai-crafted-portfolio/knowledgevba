---
title: clsKnowledgeRegisterScreen.cls
---

# clsKnowledgeRegisterScreen.cls

| 項目 | 内容 |
|---|---|
| 層 | 画面層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 登録修正.xlsm |
| 役割 | M-05 ナレッジ登録画面の構築・再描画 |
| 行数 | 92 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsKnowledgeRegisterScreen.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKnowledgeRegisterScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsKnowledgeRegisterScreen (画面層 - M-05 ナレッジ登録)
' 概要:   M-05 画面の構築・再描画 + Btn handler 経路の logId emit。
'         spec を modScreenRender に委譲する標準パターン。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenRender
' 配置:   S-登 (登録修正.xlsm)、placement_v2 §47
' 備考:   v2.1 (2026-05-20) で publish 投入。buildtest archive 旧版
'         (45L、LOG-M05-SCREENCLS-SETUP-*) を踏襲しつつ logging 規約
'         HandleButton メソッドは削除済 (2026-05-20)。ボタン処理の正規実装は modEntry* の Btn_*_v21 のみ。
'         clsMigrationScreen / clsBackupMgmtScreen と同テンプレ。
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
    Call modScreenRender.LogScreenTrace("clsKnowledgeRegisterScreen", "Setup", "ENTER sid=" & m_spec.ScreenId, "LOG-M05-SCREENCLS-SETUP-ENTRY")

    stepName = "RenderStandardScreen"
    Call modScreenRender.RenderStandardScreen(m_renderer, m_spec)

    Call modScreenRender.LogScreenTrace("clsKnowledgeRegisterScreen", "Setup", "EXIT ok", "LOG-M05-SCREENCLS-SETUP-EXIT-OK")
    Exit Sub

ErrHandler:
    Call modScreenRender.LogScreenError("clsKnowledgeRegisterScreen", "Setup", stepName, Err.Number, Err.Description, "LOG-M05-SCREENCLS-SETUP-ERR")
End Sub

Public Sub Render()
    Setup
End Sub

Public Function ValidateInput() As Boolean
    On Error GoTo ErrHandler
    ' Screen-level input precondition check (M-05 knowledge register).
    ' (a) spec must be initialised; (b) every Required field carrying an
    ' input cell address must hold a non-empty value on the bound sheet.
    ' Detailed knowledge No / required field validation stays in the
    ' entry layer (modEntryKnowledge.Btn_SaveKnowledge).
    If m_spec Is Nothing Then
        Call modScreenRender.LogScreenTrace("clsKnowledgeRegisterScreen", "ValidateInput", "spec not initialised", "VALIDATE-WARN-WW-205")
        ValidateInput = False
        Exit Function
    End If

    Dim fld As clsFieldSpec
    For Each fld In m_spec.Fields
        If fld.Required And Len(fld.InputAddr) > 0 Then
            If Len(Trim$(GetInputValue(fld.InputAddr))) = 0 Then
                Call modScreenRender.LogScreenTrace("clsKnowledgeRegisterScreen", "ValidateInput", "required field empty: " & fld.Label, "VALIDATE-WARN-WW-205")
                ValidateInput = False
                Exit Function
            End If
        End If
    Next fld

    ValidateInput = True
    Exit Function

ErrHandler:
    Call modScreenRender.LogScreenError("clsKnowledgeRegisterScreen", "ValidateInput", "validate", Err.Number, Err.Description, "VALIDATE-WARN-WW-205")
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
