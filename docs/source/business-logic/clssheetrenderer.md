---
title: clsSheetRenderer.cls
---

# clsSheetRenderer.cls

| 項目 | 内容 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | 画面描画インターフェースのシート実装。レイアウト適用を modUILoader に委譲する |
| 行数 | 112 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsSheetRenderer.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "clsSheetRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' クラス: clsSheetRenderer（v2.1、Phase 3 中核、thin shim 化）
' 概要:   IScreenRenderer の Sheet 型実装、modUILoader.ApplyUiToSheet に集約
' Version: v2.1（2026-05-16 EOD、Q20/Q5 反映）
' 関連:   ADR-0053 §6（外部 UI スタンザ駆動）, IScreenRenderer.cls
' v2.1 主要更新:
'   - v1 521 行 → v2.1 約 100 行 thin shim
'   - 個別 Render* メソッド廃止、ApplyFromStanza 1 行集約
'   - modUILoader.ApplyUiToSheet（Q20 関数名）を呼出
' ================================================================
Implements IScreenRenderer
Option Explicit

Private m_targetSheet As Worksheet
Private m_xlsmName As String

' ----------------------------------------------------------------
' IScreenRenderer 実装（8 メソッド、v2.1）
' ----------------------------------------------------------------

Private Sub IScreenRenderer_BindSheet(ByVal sheetName As String)
    On Error GoTo ErrHandler
    Set m_targetSheet = ThisWorkbook.Worksheets(sheetName)
    Exit Sub
ErrHandler:
    Err.Raise vbObjectError + 100, "clsSheetRenderer", _
              "BindSheet 失敗: " & sheetName & " - " & Err.Description
End Sub

Private Sub IScreenRenderer_ClearScreen()
    If m_targetSheet Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    m_targetSheet.Cells.Clear
    ' Shape 系（ボタン等）も削除
    Dim shp As Shape
    For Each shp In m_targetSheet.Shapes
        shp.Delete
    Next shp
    Exit Sub
ErrHandler:
    Debug.Print "[clsSheetRenderer.ClearScreen ERROR] " & Err.Description
End Sub

' v2.1 中核：1 行で UI 全構築（Q20 modUILoader.ApplyUiToSheet 呼出）
Private Sub IScreenRenderer_ApplyFromStanza(ByVal xlsmName As String, ByVal screenId As String)
    On Error GoTo ErrHandler
    If m_targetSheet Is Nothing Then
        Err.Raise vbObjectError + 101, "clsSheetRenderer", _
                  "ApplyFromStanza: BindSheet 未実行"
    End If
    m_xlsmName = xlsmName
    Call modUILoader.ApplyUiToSheet(xlsmName, screenId, m_targetSheet)
    Exit Sub
ErrHandler:
    Err.Raise vbObjectError + 102, "clsSheetRenderer", _
              "ApplyFromStanza 失敗: " & xlsmName & "/" & screenId & " - " & Err.Description
End Sub

Private Sub IScreenRenderer_ShowSheet()
    If m_targetSheet Is Nothing Then Exit Sub
    m_targetSheet.Visible = xlSheetVisible
End Sub

Private Sub IScreenRenderer_HideSheet()
    If m_targetSheet Is Nothing Then Exit Sub
    m_targetSheet.Visible = xlSheetHidden  ' Q2: VeryHidden ではない
End Sub

Private Sub IScreenRenderer_ActivateSheet()
    If m_targetSheet Is Nothing Then Exit Sub
    m_targetSheet.Activate
End Sub

Private Sub IScreenRenderer_ProtectSheet(ByVal level As String)
    If m_targetSheet Is Nothing Then Exit Sub
    On Error Resume Next
    Select Case LCase(level)
        Case "light"
            ' light 保護：ユーザがセル選択は可、書込は不可
            m_targetSheet.Protect Password:="", AllowFormattingCells:=True, _
                                   AllowFormattingColumns:=True, AllowFormattingRows:=True
        Case "strict"
            ' strict 保護：全 lock
            m_targetSheet.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True
        Case Else
            m_targetSheet.Protect Password:=""
    End Select
    ' Q7 規約 Y：On Error Resume Next 直後の Err check
    If Err.Number <> 0 Then
        Debug.Print "[clsSheetRenderer.ProtectSheet ERROR] " & Err.Description
        Err.Clear
    End If
End Sub

Private Sub IScreenRenderer_UnprotectSheet()
    If m_targetSheet Is Nothing Then Exit Sub
    On Error Resume Next
    m_targetSheet.Unprotect Password:=""
    If Err.Number <> 0 Then
        Debug.Print "[clsSheetRenderer.UnprotectSheet ERROR] " & Err.Description
        Err.Clear
    End If
End Sub
```
