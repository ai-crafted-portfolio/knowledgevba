---
title: modSpecExamples.bas
---

# modSpecExamples.bas

| 項目 | 値 |
|---|---|
| 層 | エントリポイント層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | clsFormSpec DSL 利用例 (詳細プレビュー UserForm 構築) |
| 行数 | 102 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modSpecExamples` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modSpecExamples"
Option Explicit

' ================================================================
' モジュール: modSpecExamples (エントリポイント層)
' 概要:       clsFormSpec の使い方デモ。
'             検索結果プレビュー用 spec を組み立てる。
'             UI から呼び出さなくても Application.Run "Macro_Show..."
'             で手動起動できる。
' 依存先:     clsFormSpec, modFormBuilder, clsSearchEngine, modEntryMain
' ================================================================

' ================================================================
' 関数名: NewSearchResultSpec
' 概要:   検索結果プレビュー用フォーム spec を組み立てる
' 引数:   knwNo   - ナレッジ番号 (フォームタイトルに表示)
'         imgPath - 画像絶対パス ("" なら fallback テキスト)
'         snippet - 表示するスニペット (タイトル + 本文抜粋)
' 戻り値: clsFormSpec
' ================================================================
Public Function NewSearchResultSpec(ByVal knwNo As String, _
                                      ByVal imgPath As String, _
                                      ByVal snippet As String) As clsFormSpec
    Dim spec As clsFormSpec
    Set spec = New clsFormSpec
    spec.FormTitle = "検索結果プレビュー: " & knwNo
    spec.Width = 600
    spec.Height = 460

    ' Title Label (上部)
    Call spec.AddControl("Label", "lblTitle", 10, 10, 580, 20, _
                          "ナレッジ: " & knwNo, "")

    ' 画像領域 (左)
    Call spec.AddControl("Image", "imgPreview", 10, 40, 400, 300, "", "")

    ' スニペット (右)
    Call spec.AddControl("TextBox", "txtSnippet", 420, 40, 170, 300, _
                          snippet, "")

    ' 閉じるボタン
    Call spec.AddControl("Button", "btnClose", 250, 360, 80, 30, _
                          "閉じる", "frmCallback_searchResult_close")

    Set NewSearchResultSpec = spec
End Function

' ================================================================
' 関数名: Macro_ShowSearchResultPreview
' 概要:   現在 検索シートで選択中の行のナレッジを spec 駆動フォームで表示
' 備考:   Application.Run or Alt+F8 から手動呼出可能。
'         既存の Btn_DetailDisplay (M-09 シート遷移) と並立。
' ================================================================
Public Sub Macro_ShowSearchResultPreview()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)

    Dim selRow As Long
    selRow = ws.Application.Selection.Row
    If selRow < 15 Then
        Call ShowWarning("プレビュー", _
                          "検索結果の行が選択されていません", _
                          "結果一覧から行を選んで再実行してください")
        Exit Sub
    End If

    Dim knwNo As String
    knwNo = CStr(ws.Cells(selRow, 2).Value)
    If knwNo = "" Then Exit Sub

    Dim title As String
    title = CStr(ws.Cells(selRow, 4).Value)
    Dim score As String
    score = CStr(ws.Cells(selRow, 9).Value)

    Dim spec As clsFormSpec
    Set spec = NewSearchResultSpec(knwNo, "", _
                                     "Title: " & title & vbCrLf & _
                                     "Score: " & score & vbCrLf & vbCrLf & _
                                     "(画像と本文は ChromaDB 連動真版で展開)")

    Call BuildAndShow(spec, True)
    Exit Sub

ErrHandler:
    Call ShowError("プレビュー", Err.Description, _
                    "VBA プロジェクト オブジェクト モデル信頼が ON か確認")
End Sub

' ================================================================
' 関数名: frmCallback_searchResult_close
' 概要:   spec 駆動フォームの「閉じる」ボタン Click コールバック
' 引数:   frm - UserForm インスタンス (Application.Run の第 2 引数で渡る)
' 備考:   Application.Run の引数として UserForm が渡るが、本モックでは
'         単純に Unload するだけ。
' ================================================================
Public Sub frmCallback_searchResult_close(ByVal frm As Object)
    On Error Resume Next
    Unload frm
End Sub
```
