---
title: modEntryFormat.bas
---

# modEntryFormat.bas

| 項目 | 値 |
|---|---|
| 層 | エントリポイント層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | フォーマット一覧シートのマクロボタン受け口 |
| 行数 | 223 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modEntryFormat` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modEntryFormat"
Option Explicit

' ================================================================
' モジュール: modEntryFormat（エントリポイント層）
' 概要:       フォーマット管理関連のボタン（一覧・設計・プレビュー）に
'             割り当てるマクロ群
' 依存先:     clsLogger, clsFormatManager, clsTaskController,
'             modEntryMain, modCommon
' ================================================================

' ================================================================
' 関数名: Btn_CreateNewFormat
' 概要:   フォーマット一覧「▶新規作成」ボタン
'         設計シートを空にして遷移
' ================================================================
Public Sub Btn_CreateNewFormat()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryFormat", "Btn_CreateNewFormat", _
                    "新規作成ボタン押下"
    
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init logger
    mgr.BeginCreate
    
    ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN).Activate
    Exit Sub

ErrHandler:
    Call ShowError("フォーマット新規作成", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_EditFormat
' 概要:   フォーマット一覧「▶選択行を編集」ボタン
'         選択行のフォーマットを設計シートに読み込んで遷移
' ================================================================
Public Sub Btn_EditFormat()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init logger
    
    Dim formatId As String
    formatId = mgr.GetSelectedFormatId()
    
    If formatId = "" Then
        Call ShowWarning("フォーマット編集", _
                         "フォーマットが選択されていません", _
                         "一覧から編集したい行を選択してからボタンを押してください")
        Exit Sub
    End If
    
    logger.LogInfo "modEntryFormat", "Btn_EditFormat", _
                    "編集開始: " & formatId
    
    mgr.BeginEdit formatId
    ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN).Activate
    Exit Sub

ErrHandler:
    Call ShowError("フォーマット編集", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_ShowFormatPreview
' 概要:   フォーマット一覧「▶プレビュー」ボタン
' ================================================================
Public Sub Btn_ShowFormatPreview()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init logger
    
    Dim formatId As String
    formatId = mgr.GetSelectedFormatId()
    
    If formatId = "" Then
        Call ShowWarning("フォーマットプレビュー", _
                         "フォーマットが選択されていません", _
                         "一覧から表示したい行を選択してからボタンを押してください")
        Exit Sub
    End If
    
    logger.LogInfo "modEntryFormat", "Btn_ShowFormatPreview", _
                    "プレビュー: " & formatId
    
    mgr.ShowPreview formatId
    ThisWorkbook.Worksheets(SHEET_FORMAT_PREVIEW).Activate
    Exit Sub

ErrHandler:
    Call ShowError("プレビュー表示", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_LoadFormat
' 概要:   フォーマット設計「▶読込」ボタン
'         フォーマットID欄に入力されたIDで設計シートに読み込み
' ================================================================
Public Sub Btn_LoadFormat()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    Dim formatId As String
    formatId = CStr(ws.Cells(1, 2).Value)
    
    If formatId = "" Then
        Call ShowWarning("フォーマット読込", _
                         "フォーマットIDが入力されていません", _
                         "上部のフォーマットID欄に入力してから読込ボタンを押してください")
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryFormat", "Btn_LoadFormat", _
                    "読込: " & formatId
    
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init logger
    mgr.BeginEdit formatId
    Exit Sub

ErrHandler:
    Call ShowError("フォーマット読込", Err.Description, _
                    "フォーマットIDを確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_SaveFormat
' 概要:   フォーマット設計「▶保存」ボタン
'         設計内容をフォーマット一覧に保存（新規/既存の自動判定）
' ================================================================
Public Sub Btn_SaveFormat()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryFormat", "Btn_SaveFormat", _
                    "保存ボタン押下"
    
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init logger
    
    If mgr.SaveFormat() Then
        Call ShowInfo("フォーマット保存", "保存が完了しました")
    Else
        Call ShowError("フォーマット保存", "保存に失敗しました", _
                        "入力内容を確認してから再度ボタンを押してください")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("フォーマット保存", Err.Description, _
                    "入力内容を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_PreviewFormat
' 概要:   フォーマット設計「▶プレビュー」ボタン
' ================================================================
Public Sub Btn_PreviewFormat()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORMAT_DESIGN)
    Dim formatId As String
    formatId = CStr(ws.Cells(1, 2).Value)
    
    If formatId = "" Then
        Call ShowWarning("プレビュー", _
                         "フォーマットIDが入力されていません", _
                         "設計シートのフォーマットID欄に入力してください")
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init logger
    mgr.ShowPreview formatId
    
    ThisWorkbook.Worksheets(SHEET_FORMAT_PREVIEW).Activate
    Exit Sub

ErrHandler:
    Call ShowError("プレビュー表示", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_BackToList
' 概要:   フォーマット設計「←一覧に戻る」ボタン
' ================================================================
Public Sub Btn_BackToList()
    On Error GoTo ErrHandler
    ThisWorkbook.Worksheets(SHEET_FORMAT_LIST).Activate
    Exit Sub
ErrHandler:
    Call ShowError("シート遷移", Err.Description, _
                    "再度ボタンを押してください")
End Sub

```
