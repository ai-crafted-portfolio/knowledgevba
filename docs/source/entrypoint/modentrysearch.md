---
title: modEntrySearch.bas
---

# modEntrySearch.bas

| 項目 | 値 |
|---|---|
| 層 | エントリポイント層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | 検索シート上のマクロボタン受け口 |
| 行数 | 244 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modEntrySearch` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modEntrySearch"
Option Explicit

' ================================================================
' モジュール: modEntrySearch（エントリポイント層）
' 概要:       検索・ナレッジ表示関連のボタンに割り当てるマクロ群
' 依存先:     clsLogger, clsSearchEngine, clsTaskController,
'             modEntryMain, modCommon
' ================================================================

' --- 検索シート 位置定数 ---
Private Const SS_ROW_DIRECT_NO As Long = 3
Private Const SS_COL_DIRECT_NO As Long = 3
Private Const SS_ROW_FMT_ID As Long = 6
Private Const SS_ROW_KEYWORDS As Long = 7
Private Const SS_ROW_MODE As Long = 8
Private Const SS_ROW_TARGET_FIELD As Long = 9
Private Const SS_ROW_DATE_FROM As Long = 10
Private Const SS_ROW_DATE_TO As Long = 11
Private Const SS_COL_CONDITION As Long = 3
Private Const SS_RESULT_START_ROW As Long = 15

' --- ナレッジ表示シート 位置定数 ---
Private Const KD_ROW_KNW_NO As Long = 1
Private Const KD_COL_KNW_NO_VAL As Long = 2

' ================================================================
' 関数名: Btn_SearchKnowledge
' 概要:   検索シート「▶検索」ボタン
' ================================================================
Public Sub Btn_SearchKnowledge()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntrySearch", "Btn_SearchKnowledge", _
                    "検索ボタン押下"
    
    Dim engine As clsSearchEngine
    Set engine = New clsSearchEngine
    engine.Init logger, GetDataFolder()
    
    Dim hits As Long
    hits = engine.SearchByKeywords()
    
    If hits = 0 Then
        Call ShowInfo("検索", "条件に合致するナレッジが見つかりませんでした")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ナレッジ検索", Err.Description, _
                    "検索条件を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_DirectLookup
' 概要:   検索シート「▶表示」ボタン（番号ダイレクト検索）
' ================================================================
Public Sub Btn_DirectLookup()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)
    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(SS_ROW_DIRECT_NO, SS_COL_DIRECT_NO).Value)
    
    If knowledgeNo = "" Then
        Call ShowWarning("番号ダイレクト検索", _
                         "ナレッジ番号が入力されていません", _
                         "上部の番号欄に入力してから表示ボタンを押してください")
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim engine As clsSearchEngine
    Set engine = New clsSearchEngine
    engine.Init logger, GetDataFolder()
    
    If Not engine.FindByNumber(knowledgeNo) Then
        Call ShowError("番号ダイレクト検索", _
                        "指定されたナレッジが見つかりません", _
                        "ナレッジ番号を確認してから再度ボタンを押してください")
        Exit Sub
    End If
    
    engine.DisplayKnowledge knowledgeNo
    
    Dim controller As clsTaskController
    Set controller = New clsTaskController
    controller.Init logger
    controller.SwitchToTask TASK_SEARCH
    
    ThisWorkbook.Worksheets(SHEET_KNW_DISPLAY).Activate
    Exit Sub

ErrHandler:
    Call ShowError("番号ダイレクト検索", Err.Description, _
                    "ナレッジ番号を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_SearchClear
' 概要:   検索シート「▶クリア」ボタン
' ================================================================
Public Sub Btn_SearchClear()
    On Error GoTo ErrHandler
    
    If Not ConfirmAction("検索条件クリア", _
                           "入力中の検索条件を全てクリアします") Then
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)
    
    ws.Cells(SS_ROW_DIRECT_NO, SS_COL_DIRECT_NO).Value = ""
    ws.Cells(SS_ROW_FMT_ID, SS_COL_CONDITION).Value = ""
    ws.Cells(SS_ROW_KEYWORDS, SS_COL_CONDITION).Value = ""
    ws.Cells(SS_ROW_MODE, SS_COL_CONDITION).Value = ""
    ws.Cells(SS_ROW_TARGET_FIELD, SS_COL_CONDITION).Value = ""
    ws.Cells(SS_ROW_DATE_FROM, SS_COL_CONDITION).Value = ""
    ws.Cells(SS_ROW_DATE_TO, SS_COL_CONDITION).Value = ""
    
    ' 結果一覧もクリア
    ' rev22: 検索シートに操作フロー注記用の結合セル (B19:H19) があり
    ' A:G の一括 ClearContents が "結合したセルには行えません" で
    ' 落ちる。行ごとに Resume Next で結合行をスキップ。
    Dim i As Long
    For i = SS_RESULT_START_ROW To SS_RESULT_START_ROW + 100
        On Error Resume Next
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 7)).ClearContents
        Err.Clear
        On Error GoTo 0
    Next i
    Exit Sub

ErrHandler:
    Call ShowError("検索条件クリア", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_DetailDisplay
' 概要:   検索結果「▶詳細」ボタン
'         選択行のナレッジを表示シートに展開
' ================================================================
Public Sub Btn_DetailDisplay()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)
    
    Dim selRow As Long
    selRow = ws.Application.Selection.Row
    
    If selRow < SS_RESULT_START_ROW Then
        Call ShowWarning("詳細表示", _
                         "結果行が選択されていません", _
                         "結果一覧から表示したい行を選択してから再度実行してください")
        Exit Sub
    End If
    
    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(selRow, 2).Value)
    
    If knowledgeNo = "" Then
        Call ShowWarning("詳細表示", _
                         "選択行にナレッジ番号がありません", _
                         "先に検索を実行してから結果行を選択してください")
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim engine As clsSearchEngine
    Set engine = New clsSearchEngine
    engine.Init logger, GetDataFolder()
    engine.DisplayKnowledge knowledgeNo
    
    ThisWorkbook.Worksheets(SHEET_KNW_DISPLAY).Activate
    Exit Sub

ErrHandler:
    Call ShowError("詳細表示", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_BackToSearch
' 概要:   ナレッジ表示「←検索に戻る」ボタン
' ================================================================
Public Sub Btn_BackToSearch()
    On Error GoTo ErrHandler
    ThisWorkbook.Worksheets(SHEET_SEARCH).Activate
    Exit Sub
ErrHandler:
    Call ShowError("シート遷移", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_GoToEdit
' 概要:   ナレッジ表示「▶修正に遷移」ボタン
'         表示中のナレッジ番号で修正シートに遷移
' ================================================================
Public Sub Btn_GoToEdit()
    On Error GoTo ErrHandler
    
    Dim displayWs As Worksheet
    Set displayWs = ThisWorkbook.Worksheets(SHEET_KNW_DISPLAY)
    
    Dim knowledgeNo As String
    knowledgeNo = CStr(displayWs.Cells(KD_ROW_KNW_NO, KD_COL_KNW_NO_VAL).Value)
    
    If knowledgeNo = "" Then
        Call ShowWarning("修正遷移", _
                         "表示中のナレッジ番号が取得できません", _
                         "検索または番号ダイレクト表示でナレッジを開いてから再度実行してください")
        Exit Sub
    End If
    
    Dim editWs As Worksheet
    Set editWs = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)
    editWs.Cells(1, 3).Value = knowledgeNo
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim controller As clsTaskController
    Set controller = New clsTaskController
    controller.Init logger
    controller.SwitchToTask TASK_EDIT
    
    editWs.Activate
    Exit Sub

ErrHandler:
    Call ShowError("修正遷移", Err.Description, _
                    "再度ボタンを押してください")
End Sub

```
