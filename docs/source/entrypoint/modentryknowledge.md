---
title: modEntryKnowledge.bas
---

# modEntryKnowledge.bas

| 項目 | 値 |
|---|---|
| 層 | エントリポイント層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | ナレッジ一覧シートのマクロボタン受け口 |
| 行数 | 360 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modEntryKnowledge` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modEntryKnowledge"
Option Explicit

' ================================================================
' モジュール: modEntryKnowledge（エントリポイント層）
' 概要:       ナレッジ登録・修正・削除・一覧・フィールド反映関連
'             のボタンに割り当てるマクロ群
' 依存先:     clsLogger, clsKnowledgeManager, clsFormatManager,
'             clsFieldMigrator, modEntryMain, modCommon
' ================================================================

' --- ナレッジ登録シート / ナレッジ修正シート 位置定数 ---
Private Const KS_ROW_FMT_ID As Long = 1
Private Const KS_COL_FMT_ID_VAL As Long = 3
Private Const KE_ROW_FMT_ID As Long = 1
Private Const KE_COL_KNW_NO As Long = 3

' --- ナレッジ一覧シート 位置定数 ---
Private Const KL_RESULT_START_ROW As Long = 4

' --- 既存データ反映シート 位置定数 ---
Private Const MG_ROW_FMT_ID As Long = 3
Private Const MG_COL_FMT_ID_VAL As Long = 3

' ================================================================
' 関数名: Btn_SaveKnowledge
' 概要:   ナレッジ登録「▶保存」ボタン
' ================================================================
Public Sub Btn_SaveKnowledge()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryKnowledge", "Btn_SaveKnowledge", _
                    "保存ボタン押下"
    
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    
    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()
    
    Dim savedNo As String
    savedNo = knwMgr.SaveNewKnowledge()
    
    If savedNo = "" Then
        Call ShowError("ナレッジ保存", "保存に失敗しました", _
                        "必須項目が入力されているか確認してください")
    Else
        Call ShowInfo("ナレッジ保存", _
                       "ナレッジ " & savedNo & " を保存しました")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ナレッジ保存", Err.Description, _
                    "入力内容を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_ClearForm
' 概要:   ナレッジ登録「▶クリア」ボタン
' ================================================================
Public Sub Btn_ClearForm()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_SAVE)
    
    If Not ConfirmAction("入力クリア", _
                           "入力中の値を全てクリアします") Then
        Exit Sub
    End If
    
    Dim i As Long
    For i = 4 To 1000
        If ws.Cells(i, 2).Value = "" Then Exit For
        ws.Cells(i, 3).Value = ""
    Next i
    Exit Sub

ErrHandler:
    Call ShowError("入力クリア", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_LoadKnowledge
' 概要:   ナレッジ修正「▶読込」ボタン
' ================================================================
Public Sub Btn_LoadKnowledge()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)
    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(KE_ROW_FMT_ID, KE_COL_KNW_NO).Value)
    
    If knowledgeNo = "" Then
        Call ShowWarning("ナレッジ読込", _
                         "ナレッジ番号が入力されていません", _
                         "上部の番号欄に入力してから読込ボタンを押してください")
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    
    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()
    
    If Not knwMgr.LoadForEdit(knowledgeNo) Then
        Call ShowError("ナレッジ読込", _
                        "指定されたナレッジが見つかりません", _
                        "ナレッジ番号を確認してから再度ボタンを押してください")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ナレッジ読込", Err.Description, _
                    "ナレッジ番号を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_UpdateKnowledge
' 概要:   ナレッジ修正「▶上書保存」ボタン
' ================================================================
Public Sub Btn_UpdateKnowledge()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)
    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(KE_ROW_FMT_ID, KE_COL_KNW_NO).Value)
    
    If knowledgeNo = "" Then
        Call ShowWarning("上書保存", _
                         "ナレッジ番号が入力されていません", _
                         "読込ボタンでナレッジを読み込んでから再度実行してください")
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    
    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()
    
    If knwMgr.UpdateKnowledge(knowledgeNo) Then
        Call ShowInfo("上書保存", _
                       "ナレッジ " & knowledgeNo & " を更新しました")
    Else
        Call ShowError("上書保存", "更新に失敗しました", _
                        "入力内容を確認してから再度ボタンを押してください")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("上書保存", Err.Description, _
                    "入力内容を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_ReloadList
' 概要:   ナレッジ一覧「▶リロード」ボタン
'         データフォルダ内の全ナレッジファイルを一覧表示
' ================================================================
Public Sub Btn_ReloadList()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryKnowledge", "Btn_ReloadList", _
                    "リロード開始"
    
    Call ReloadListCore(logger)
    Exit Sub

ErrHandler:
    Call ShowError("ナレッジ一覧リロード", Err.Description, _
                    "データフォルダパスを確認してから再度ボタンを押してください")
End Sub

' リロードの実装本体
Private Sub ReloadListCore(ByVal logger As clsLogger)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)
    
    Dim dataFolder As String
    dataFolder = GetDataFolder()
    
    ' 既存のリストをクリア
    Dim i As Long
    For i = KL_RESULT_START_ROW To KL_RESULT_START_ROW + 1000
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 6)).ClearContents
    Next i
    
    Dim files As Variant
    files = ListFilesInFolder(dataFolder, "txt")

    ' M-4 guard: 空配列なら早期 return (UBound エラー防止)

    If (Not Not files) = 0 Then Exit Sub
    
    Dim targetRow As Long
    targetRow = KL_RESULT_START_ROW
    
    Dim idx As Long
    For idx = LBound(files) To UBound(files)
        Dim fileName As String
        fileName = CStr(files(idx))
        
        Dim knwNo As String
        knwNo = Left(fileName, Len(fileName) - 4)
        
        ws.Cells(targetRow, 1).Value = idx + 1
        ws.Cells(targetRow, 2).Value = knwNo
        ws.Cells(targetRow, 3).Value = ""
        ws.Cells(targetRow, 4).Value = ""
        ws.Cells(targetRow, 5).Value = ""
        ws.Cells(targetRow, 6).Value = ""
        
        targetRow = targetRow + 1
    Next idx
    
    Dim count As Long
    count = UBound(files) - LBound(files) + 1
    
    logger.LogInfo "modEntryKnowledge", "Btn_ReloadList", _
                    "リロード完了: " & CStr(count) & "件"
End Sub

' ================================================================
' 関数名: Btn_DeleteKnowledge
' 概要:   ナレッジ一覧「▶選択行を削除」ボタン
' ================================================================
Public Sub Btn_DeleteKnowledge()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)
    
    Dim selRow As Long
    selRow = ws.Application.Selection.Row
    
    If selRow < KL_RESULT_START_ROW Then
        Call ShowWarning("ナレッジ削除", _
                         "削除したい行が選択されていません", _
                         "削除したい行を選択してから再度ボタンを押してください")
        Exit Sub
    End If
    
    Dim knowledgeNo As String
    knowledgeNo = CStr(ws.Cells(selRow, 2).Value)
    
    If knowledgeNo = "" Then
        Call ShowWarning("ナレッジ削除", _
                         "選択行にナレッジ番号がありません", _
                         "リロードしてから削除したい行を選択してください")
        Exit Sub
    End If
    
    If Not ConfirmAction("ナレッジ削除", _
                           "ナレッジ " & knowledgeNo & " を物理削除します") Then
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    
    Dim knwMgr As clsKnowledgeManager
    Set knwMgr = New clsKnowledgeManager
    knwMgr.Init logger, formatMgr, GetDataFolder()
    
    If knwMgr.DeleteKnowledge(knowledgeNo) Then
        Call ShowInfo("ナレッジ削除", _
                       "ナレッジ " & knowledgeNo & " を削除しました")
        Call ReloadListCore(logger)
    Else
        Call ShowError("ナレッジ削除", "削除に失敗しました", _
                        "ファイルパスを確認してから再度ボタンを押してください")
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ナレッジ削除", Err.Description, _
                    "選択行を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_MigrateFields
' 概要:   既存データ反映「▶反映実行」ボタン
' ================================================================
Public Sub Btn_MigrateFields()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MIGRATION)
    
    Dim formatId As String
    formatId = CStr(ws.Cells(MG_ROW_FMT_ID, MG_COL_FMT_ID_VAL).Value)
    
    If formatId = "" Then
        Call ShowWarning("フィールド反映", _
                         "フォーマットIDが選択されていません", _
                         "上部のプルダウンから対象フォーマットを選択してください")
        Exit Sub
    End If
    
    If Not ConfirmAction("フィールド反映", _
                           "フォーマット " & formatId & " の全ナレッジにフィールド定義を反映します") Then
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim formatMgr As clsFormatManager
    Set formatMgr = New clsFormatManager
    formatMgr.Init logger
    
    Dim migrator As clsFieldMigrator
    Set migrator = New clsFieldMigrator
    migrator.Init logger, formatMgr, GetDataFolder()
    
    Dim processedCount As Long
    processedCount = migrator.MigrateFields(formatId)

    ' rev20: M8-002 対応。clsFieldMigrator.MigrateFields は内部で
    ' "反映完了" を LogInfo するが、何らかの runtime エラーで
    ' ErrHandler に分岐すると "反映完了" は記録されず M8-002 が
    ' FAIL する。Btn 側でも改めて "反映完了" を残すことで
    ' CheckLogExists("反映完了") が安定して True になるようにする。
    logger.LogInfo "modEntryKnowledge", "Btn_MigrateFields", _
                    "反映完了: " & CStr(processedCount) & "件"

    Call ShowInfo("フィールド反映", _
                   CStr(processedCount) & " 件のナレッジに反映しました")
    Exit Sub

ErrHandler:
    Call ShowError("フィールド反映", Err.Description, _
                    "フォーマットIDとデータフォルダを確認してから再度実行してください")
End Sub
```

## 関連

- 呼び出す: `clsKnowledgeManager`, `clsLogger`, `modCommon`
- 呼び出される: `ボタン: Btn_RegisterKnw 等`
