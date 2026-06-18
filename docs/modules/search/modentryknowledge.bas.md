---
title: modEntryKnowledge.bas
description: modEntryKnowledge.bas のソースコード（コピペ用）
---

# modEntryKnowledge.bas

**配置先**: `検索.xlsm` 用の VBA モジュール
**種類**: 標準モジュール
**更新日**: 2026-06-09 22:15 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\search\`
- ファイル名: `modEntryKnowledge.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
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
' BUG-006 fix: ナレッジ登録 (M-05) は フォーマット選択 UI を行 6 に配置。
'              clsKnowledgeManager の KS_ROW_FMT_ID/KS_COL_FMT_ID_VAL と整合。
Public Const KS_ROW_FMT_ID As Long = 6
Public Const KS_COL_FMT_ID_VAL As Long = 3
Public Const KS_FORM_START_ROW As Long = 8
Public Const KS_FIELD_COL_NAME As Long = 3
Public Const KS_FIELD_COL_VALUE As Long = 5
Private Const KE_ROW_FMT_ID As Long = 1
Private Const KE_COL_KNW_NO As Long = 3

' --- ナレッジ一覧シート 位置定数 ---
Private Const KL_RESULT_START_ROW As Long = 4

' --- 既存データ反映シート 位置定数 ---
Private Const MG_ROW_FMT_ID As Long = 3
Private Const MG_COL_FMT_ID_VAL As Long = 3

' ================================================================
' 関数名: Btn_SaveKnowledge
' 概要:   ナレッジ登録「?保存」ボタン
' ================================================================
Public Sub Btn_SaveKnowledge()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryKnowledge", "Btn_SaveKnowledge", "保存ボタン押下", , "LOG-M05-SAVEKNOWLEDGE-ENTRY"
    
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
' 概要:   ナレッジ登録「?クリア」ボタン
' 備考:   BUG-006 fix - 新 5 列レイアウト (No / 必須 / ラベル / 型 / 入力欄)
'         に対応。空セル判定はラベル列 (KS_FIELD_COL_NAME) で行い、
'         入力欄 (KS_FIELD_COL_VALUE = E 列) のみクリアする
'         (ラベルや型は保持して再入力可とする)。
' ================================================================
Public Sub Btn_ClearForm()
    On Error GoTo ErrHandler

    ' BUG-C0-003 修正: LOG-M05-CLEARFORM-{ENTRY,EXIT-OK} を emit
    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_ClearForm", _
                       "ENTRY", , LOG_M05_CLEARFORM_ENTRY
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_SAVE)

    If Not ConfirmAction("入力クリア", _
                           "入力中の値を全てクリアします") Then
        Exit Sub
    End If

    Dim i As Long
    For i = KS_FORM_START_ROW To KS_FORM_START_ROW + 100
        If ws.Cells(i, KS_FIELD_COL_NAME).Value = "" Then Exit For
        ws.Cells(i, KS_FIELD_COL_VALUE).Value = ""
    Next i

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_ClearForm", _
                       "EXIT-OK", , LOG_M05_CLEARFORM_EXIT_OK
    End If
    Exit Sub

ErrHandler:
    Call ShowError("入力クリア", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' BUG-006: M-05 ロードボタン (modEntryFormat.Btn_LoadFormat と名前衝突回避)
Public Sub Btn_LoadKnowledgeFormat()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_SAVE)
    Dim formatId As String
    formatId = Trim(CStr(ws.Cells(KS_ROW_FMT_ID, KS_COL_FMT_ID_VAL).Value))
    If formatId = "" Then
        Call ShowWarning("フォーマットロード", "フォーマットID未入力", "C6にIDを入力")
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
    knwMgr.BuildRegistrationForm formatId
    Call ShowInfo("フォーマットロード", "フォーマット " & formatId & " をロードしました")
    Exit Sub
ErrHandler:
    Call ShowError("フォーマットロード", Err.Description, "IDを確認")
End Sub

' ================================================================
' 関数名: Btn_LoadKnowledge
' 概要:   ナレッジ修正「?読込」ボタン
' ================================================================
Public Sub Btn_LoadKnowledge()
    On Error GoTo ErrHandler

    ' BUG-C0-003 修正: LOG-M06-LOADKNOWLEDGE-{ENTRY,EXIT-OK} を emit
    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_LoadKnowledge", _
                       "ENTRY", , LOG_M06_LOADKNOWLEDGE_ENTRY
    End If

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

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_LoadKnowledge", _
                       "EXIT-OK", , LOG_M06_LOADKNOWLEDGE_EXIT_OK
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ナレッジ読込", Err.Description, _
                    "ナレッジ番号を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_UpdateKnowledge
' 概要:   ナレッジ修正「?上書保存」ボタン
' ================================================================
Public Sub Btn_UpdateKnowledge()
    On Error GoTo ErrHandler

    ' BUG-C0-003 修正: LOG-M06-UPDATEKNOWLEDGE-{ENTRY,EXIT-OK} を emit
    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_UpdateKnowledge", _
                       "ENTRY", , LOG_M06_UPDATEKNOWLEDGE_ENTRY
    End If

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

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_UpdateKnowledge", _
                       "EXIT-OK", , LOG_M06_UPDATEKNOWLEDGE_EXIT_OK
    End If
    Exit Sub

ErrHandler:
    Call ShowError("上書保存", Err.Description, _
                    "入力内容を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_ReloadList
' 概要:   ナレッジ一覧「?リロード」ボタン
'         データフォルダ内の全ナレッジファイルを一覧表示
' ================================================================
Public Sub Btn_ReloadList()
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryKnowledge", "Btn_ReloadList", "リロード開始", , "LOG-M07-RELOADLIST-ENTRY"
    
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
    
    logger.LogInfo "modEntryKnowledge", "Btn_ReloadList", "リロード完了: " & CStr(count) & "件", , "LOG-M07-RELOADLIST-EXIT-OK"
End Sub

' ================================================================
' 関数名: Btn_DeleteKnowledge
' 概要:   ナレッジ一覧「?選択行を削除」ボタン
' ================================================================
Public Sub Btn_DeleteKnowledge()
    On Error GoTo ErrHandler

    ' BUG-C0-003 修正: LOG-M06-DELETEKNOWLEDGE-{ENTRY,EXIT-OK} を emit
    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_DeleteKnowledge", _
                       "ENTRY", , LOG_M06_DELETEKNOWLEDGE_ENTRY
    End If

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

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_DeleteKnowledge", _
                       "EXIT-OK", , LOG_M06_DELETEKNOWLEDGE_EXIT_OK
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ナレッジ削除", Err.Description, _
                    "選択行を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_MigrateFields
' 概要:   既存データ反映「?反映実行」ボタン
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
    logger.LogInfo "modEntryKnowledge", "Btn_MigrateFields", "反映完了: " & CStr(processedCount) & "件", , "LOG-M12-MIGRATE-EXIT-OK"

    Call ShowInfo("フィールド反映", _
                   CStr(processedCount) & " 件のナレッジに反映しました")
    Exit Sub

ErrHandler:
    Call ShowError("フィールド反映", Err.Description, _
                    "フォーマットIDとデータフォルダを確認してから再度実行してください")
End Sub

' ================================================================
' 関数名: Btn_RestoreBackup
' 概要:   既存データ反映「バックアップから復旧」ボタン (M-12 F20)
' 仕様:   BUG-004-impl (Sprint 0 P0, ADR-0045 SSOT)
' ================================================================
Public Sub Btn_RestoreBackup()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_RestoreBackup", "ENTRY", , "LOG-M12-RESTOREBACKUP-ENTRY"
    End If

    Dim dataFolder As String
    dataFolder = GetDataFolder()

    If LenB(dataFolder) = 0 Then
        Call ShowWarning("バックアップ復旧", _
                         "データフォルダが未設定です", _
                         "[設定] 画面でデータフォルダを設定してください")
        Exit Sub
    End If

    Dim backupFolder As String
    ' iter23 (2026-06-01): modCommon の BACKUP_SUBFOLDER 解決が canonical コメント mojibake
    ' で時々失敗するため、ASCII literal "backup" を inline 化して module 間依存を切る。
    backupFolder = dataFolder & "\backup"

    Dim folderExists As Boolean
    folderExists = (Dir(backupFolder, vbDirectory) <> "")

    Dim hasFiles As Boolean
    hasFiles = False
    If folderExists Then
        Dim probe As String
        probe = Dir(backupFolder & "\*.*")
        Do While LenB(probe) > 0
            If probe <> "." And probe <> ".." Then
                hasFiles = True
                Exit Do
            End If
            probe = Dir()
        Loop
    End If

    If Not folderExists Or Not hasFiles Then
        Call ShowError("バックアップ復旧", _
                       "バックアップが見つかりません", _
                       "バックアップフォルダにファイルを配置してから再度ボタンを押してください")
        If Not logger Is Nothing Then
            logger.LogInfo "modEntryKnowledge", "Btn_RestoreBackup", _
                           "EXIT-OK no backup", , _
                           "LOG-M12-RESTOREBACKUP-NOBACKUP"
        End If
        Exit Sub
    End If

    Call ShowInfo("バックアップ復旧", _
                  "復元処理は未実装です" & vbCrLf & vbCrLf & _
                  "(検出: " & backupFolder & ")")
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_RestoreBackup", _
                       "EXIT-OK placeholder", , _
                       "LOG-M12-RESTOREBACKUP-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("バックアップ復旧", _
                   Err.Description, _
                   "バックアップフォルダのアクセス権限を確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_PageFirst
' 概要:   ナレッジ一覧「<<最初」ボタン (M-07 B12)
' 備考:   ページング: 表示行範囲を先頭リセット (一覧シートの A1 にページ番号を保持)
' ================================================================
Public Sub Btn_PageFirst()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PageFirst", "ENTRY", , "LOG-M07-PAGEFIRST-ENTRY"
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)
    Const PAGE_CELL As String = "A1"
    ws.Range(PAGE_CELL).Value = 1

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PageFirst", "EXIT-OK page=1", , "LOG-M07-PAGEFIRST-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ページング先頭", Err.Description, "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_PagePrev
' 概要:   ナレッジ一覧「<前」ボタン (M-07 D12)
' ================================================================
Public Sub Btn_PagePrev()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PagePrev", "ENTRY", , "LOG-M07-PAGEPREV-ENTRY"
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)
    Const PAGE_CELL As String = "A1"
    Dim current As Long
    current = CLng(Val(CStr(ws.Range(PAGE_CELL).Value)))
    If current <= 1 Then current = 1 Else current = current - 1
    ws.Range(PAGE_CELL).Value = current

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PagePrev", "EXIT-OK page=" & current, , "LOG-M07-PAGEPREV-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ページング前", Err.Description, "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_PageNext
' 概要:   ナレッジ一覧「次>」ボタン (M-07 G12)
' ================================================================
Public Sub Btn_PageNext()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PageNext", "ENTRY", , "LOG-M07-PAGENEXT-ENTRY"
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_LIST)
    Const PAGE_CELL As String = "A1"
    Dim current As Long
    current = CLng(Val(CStr(ws.Range(PAGE_CELL).Value)))
    If current < 1 Then current = 1
    current = current + 1
    ws.Range(PAGE_CELL).Value = current

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PageNext", "EXIT-OK page=" & current, , "LOG-M07-PAGENEXT-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ページング次", Err.Description, "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_PageLast
' 概要:   ナレッジ一覧「最後>>」ボタン (M-07 I12)
' 備考:   Stub: 全レコード数からの末尾ページ算出は次回 Sprint
' ================================================================
Public Sub Btn_PageLast()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PageLast", "ENTRY", , "LOG-M07-PAGELAST-ENTRY"
    End If

    Call ShowInfo("ページング末尾", _
                  "末尾ページへのジャンプは次回 Sprint で実装予定です")

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_PageLast", "EXIT-OK stub", , "LOG-M07-PAGELAST-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("ページング末尾", Err.Description, "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_ConfirmDiff
' 概要:   既存データ反映「差分確認」ボタン (M-12 I3)
' 引数:   なし (Excel フォームコントロールから直接呼び出される)
' 戻り値: なし
' 備考:   フォーマットIDの入力チェック + clsFieldMigrator 経由で
'         対象ナレッジ件数を取得し、ユーザに差分サマリを提示する。
'         完全な差分テーブル描画 (現状値 vs 変更後) は次回 Sprint で実装予定。
'         BUG-005 対応 (2026-05-15): spec 側 AddBtn (modScreenSpecRegistry.bas:486)
'         に対する Sub 実装が欠落していたため新規実装。
' ================================================================
Public Sub Btn_ConfirmDiff()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_ConfirmDiff", "ENTRY", , "LOG-M12-CONFIRMDIFF-ENTRY"
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MIGRATION)

    Dim formatId As String
    formatId = Trim(CStr(ws.Cells(MG_ROW_FMT_ID, MG_COL_FMT_ID_VAL).Value))

    If LenB(formatId) = 0 Then
        Call ShowWarning("差分確認", _
                         "フォーマットIDが選択されていません", _
                         "上部のプルダウンから対象フォーマットを選択してから再度ボタンを押してください")
        If Not logger Is Nothing Then
            logger.LogInfo "modEntryKnowledge", "Btn_ConfirmDiff", _
                           "EXIT-OK no formatId", , "LOG-M12-CONFIRMDIFF-NOFMT"
        End If
        Exit Sub
    End If

    Call ShowInfo("差分確認", _
                  "フォーマット " & formatId & " の差分プレビュー機能は次回 Sprint で実装予定です" & vbCrLf & _
                  "現状は [反映実行] ボタンで一括反映してください")

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_ConfirmDiff", _
                       "EXIT-OK stub formatId=" & formatId, , "LOG-M12-CONFIRMDIFF-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("差分確認", Err.Description, _
                    "フォーマットIDを確認してから再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_CancelMigrate
' 概要:   既存データ反映「中断」ボタン (M-12 D20)
' 引数:   なし (Excel フォームコントロールから直接呼び出される)
' 戻り値: なし
' 備考:   反映実行中の取消 (実装本体は次 Sprint) + メイン画面へ戻る。
'         確認ダイアログでユーザの中断意図を明示的に取得する。
'         BUG-005 対応 (2026-05-15): spec 側 AddBtn (modScreenSpecRegistry.bas:488)
'         に対する Sub 実装が欠落していたため新規実装。
' ================================================================
Public Sub Btn_CancelMigrate()
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_CancelMigrate", "ENTRY", , "LOG-M12-CANCELMIGRATE-ENTRY"
    End If

    If Not ConfirmAction("フィールド反映 中断", _
                          "現在のフィールド反映を中断してメイン画面へ戻ります") Then
        If Not logger Is Nothing Then
            logger.LogInfo "modEntryKnowledge", "Btn_CancelMigrate", _
                           "EXIT-OK user declined", , "LOG-M12-CANCELMIGRATE-DECLINED"
        End If
        Exit Sub
    End If

    ThisWorkbook.Worksheets(SHEET_MAIN).Activate

    If Not logger Is Nothing Then
        logger.LogInfo "modEntryKnowledge", "Btn_CancelMigrate", _
                       "EXIT-OK cancelled", , "LOG-M12-CANCELMIGRATE-EXIT-OK"
    End If
    Exit Sub

ErrHandler:
    Call ShowError("フィールド反映 中断", Err.Description, _
                    "再度ボタンを押してください")
End Sub
```
