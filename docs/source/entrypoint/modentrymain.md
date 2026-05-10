---
title: modEntryMain.bas
---

# modEntryMain.bas

| 項目 | 値 |
|---|---|
| 層 | エントリポイント層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | 本番モード起点 / 共通 BuildLogger 提供 |
| 行数 | 241 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modEntryMain` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modEntryMain"
Option Explicit

' ================================================================
' モジュール: modEntryMain（エントリポイント層）
' 概要:   メインシートの 12 タスク選択ボタン + 各画面の「メインに戻る」
'         ボタンに割り当てるマクロ群。
'         polished mock M-01 v19 準拠で 8 → 12 ボタン化。
' 依存先: clsLogger, clsTaskController, modCommon, modFactory
' ================================================================

' ================================================================
' --- 12 タスク切替ボタン (M-01 メイン) ---
' ================================================================

Public Sub Btn_TaskSearch()
    Call SwitchTaskCommon(TASK_SEARCH, "Btn_TaskSearch")
End Sub

Public Sub Btn_TaskRegister()
    Call SwitchTaskCommon(TASK_REGISTER, "Btn_TaskRegister")
End Sub

Public Sub Btn_TaskModify()
    Call SwitchTaskCommon(TASK_MODIFY, "Btn_TaskModify")
End Sub

Public Sub Btn_TaskList()
    Call SwitchTaskCommon(TASK_LIST, "Btn_TaskList")
End Sub

Public Sub Btn_TaskFormat()
    Call SwitchTaskCommon(TASK_FORMAT, "Btn_TaskFormat")
End Sub

Public Sub Btn_TaskFieldReflect()
    Call SwitchTaskCommon(TASK_FIELD_REFLECT, "Btn_TaskFieldReflect")
End Sub

Public Sub Btn_TaskStorage()
    Call SwitchTaskCommon(TASK_STORAGE, "Btn_TaskStorage")
End Sub

Public Sub Btn_TaskSysSettings()
    Call SwitchTaskCommon(TASK_SYS_SETTINGS, "Btn_TaskSysSettings")
End Sub

Public Sub Btn_TaskLog()
    Call SwitchTaskCommon(TASK_LOG, "Btn_TaskLog")
End Sub

Public Sub Btn_TaskFileFormat()
    Call SwitchTaskCommon(TASK_FILE_FORMAT, "Btn_TaskFileFormat")
End Sub

Public Sub Btn_TaskInitSetup()
    On Error GoTo ErrHandler
    Call modSetup.SetupSheetsAndButtons(False)
    Exit Sub
ErrHandler:
    Call ShowError("再セットアップ", Err.Description, "再度ボタンを押してください")
End Sub

Public Sub Btn_TaskHelpVersion()
    On Error GoTo ErrHandler
    MsgBox "ナレッジ管理システム v2.0" & vbCrLf & _
           "ビルド: 2026-05-10" & vbCrLf & _
           "ライセンス: 社内利用限定", _
           vbInformation, "ヘルプ / バージョン"
    Exit Sub
ErrHandler:
    Call ShowError("バージョン表示", Err.Description, "")
End Sub

' ================================================================
' --- 全画面共通: メインに戻る ボタン ---
' ================================================================

Public Sub Btn_BackToMain()
    On Error GoTo ErrHandler
    Dim w As Worksheet
    For Each w In ThisWorkbook.Worksheets
        If w.Name = SHEET_MAIN Then
            w.Visible = xlSheetVisible
        Else
            w.Visible = xlSheetHidden
        End If
    Next w
    ThisWorkbook.Worksheets(SHEET_MAIN).Activate
    ThisWorkbook.Worksheets(SHEET_MAIN).Range("D7").Value = "(未選択)"
    Exit Sub
ErrHandler:
    Call ShowError("メインに戻る", Err.Description, "再度ボタンを押してください")
End Sub

' ================================================================
' --- 後方互換ボタン（旧 8 タスク → 新 12 タスクへのリダイレクト） ---
' ================================================================

Public Sub Btn_TaskSetup()
    Call Btn_TaskInitSetup
End Sub

Public Sub Btn_TaskConfig()
    Call Btn_TaskSysSettings
End Sub

Public Sub Btn_TaskEdit()
    Call Btn_TaskModify
End Sub

Public Sub Btn_TaskDelete()
    Call Btn_TaskModify
End Sub

Public Sub Btn_TaskMigrate()
    Call Btn_TaskFieldReflect
End Sub

' ================================================================
' --- 共通処理 ---
' ================================================================

Private Sub SwitchTaskCommon(ByVal taskName As String, ByVal callerName As String)
    On Error GoTo ErrHandler
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryMain", callerName, _
                    "タスク切替開始: " & taskName

    Dim controller As clsTaskController
    Set controller = New clsTaskController
    controller.Init logger
    controller.SwitchToTask taskName
    Exit Sub
ErrHandler:
    Call ShowError("タスク切替", Err.Description, "再度ボタンを押してください")
End Sub

Public Function BuildLogger() As clsLogger
    Dim logger As clsLogger
    Set logger = New clsLogger
    logger.Init ThisWorkbook.Worksheets(SHEET_LOG), GetDebugLevel()
    Set BuildLogger = logger
End Function

Public Function GetDebugLevel() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    Dim lvl As String
    lvl = CStr(ws.Cells(SETTINGS_ROW_DEBUGLEVEL, SETTINGS_COL_VALUE).Value)
    If lvl <> DEBUG_ON Then lvl = DEBUG_OFF
    GetDebugLevel = lvl
    Exit Function
ErrHandler:
    GetDebugLevel = DEBUG_OFF
End Function

Public Function GetDataFolder() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    GetDataFolder = CStr(ws.Cells(SETTINGS_ROW_DATAFOLDER, SETTINGS_COL_VALUE).Value)
    Exit Function
ErrHandler:
    GetDataFolder = ""
End Function

Public Function IsTestMode() As Boolean
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    Dim v As String
    v = UCase(CStr(ws.Cells(SETTINGS_ROW_TESTMODE, SETTINGS_COL_VALUE).Value))
    IsTestMode = (v = TESTMODE_ON)
    Exit Function
ErrHandler:
    IsTestMode = False
End Function

Public Sub ShowError(ByVal operation As String, ByVal detail As String, ByVal action As String)
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowError", operation, detail, action)
        Exit Sub
    End If
    MsgBox "操作名: " & operation & vbCrLf & _
            "内容: " & detail & vbCrLf & _
            "対処: " & action, _
            vbCritical, "エラー"
End Sub

Public Sub ShowWarning(ByVal operation As String, ByVal detail As String, ByVal action As String)
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowWarning", operation, detail, action)
        Exit Sub
    End If
    MsgBox "操作名: " & operation & vbCrLf & _
            "内容: " & detail & vbCrLf & _
            "対処: " & action, _
            vbExclamation, "警告"
End Sub

Public Sub ShowInfo(ByVal operation As String, ByVal detail As String)
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowInfo", operation, detail, "")
        Exit Sub
    End If
    MsgBox "操作名: " & operation & vbCrLf & _
            "内容: " & detail, _
            vbInformation, "情報"
End Sub

Public Function ConfirmAction(ByVal operation As String, ByVal detail As String) As Boolean
    If IsTestMode() Then
        Call LogDialogSuppressed("ConfirmAction", operation, detail, "自動Yes")
        ConfirmAction = True
        Exit Function
    End If
    Dim result As VbMsgBoxResult
    result = MsgBox("操作名: " & operation & vbCrLf & _
                      "内容: " & detail & vbCrLf & _
                      "実行してもよろしいですか？", _
                      vbQuestion + vbYesNo, "確認")
    ConfirmAction = (result = vbYes)
End Function

Private Sub LogDialogSuppressed(ByVal dialogType As String, _
                                  ByVal operation As String, _
                                  ByVal detail As String, _
                                  ByVal action As String)
    On Error Resume Next
    Dim logger As clsLogger
    Set logger = BuildLogger()
    Dim msg As String
    msg = "[" & dialogType & "抑制] 操作=" & operation & ", 内容=" & detail
    If action <> "" Then msg = msg & ", 対処=" & action
    logger.LogInfo "modEntryMain", "LogDialogSuppressed", msg
    Err.Clear
End Sub
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
```
