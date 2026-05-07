---
title: modEntryMain.bas
---

# modEntryMain.bas

| 項目 | 値 |
|---|---|
| 層 | エントリポイント層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | 本番モード起点 / 共通 BuildLogger 提供 |
| 行数 | 315 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modEntryMain` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modEntryMain"
Option Explicit

' ================================================================
' モジュール: modEntryMain（エントリポイント層）
' 概要:       メインシートの8つのタスク選択ボタンに割り当てるマクロ群
' 依存先:     clsLogger, clsTaskController, modCommon
' ================================================================

' ================================================================
' 関数名: Btn_TaskSetup
' 概要:   タスク「初回セットアップ」に切替
' 引数:   なし
' 戻り値: なし
' ================================================================
Public Sub Btn_TaskSetup()
    On Error GoTo ErrHandler
    Call SwitchTaskCommon(TASK_SETUP, "Btn_TaskSetup")
    Exit Sub
ErrHandler:
    Call ShowError("タスク切替", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_TaskConfig
' 概要:   タスク「設定変更」に切替
' ================================================================
Public Sub Btn_TaskConfig()
    On Error GoTo ErrHandler
    Call SwitchTaskCommon(TASK_CONFIG, "Btn_TaskConfig")
    Exit Sub
ErrHandler:
    Call ShowError("タスク切替", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_TaskFormat
' 概要:   タスク「フォーマット管理」に切替
' ================================================================
Public Sub Btn_TaskFormat()
    On Error GoTo ErrHandler
    Call SwitchTaskCommon(TASK_FORMAT, "Btn_TaskFormat")
    Exit Sub
ErrHandler:
    Call ShowError("タスク切替", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_TaskRegister
' 概要:   タスク「ナレッジ登録」に切替
' ================================================================
Public Sub Btn_TaskRegister()
    On Error GoTo ErrHandler
    Call SwitchTaskCommon(TASK_REGISTER, "Btn_TaskRegister")
    Exit Sub
ErrHandler:
    Call ShowError("タスク切替", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_TaskSearch
' 概要:   タスク「検索・確認」に切替
' ================================================================
Public Sub Btn_TaskSearch()
    On Error GoTo ErrHandler
    Call SwitchTaskCommon(TASK_SEARCH, "Btn_TaskSearch")
    Exit Sub
ErrHandler:
    Call ShowError("タスク切替", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_TaskEdit
' 概要:   タスク「ナレッジ修正」に切替
' ================================================================
Public Sub Btn_TaskEdit()
    On Error GoTo ErrHandler
    Call SwitchTaskCommon(TASK_EDIT, "Btn_TaskEdit")
    Exit Sub
ErrHandler:
    Call ShowError("タスク切替", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_TaskDelete
' 概要:   タスク「ナレッジ削除」に切替
' ================================================================
Public Sub Btn_TaskDelete()
    On Error GoTo ErrHandler
    Call SwitchTaskCommon(TASK_DELETE, "Btn_TaskDelete")
    Exit Sub
ErrHandler:
    Call ShowError("タスク切替", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: Btn_TaskMigrate
' 概要:   タスク「既存データ反映」に切替
' ================================================================
Public Sub Btn_TaskMigrate()
    On Error GoTo ErrHandler
    Call SwitchTaskCommon(TASK_MIGRATE, "Btn_TaskMigrate")
    Exit Sub
ErrHandler:
    Call ShowError("タスク切替", Err.Description, _
                    "再度ボタンを押してください")
End Sub

' ================================================================
' 関数名: SwitchTaskCommon
' 概要:   タスク切替の共通処理（Logger生成、Controller実行）
' 引数:   taskName    - 切替先タスク名
'         callerName  - 呼び出し元関数名（ログ用）
' 戻り値: なし
' ================================================================
Private Sub SwitchTaskCommon(ByVal taskName As String, _
                               ByVal callerName As String)
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "modEntryMain", callerName, _
                    "タスク切替開始: " & taskName
    
    Dim controller As clsTaskController
    Set controller = New clsTaskController
    controller.Init logger
    controller.SwitchToTask taskName
End Sub

' ================================================================
' 関数名: BuildLogger
' 概要:   設定シートの値を読み取ってLoggerインスタンスを生成
' 引数:   なし
' 戻り値: clsLogger - 初期化済みLogger
' ================================================================
Public Function BuildLogger() As clsLogger
    Dim logger As clsLogger
    Set logger = New clsLogger
    logger.Init ThisWorkbook.Worksheets(SHEET_LOG), _
                 GetDebugLevel()
    Set BuildLogger = logger
End Function

' ================================================================
' 関数名: GetDebugLevel
' 概要:   設定シートからデバッグレベルの値を取得
' 引数:   なし
' 戻り値: String - "OFF" または "ON"
' ================================================================
Public Function GetDebugLevel() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    Dim lvl As String
    lvl = CStr(ws.Cells(SETTINGS_ROW_DEBUGLEVEL, SETTINGS_COL_VALUE).Value)
    If lvl <> DEBUG_ON Then
        lvl = DEBUG_OFF
    End If
    GetDebugLevel = lvl
End Function

' ================================================================
' 関数名: GetDataFolder
' 概要:   設定シートからデータフォルダパスを取得
' 引数:   なし
' 戻り値: String - データフォルダのUNCパス
' ================================================================
Public Function GetDataFolder() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    GetDataFolder = CStr(ws.Cells(SETTINGS_ROW_DATAFOLDER, SETTINGS_COL_VALUE).Value)
End Function

' ================================================================
' 関数名: IsTestMode
' 概要:   テストモードが有効か判定
'         有効時は MsgBox 系関数がダイアログを出さずログのみ出力
' 引数:   なし
' 戻り値: Boolean - テストモード有効なら True
' 備考:   設定シート 行6 C列 = "TRUE" でテストモード有効
'         通常運用時は空または"FALSE"
' ================================================================
Public Function IsTestMode() As Boolean
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    Dim v As String
    v = UCase(CStr(ws.Cells(6, SETTINGS_COL_VALUE).Value))
    IsTestMode = (v = "TRUE")
    Exit Function
ErrHandler:
    IsTestMode = False
End Function

' ================================================================
' 関数名: ShowError
' 概要:   統一フォーマットのエラーメッセージ表示
'         ユーザビリティ点検の積み残し#6「エラーメッセージ統一」に対応
'         テストモード時はログのみ出力してダイアログは抑制
' 引数:   operation - 操作名
'         detail    - エラー内容
'         action    - 推奨アクション
' 戻り値: なし
' ================================================================
Public Sub ShowError(ByVal operation As String, _
                       ByVal detail As String, _
                       ByVal action As String)
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowError", operation, detail, action)
        Exit Sub
    End If
    MsgBox "操作名: " & operation & vbCrLf & _
            "内容: " & detail & vbCrLf & _
            "対処: " & action, _
            vbCritical, "エラー"
End Sub

' ================================================================
' 関数名: ShowWarning
' 概要:   統一フォーマットの警告メッセージ表示
'         テストモード時はログのみ出力してダイアログは抑制
' 引数:   operation - 操作名
'         detail    - 警告内容
'         action    - 推奨アクション
' 戻り値: なし
' ================================================================
Public Sub ShowWarning(ByVal operation As String, _
                        ByVal detail As String, _
                        ByVal action As String)
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowWarning", operation, detail, action)
        Exit Sub
    End If
    MsgBox "操作名: " & operation & vbCrLf & _
            "内容: " & detail & vbCrLf & _
            "対処: " & action, _
            vbExclamation, "警告"
End Sub

' ================================================================
' 関数名: ShowInfo
' 概要:   統一フォーマットの情報メッセージ表示
'         テストモード時はログのみ出力してダイアログは抑制
' 引数:   operation - 操作名
'         detail    - 情報内容
' 戻り値: なし
' ================================================================
Public Sub ShowInfo(ByVal operation As String, ByVal detail As String)
    If IsTestMode() Then
        Call LogDialogSuppressed("ShowInfo", operation, detail, "")
        Exit Sub
    End If
    MsgBox "操作名: " & operation & vbCrLf & _
            "内容: " & detail, _
            vbInformation, "情報"
End Sub

' ================================================================
' 関数名: ConfirmAction
' 概要:   ユーザーに確認を求めるダイアログ（Yes/No）
'         テストモード時は無条件で True を返し、ログに記録
' 引数:   operation - 操作名
'         detail    - 確認内容
' 戻り値: Boolean - Yesなら True（テストモード時は常に True）
' ================================================================
Public Function ConfirmAction(ByVal operation As String, _
                                ByVal detail As String) As Boolean
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

' ================================================================
' 関数名: LogDialogSuppressed
' 概要:   テストモード時に抑制されたダイアログ内容をログに記録
' 引数:   dialogType - ダイアログ種別（ShowError等）
'         operation  - 操作名
'         detail     - 内容
'         action     - 推奨アクション（Infoの場合は空）
' 戻り値: なし
' ================================================================
Private Sub LogDialogSuppressed(ByVal dialogType As String, _
                                  ByVal operation As String, _
                                  ByVal detail As String, _
                                  ByVal action As String)
    On Error GoTo ErrHandler
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    Dim msg As String
    msg = "[" & dialogType & "抑制] 操作=" & operation & _
           ", 内容=" & detail
    If action <> "" Then
        msg = msg & ", 対処=" & action
    End If
    
    logger.LogInfo "modEntryMain", "LogDialogSuppressed", msg
    Exit Sub

ErrHandler:
    ' ログ失敗時は握りつぶす
End Sub
```

## 関連

- 呼び出す: `modCommon`, `clsLogger`
- 呼び出される: `ThisWorkbook (本番)`
