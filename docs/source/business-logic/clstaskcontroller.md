---
title: clsTaskController.cls
---

# clsTaskController.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | ナレッジ操作 (登録 / 編集) のトランザクション制御 |
| 行数 | 187 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsTaskController` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTaskController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Phase 6 レビュー: クラス内状態は transient (シート切替の一時記録のみ)。
' D-3 余地: 現状 ThisWorkbook 直結なし、副作用は呼出側 ws を経由。テスト容易性 OK。

' ================================================================
' クラス: clsTaskController（ビジネスロジック層）
' 概要:   利用シーン別シート表示・非表示を制御する
'         8タスクに応じて関連シートのみを表示し、他を非表示にする
'         メインシートは常に表示
' 依存先: clsLogger, modCommon
' ================================================================

' メインシートの現在タスク表示位置
Private Const CURRENT_TASK_ROW As Long = 3
Private Const CURRENT_TASK_COL As Long = 3

Private m_logger As clsLogger

' ================================================================
' 関数名: Init
' 概要:   初期化
' 引数:   logger - ログ出力用
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logger As clsLogger)
    Set m_logger = logger
End Sub

' ================================================================
' 関数名: SwitchToTask
' 概要:   指定タスクに切り替える（関連シートのみ表示、他は非表示）
'         メインシートの現在タスク表示も更新
' 引数:   taskName - タスク名（例: "ナレッジ登録"）
' 戻り値: なし
' 備考:   未定義のタスク名の場合はメインシートのみ表示
' ================================================================
Public Sub SwitchToTask(ByVal taskName As String)
    On Error GoTo ErrHandler
    
    Dim sheetNames As Variant
    sheetNames = GetTaskDefinition(taskName)
    
    Call SetVisibleSheets(sheetNames)
    Call UpdateCurrentTaskLabel(taskName)
    
    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsTaskController", "SwitchToTask", _
                          "タスク切替: " & taskName
    End If
    Exit Sub

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsTaskController", "SwitchToTask", _
                           "切替失敗: " & Err.Description
    End If
End Sub

' ================================================================
' 関数名: GetCurrentTask
' 概要:   メインシートに表示されている現在のタスク名を取得
' 引数:   なし
' 戻り値: String - 現在のタスク名（未選択なら空文字列）
' ================================================================
Public Function GetCurrentTask() As String
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    GetCurrentTask = CStr(ws.Cells(CURRENT_TASK_ROW, CURRENT_TASK_COL).Value)
    Exit Function

ErrHandler:
    GetCurrentTask = ""
End Function

' ================================================================
' 関数名: SetVisibleSheets
' 概要:   指定されたシート群のみを表示、他は非表示にする
' 引数:   sheetNames - 表示するシート名の配列
' 戻り値: なし
' 備考:   xlSheetHidden を使用（ユーザが右クリック→再表示で戻せる）
'         xlVeryHidden は使わない
' ================================================================
Private Sub SetVisibleSheets(ByVal sheetNames As Variant)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim shouldShow As Boolean
    
    For Each ws In ThisWorkbook.Worksheets
        shouldShow = IsInArray(ws.Name, sheetNames)
        
        If shouldShow Then
            ws.Visible = xlSheetVisible
        Else
            ' ログシート・データファイル形式・メイン以外を非表示化
            ' ただし全てのシートを非表示にはできないため、最低1枚は残す
            If ws.Name <> SHEET_MAIN Then
                ws.Visible = xlSheetHidden
            End If
        End If
    Next ws
    Exit Sub

ErrHandler:
    ' エラー時は何もしない（切替失敗はログで記録済み）
End Sub

' ================================================================
' 関数名: UpdateCurrentTaskLabel
' 概要:   メインシートの現在タスク表示欄を更新
' 引数:   taskName - 表示するタスク名
' 戻り値: なし
' ================================================================
Private Sub UpdateCurrentTaskLabel(ByVal taskName As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    ws.Cells(CURRENT_TASK_ROW, CURRENT_TASK_COL).NumberFormat = "@"
    ws.Cells(CURRENT_TASK_ROW, CURRENT_TASK_COL).Value = taskName
    Exit Sub

ErrHandler:
    ' エラー時は何もしない
End Sub

' ================================================================
' 関数名: GetTaskDefinition
' 概要:   タスク名に対応する表示シート名の配列を返す
' 引数:   taskName - タスク名
' 戻り値: Variant - シート名の配列
' 備考:   未定義のタスク名の場合はメインシートのみの配列を返す
' ================================================================
Private Function GetTaskDefinition(ByVal taskName As String) As Variant
    Select Case taskName
        Case TASK_SETUP
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_SETTINGS, SHEET_STORAGE)
        Case TASK_CONFIG
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_SETTINGS, SHEET_STORAGE)
        Case TASK_FORMAT
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_FORMAT_LIST, _
                                       SHEET_FORMAT_DESIGN, SHEET_FORMAT_PREVIEW)
        Case TASK_REGISTER
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_KNW_SAVE)
        Case TASK_SEARCH
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_SEARCH, SHEET_KNW_DISPLAY)
        Case TASK_EDIT
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_KNW_EDIT, _
                                       SHEET_SEARCH, SHEET_KNW_DISPLAY)
        Case TASK_DELETE
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_KNW_LIST)
        Case TASK_MIGRATE
            GetTaskDefinition = Array(SHEET_MAIN, SHEET_MIGRATION)
        Case Else
            GetTaskDefinition = Array(SHEET_MAIN)
    End Select
End Function

' ================================================================
' 関数名: IsInArray
' 概要:   配列内に指定された値が含まれるか判定
' 引数:   target - 検索対象
'         arr    - 配列
' 戻り値: Boolean - 含まれるなら True
' ================================================================
Private Function IsInArray(ByVal target As String, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If CStr(arr(i)) = target Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function
```

## 関連

- 呼び出す: `clsKnowledgeManager`, `clsFormatManager`, `clsLogger`
- 呼び出される: `modEntryKnowledge`
