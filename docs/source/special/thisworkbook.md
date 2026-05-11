---
title: ThisWorkbook.cls
---

# ThisWorkbook.cls

| 項目 | 値 |
|---|---|
| 層 | 特殊モジュール |
| 種別 | ドキュメントモジュール (ThisWorkbook) |
| 役割 | 本番版 ThisWorkbook (Workbook_Open / Workbook_BeforeClose イベント / メイン画面初期表示) |
| 行数 | 85 行 |

## 配置先

VBE のプロジェクトツリーで `ThisWorkbook` モジュールを開き、コードペインに **置換貼り付け** します。新規モジュールとしてインポートしないでください。コード冒頭の 9 行（`VERSION 1.0 CLASS` から始まるヘッダー）は **貼り付け対象外** です。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ================================================================
' クラス: ThisWorkbook（ドキュメントモジュール）
' 概要:   ブックイベントを処理
'         - Workbook_Open: ログクリア、起動ログ、初期表示
'         - Workbook_BeforeClose: 終了ログ
' 依存先: clsLogger, clsTaskController, modEntryMain, modCommon
' 備考:   v21 (E2E rerun) で起動/終了に詳細ログを注入。
'         外部ログファイル C:\kvba\runtime.log にもセッションマーカーを残す。
' ================================================================

Private Const MOD_NAME As String = "ThisWorkbook"

' ================================================================
' 関数名: Workbook_Open
' 概要:   ブックオープン時のイベント
'         (1) ログシートの自動クリア（セッション単位リセット）
'         (2) 起動ログ出力
'         (3) 初期タスクに設定（メインシートのみ表示）
' 引数:   なし
' 戻り値: なし
' ================================================================
Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"

    stepName = "BuildLogger"
    Dim logger As clsLogger
    Set logger = BuildLogger()
    If Not logger Is Nothing Then
        logger.LogInfo MOD_NAME, "Workbook_Open", "=== セッション開始 (book=" & ThisWorkbook.Name & ") ==="
        logger.LogTrace MOD_NAME, "Workbook_Open", "ENTER"
    End If

    stepName = "ClearLog"
    If Not logger Is Nothing Then logger.ClearLog

    ' クリア後の起動ログ
    stepName = "startup log"
    If Not logger Is Nothing Then
        logger.LogInfo MOD_NAME, "Workbook_Open", _
                        "システム起動 (book=" & ThisWorkbook.Name & ", sheets=" & ThisWorkbook.Worksheets.Count & ")"
    End If

    ' 初期表示: メインシートのみ表示
    stepName = "SetInitialVisibility"
    Call SetInitialVisibility

    ' メインシートをアクティブに
    stepName = "Activate メイン"
    ThisWorkbook.Worksheets(SHEET_MAIN).Activate

    If Not logger Is Nothing Then
        logger.LogTrace MOD_NAME, "Workbook_Open", "EXIT ok"
    End If
    Exit Sub

ErrHandler:
    ' 起動時エラーはメッセージボックスのみ（ログは使えない可能性あり）
    On Error Resume Next
    Dim recoveryLogger As clsLogger
    Set recoveryLogger = BuildLogger()
    If Not recoveryLogger Is Nothing Then
        recoveryLogger.LogErrorWithErr MOD_NAME, "Workbook_Open", stepName, Err.Number, Err.Description
    End If
    Err.Clear
    MsgBox "起動時にエラーが発生しました:" & vbCrLf & _
           "  step  : " & stepName & vbCrLf & _
           "  desc  : " & Err.Description & vbCrLf & vbCrLf & _
           "詳細は C:\kvba\runtime.log を確認してください。", _
           vbCritical
End Sub

' ================================================================
' 関数名: Workbook_BeforeCl
```
