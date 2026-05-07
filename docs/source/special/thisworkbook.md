---
title: ThisWorkbook.cls
---

# ThisWorkbook.cls

| 項目 | 値 |
|---|---|
| 層 | 特殊モジュール |
| 種別 | ドキュメントモジュール (ThisWorkbook) |
| 役割 | 本番版 ThisWorkbook (Workbook_Open 等のイベント / 自動初期化なし) |
| 行数 | 93 行 |

## 配置先

VBE のプロジェクトツリーで `ThisWorkbook` モジュールを開き、コードペインに**置換貼り付け**します。新規モジュールとしてインポートしないでください。

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
' ================================================================

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
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    
    ' ログシート自動クリア（セッション単位リセット）
    logger.ClearLog
    
    ' 起動ログ
    logger.LogInfo "ThisWorkbook", "Workbook_Open", "システム起動"
    
    ' 初期表示: メインシートのみ表示
    Call SetInitialVisibility
    
    ' メインシートをアクティブに
    ThisWorkbook.Worksheets(SHEET_MAIN).Activate
    Exit Sub

ErrHandler:
    ' 起動時エラーはメッセージボックスのみ（ログは使えない可能性あり）
    MsgBox "起動時にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' ================================================================
' 関数名: Workbook_BeforeClose
' 概要:   ブックを閉じる直前のイベント
' 引数:   Cancel - キャンセルフラグ
' 戻り値: なし
' ================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo ErrHandler
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "ThisWorkbook", "Workbook_BeforeClose", "システム終了"
    Exit Sub

ErrHandler:
    ' 終了時エラーは握りつぶす（ブックを閉じる妨げにしない）
End Sub

' ================================================================
' 関数名: SetInitialVisibility
' 概要:   初期表示時のシート可視性を設定
'         メインシート以外を非表示にする
' 引数:   なし
' 戻り値: なし
' ================================================================
Private Sub SetInitialVisibility()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_MAIN Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetHidden
        End If
    Next ws
    Exit Sub

ErrHandler:
    ' エラー時は何もしない
End Sub
```

## 関連

- 呼び出す: `modEntryMain`
- 呼び出される: `Excel イベントから直接呼び出し`
