---
title: clsSetupOrchestrator.cls
---

# clsSetupOrchestrator.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | SetupSheetsAndButtons の実装本体 (14 シート + 29 ボタン生成) |
| 行数 | 196 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsSetupOrchestrator` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSetupOrchestrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsSetupOrchestrator（インストーラ層）
' 概要:   初期セットアップの司令塔。14 シート作成 → 各画面の Setup 呼出 →
'         初期可視性設定 → 既定 Sheet 削除 を順次実行。
'         旧 modSetup.SetupSheetsAndButtons の責務を OOP で再構成したもの。
' 依存先: IScreenRenderer, modFactory, modScreenSpecRegistry, modCommon
' ================================================================

Private m_renderer As IScreenRenderer

Private Const REQUIRED_SHEETS_CSV As String = _
    "メイン,フォーマット一覧,フォーマット設計,フォーマットプレビュー," & _
    "ナレッジ登録,ナレッジ修正,ナレッジ一覧,検索,ナレッジ表示," & _
    "格納先設定,設定,既存データへのフィールド反映,データファイル形式,ログ"
Private Const DEFAULT_SHEETS_CSV As String = "Sheet1,Sheet2,Sheet3,Sheet4"

' ================================================================
' 関数名: Init
' 概要:   Renderer を依存注入
' ================================================================
Public Sub Init(ByVal renderer As IScreenRenderer)
    Set m_renderer = renderer
End Sub

' ================================================================
' 関数名: RunFullSetup
' 概要:   フルセットアップを実行
' ================================================================
Public Sub RunFullSetup()
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Call CreateRequiredSheets
    Call SetupAllScreens
    Call InitializeLogSheet
    Call InitializeSettingsSheet
    Call SetInitialVisibility
    Call DeleteEmptyDefaultSheets

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Err.Raise Err.Number, "clsSetupOrchestrator.RunFullSetup", Err.Description
End Sub

' ================================================================
' 関数名: CreateRequiredSheets
' 概要:   14 シートを順次作成（既存はスキップ）
' ================================================================
Private Sub CreateRequiredSheets()
    Dim names() As String
    names = Split(REQUIRED_SHEETS_CSV, ",")
    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim sheetName As String
        sheetName = Trim$(names(i))
        If Len(sheetName) > 0 Then
            If Not modCommon.SheetExists(sheetName) Then
                Call AppendNewSheet(sheetName)
            End If
        End If
    Next i
End Sub

Private Sub AppendNewSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
End Sub

' ================================================================
' 関数名: SetupAllScreens
' 概要:   14 画面の Setup を順次実行（spec → screen → renderer）
' ================================================================
Private Sub SetupAllScreens()
    Dim screenIds As Variant
    screenIds = modScreenSpecRegistry.GetAllScreenIds()
    Dim i As Long
    For i = LBound(screenIds) To UBound(screenIds)
        Dim screenId As String
        screenId = CStr(screenIds(i))
        Call SetupOneScreen(screenId)
    Next i
End Sub

' ================================================================
' 関数名: SetupOneScreen
' 概要:   1 画面の Setup を実行（エラー耐性 ― 1 画面失敗しても他は続行）
' ================================================================
Private Sub SetupOneScreen(ByVal screenId As String)
    On Error GoTo NextScreen
    Dim screen As Object
    Set screen = modFactory.CreateScreen(screenId, m_renderer)
    If screen Is Nothing Then GoTo NextScreen
    screen.Setup
NextScreen:
    If Err.Number <> 0 Then
        Debug.Print "[clsSetupOrchestrator.SetupOneScreen] WARN " & screenId & ": " & Err.Description
        Err.Clear
    End If
End Sub

' ================================================================
' 関数名: InitializeLogSheet
' 概要:   ログシートのヘッダ行を初期化（既存ロジック踏襲）
' ================================================================
Private Sub InitializeLogSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    If ws Is Nothing Then Exit Sub
    ws.Cells(1, LOG_COL_TIMESTAMP).Value = "日時"
    ws.Cells(1, LOG_COL_MODULE).Value = "モジュール名"
    ws.Cells(1, LOG_COL_FUNCTION).Value = "関数名"
    ws.Cells(1, LOG_COL_LEVEL).Value = "メッセージ種別"
    ws.Cells(1, LOG_COL_MESSAGE).Value = "メッセージ内容"
    Err.Clear
    On Error GoTo 0
End Sub

' ================================================================
' 関数名: InitializeSettingsSheet
' 概要:   設定シートのテストモード初期値を書き込み
' ================================================================
Private Sub InitializeSettingsSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    If ws Is Nothing Then Exit Sub
    ws.Cells(SETTINGS_ROW_TESTMODE, SETTINGS_COL_NAME).Value = "テストモード"
    ws.Cells(SETTINGS_ROW_TESTMODE, SETTINGS_COL_VALUE).Value = TESTMODE_OFF
    Err.Clear
    On Error GoTo 0
End Sub

' ================================================================
' 関数名: SetInitialVisibility
' 概要:   メインのみ可視、その他は xlSheetHidden
' ================================================================
Private Sub SetInitialVisibility()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_MAIN Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetHidden
        End If
    Next ws
End Sub

' ================================================================
' 関数名: DeleteEmptyDefaultSheets
' 概要:   空の Sheet1〜Sheet4 を削除
' ================================================================
Private Sub DeleteEmptyDefaultSheets()
    Dim names() As String
    names = Split(DEFAULT_SHEETS_CSV, ",")
    Dim i As Long
    For i = LBound(names) To UBound(names)
        Call DeleteSheetIfEmpty(Trim$(names(i)))
    Next i
End Sub

Private Sub DeleteSheetIfEmpty(ByVal sheetName As String)
    If Len(sheetName) = 0 Then Exit Sub
    If Not modCommon.SheetExists(sheetName) Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim cellCount As Long
    cellCount = WorksheetFunction.CountA(ws.Cells)
    If cellCount > 0 Then Exit Sub
    If ws.Shapes.Count > 0 Then Exit Sub
    On Error Resume Next
    ws.Visible = xlSheetVisible
    ws.Delete
    Err.Clear
    On Error GoTo 0
End Sub

```
