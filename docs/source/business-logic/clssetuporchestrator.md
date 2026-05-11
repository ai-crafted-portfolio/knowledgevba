---
title: clsSetupOrchestrator.cls
---

# clsSetupOrchestrator.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | セットアップ全工程 (シート生成 / spec 読込 / 描画依頼) を統括する制御クラス |
| 行数 | 187 行 |

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
' 依存先: IScreenRenderer, modFactory, modScreenSpecRegistry, modCommon, clsLogger
' 備考:   v21 (E2E rerun) で全 Sub に ENTER/EXIT/step ログを注入。
'         SetupOneScreen の silent NextScreen → LogError で必ず痕跡を残す。
' ================================================================

Private Const MOD_NAME As String = "clsSetupOrchestrator"

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
    Dim stepName As String : stepName = "begin"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    stepName = "CreateRequiredSheets"
    Call LogEnter("RunFullSetup", stepName)
    Call CreateRequiredSheets

    ' この時点で SHEET_LOG が存在するため、これ以降は logger を使える
    Call LogTraceSafe("RunFullSetup", "after CreateRequiredSheets - sheets created")

    stepName = "SetupAllScreens"
    Call LogTraceSafe("RunFullSetup", "STEP " & stepName)
    Call SetupAllScreens

    stepName = "InitializeSettingsSheet"
    Call LogTraceSafe("RunFullSetup", "STEP " & stepName)
    Call InitializeSettingsSheet

    stepName = "SetInitialVisibility"
    Call LogTraceSafe("RunFullSetup", "STEP " & stepName)
    Call SetInitialVisibility

    stepName = "DeleteEmptyDefaultSheets"
    Call LogTraceSafe("RunFullSetup", "STEP " & stepName)
    Call DeleteEmptyDefaultSheets

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Call LogTraceSafe("RunFullSetup", "EXIT ok")
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Call LogErrorSafe("RunFullSetup", stepName, Err.Number, Err.Description)
    Err.Raise Err.Number, "clsSetupOrchestrator.RunFullSetup", Err.Description
End Sub

' ================================================================
' 関数名: CreateRequiredSheets
' 概要:   14 シートを順次作成（既存はスキップ）
' ================================================================
Private Sub CreateRequiredSheets()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "split CSV"

    Dim names() As String
    names = Split(REQUIRED_SHEETS_CSV, ",")

    Dim createdCount As Long : createdCount = 0
    Dim skippedCount As Long : skippedCount = 0

    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim sheetName As String
        sheetName = Trim$(names(i))
        stepName = "check " & sheetName
        If Len(sheetName) > 0 Then
            If Not modCommon.SheetExists(sheetName) Then
                stepName = "append " & sheetName
                Call AppendNewSheet(sheetName)
                createdCount = createdCount + 1
            Else
                skippedCount = skippedCount + 1
            End If
        End If
    Next i

    Call LogTraceSafe("CreateRequiredSheets", _
                       "created=" & createdCount & " skipped=" & skippedCount)
    Exit Sub

ErrHandler:
    Call LogErrorSafe("CreateRequiredSheets", stepName, Err.Number, Err.Description)
    Err.Raise Err.Number, MOD_NAME & ".CreateRequiredSheets", Err.Description
End Sub

Private Sub AppendNewSheet(ByVal sheetName As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    Exit Sub
ErrHandler:
    Call LogErrorSafe("AppendNewSheet", "Add+Name " & sheetName, Err.Number, Err.Description)
    Err.Raise Err.Number, MOD_NAME & ".AppendNewSheet", Err.Description
End Sub

' ================================================================
' 関数名: SetupAllScreens
' 概要:   14 画面の Setup を順次実行（spec → screen → renderer）
' ================================================================
Private Sub SetupAllScreens()
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "GetAllScreenIds"

    Dim screenIds As Variant
    screenIds = modScreenSpecRegistry.GetAllScreenIds()

    Dim okCount As Long : okCount = 0
    Dim ngCount As Long : ngCount = 0

    Dim i As Long
    For i = LBound(screenIds) To UBound(screenIds)
        Dim screenId As String
        screenId = CStr(screenIds(i))
        stepName = "Setup " & screenId
        Call LogTraceSafe("SetupAllScreens", "BEGIN " & screenId)
        If SetupOneScreenLogged(screenId) Then
            okCount = okCount + 1
        Else
            ngCount = ngCount + 1
        End If
    Next i

    Call LogTraceSafe("SetupAllScreens", _
                       "EXIT ok=" & okCount & " ng=" & ngCount)
    Exit Sub

ErrHandler:
    Call LogErrorSafe("SetupAllScreens", stepName, Err.Number, Err.Description)
    Err.Raise Err.Number, MOD_NAME & ".SetupAllScreens", Err.Description
End Sub

' ================================================================
' 関数名: SetupOneScreenLogged
' 概要:   1 画面の Setup を実行。失敗しても他画面の続行のため True/False で返す。
'         失敗時は LogError で痕跡を残す（旧 SetupOneScreen の silent NextScreen
'         を廃止し、必ず Err.Number/Description を記録）。
' 戻り値: Boolean - Setup 成功なら True
' ================================================================
Private Function SetupOneScreenLogged(ByVal screenId As String) As Boolean
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "CreateScreen " & screenId

    Dim screen As Object
    Set screen = modFactory.CreateScreen(screenId, m_renderer)
    If screen Is Nothing Then
        Call LogWarnSafe("SetupOneScreenLogged", screenId & " screen=Nothing (spec missing or unknown id)"
```
