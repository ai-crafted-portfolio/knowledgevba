---
title: clsSetupOrchestrator.cls
---

# clsSetupOrchestrator.cls

| 項目 | 内容 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | ブック起動時のセットアップ一括処理（設定読込→ログ初期化→シート構築→保護→起動シート表示） |
| 行数 | 254 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsSetupOrchestrator.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "clsSetupOrchestrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' クラス: clsSetupOrchestrator（v2.1、Phase 5-7、3 xlsm 共 setup）
' 概要:   xlsm 起動時の setup orchestration（config 読込 → ログ初期化 → シート構築 → 保護 → ActiveSheet）
' Version: v2.1（2026-05-16 EOD、Q1-Q57 回答反映）
' 関連:   ADR-0053 §2.1 / §2.5 / §2.8 #4, Q44 / Q39-Q44 / Q34
' SSOT:   本クラスは本番 runtime setup の SSOT（ADR-0053 §2 / SSOT-Q17）。
'         3 xlsm のシート構成定数・起動シート定数は ADR-0053 §2.1 表が文書 SSOT。
' v2.1 主要更新:
'   - 画面構成（ADR-0053 §2.1.1 サブグループ順）: 登録修正=M-05/M-06、検索=M-07/M-08/M-09、管理=M-02/M-03/M-04/M-12/M-13/M-10/M-11/M-14（LOG 共通）
'   - 起動時 ActiveSheet（ADR-0053 §2.1 / §9 M-2）: 登録修正=M-05、検索=M-08、管理=M-02
'   - Q39: 管理.xlsm の SheetChange ハンドラ不要（M-11 は旧 ID 体系の『M-11 削除』記述を訂正、新 ID 体系では M-11=設定で現役。architecture §10-3 L37 注記と整合）
'   - Q34: 管理.xlsm 起動時に CleanupOldBackups を呼出
'   - 工数記述廃止（Q30）
' ================================================================
Option Explicit

Private m_logger As clsLogger

' xlsm 別 sheet 構成（ADR-0053 §2.1 / modSetup.bas SHEETS_* と同値、R6-01）
' SSOT: ADR-0053 §2.1 表（§2.1.1 サブグループ順）。登録修正=M-05/M-06、検索=M-07/M-08/M-09、管理=M-02/M-03/M-04/M-12/M-13/M-10/M-11/M-14（LOG は全 xlsm 共通）
Private Const SHEETS_TOUROKU As String = "M-05|M-06|LOG"
Private Const SHEETS_KENSAKU As String = "M-07|M-08|M-09|LOG"
Private Const SHEETS_KANRI As String = "M-02|M-03|M-04|M-12|M-13|M-10|M-11|M-14|LOG"

' タブ色（ADR-0053 §2.1.1）
Private Const TAB_COLOR_TOUROKU As Long = 16711892  ' #FFB6C1 pink (BGR: B6C1FF -> FFC1B6 reversed)
Private Const TAB_COLOR_KENSAKU As Long = 15128749  ' #ADD8E6 lightblue
Private Const TAB_COLOR_KANRI As Long = 9498256     ' #90EE90 lightgreen

' 起動時 ActiveSheet（ADR-0053 §2.1 表 / §9 M-2、R6-01）
' SSOT: 登録修正=M-05、検索=M-08、管理=M-02
Private Const STARTUP_SHEET_TOUROKU As String = "M-05"
Private Const STARTUP_SHEET_KENSAKU As String = "M-08"
Private Const STARTUP_SHEET_KANRI As String = "M-02"

' ----------------------------------------------------------------
' Public API: spec §6.1 convenience wrappers
' ----------------------------------------------------------------

Public Sub RunForRegister()
    RunFullSetup "登録修正"
End Sub

Public Sub RunForSearch()
    RunFullSetup "検索"
End Sub

Public Sub RunForAdmin()
    RunFullSetup "管理"
End Sub

' ----------------------------------------------------------------
' Public API
' ----------------------------------------------------------------

' 概要: xlsm 起動時の全 setup を実行
' 引数: xlsmName = "登録修正" / "検索" / "管理"
Public Sub RunFullSetup(ByVal xlsmName As String)
    On Error GoTo ErrHandler

    ' (1) modConfigLoader で xlsm 別 config.txt を read し modConfigHolder にセット（Q8）
    Call modConfigLoader.LoadConfig(xlsmName)

    ' (2) clsLogger 初期化（Q7 既定 ERROR）
    EnsureLogSheet xlsmName
    Set m_logger = New clsLogger
    Call m_logger.Init(ThisWorkbook.Worksheets("LOG"))
    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "起動: " & xlsmName, "", "LOG-SETUP-START"

    ' (3) 管理.xlsm のみ: 起動時に 90 日超バックアップ自動削除（Q34）
    If xlsmName = "管理" Then
        Dim deleted As Long
        deleted = modKnowledgeFileIO.CleanupOldBackups()
        If deleted > 0 Then
            m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", _
                "90 日超バックアップ削除: " & deleted & " 件", "", "LOG-BCK-CLEAN"
        End If
    End If

    ' (4) シート確保（v2.1 構成、画面廃止は M-01 のみ）
    EnsureSheets xlsmName

    ' (5) タブ色設定（ADR-0053 §2.1.1）
    ApplyTabColors xlsmName

    ' (6) UI スタンザ適用（Q20 modUILoader.ApplyUiToSheet）
    ApplyUiStanzas xlsmName

    ' (7) Workbook.Protect Structure + シート保護
    ApplyProtection

    ' (8) ActiveSheet = 起動時画面（Q44）
    ActivateStartupSheet xlsmName

    m_logger.LogInfo "clsSetupOrchestrator", "RunFullSetup", "完了: " & xlsmName, "", "LOG-SETUP-OK"
    Exit Sub

ErrHandler:
    ' Q7 規約 X: error handler 内で必ず LogError
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsSetupOrchestrator", "RunFullSetup", Err.Description, "", "LOG-SETUP-ERR"
    Else
        Debug.Print "[clsSetupOrchestrator.RunFullSetup ERROR] " & Err.Description
    End If
End Sub

' ----------------------------------------------------------------
' Private: setup ステップ
' ----------------------------------------------------------------

Private Sub EnsureLogSheet(ByVal xlsmName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("LOG")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "LOG"
    End If
    ' Q7 規約 Y
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub EnsureSheets(ByVal xlsmName As String)
    Dim sheetList As String
    sheetList = GetSheetsCsv(xlsmName)
    If Len(sheetList) = 0 Then Exit Sub

    Dim names() As String
    names = Split(sheetList, "|")

    Dim i As Long
    For i = LBound(names) To UBound(names)
        EnsureSheetExists names(i)
    Next i
End Sub

Private Sub EnsureSheetExists(ByVal sheetName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub ApplyTabColors(ByVal xlsmName As String)
    Dim sheetList As String
    sheetList = GetSheetsCsv(xlsmName)
    If Len(sheetList) = 0 Then Exit Sub

    Dim tabColor As Long
    tabColor = GetTabColor(xlsmName)

    Dim names() As String
    names = Split(sheetList, "|")

    Dim i As Long
    On Error Resume Next
    For i = LBound(names) To UBound(names)
        If names(i) <> "LOG" Then
            ThisWorkbook.Worksheets(names(i)).Tab.Color = tabColor
     