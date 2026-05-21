---
title: ThisWorkbook（管理.xlsm）
---

# ThisWorkbook（管理.xlsm）

| 項目 | 内容 |
|---|---|
| 層 | 特殊モジュール |
| 種別 | ドキュメントモジュール (ThisWorkbook) |
| 配置ブック | 管理.xlsm |
| 役割 | 管理.xlsm 用の ThisWorkbook。起動時の設定読込・セットアップ・終了ログ |
| 行数 | 107 行 |

## 取り込み先

ブックに最初から存在する `ThisWorkbook` モジュール（VBE のプロジェクトツリーにある `ThisWorkbook`）に貼り付けます。新規モジュールとしては取り込めません。先頭の `VERSION 1.0 CLASS` から `Attribute` 行までのファイル先頭ブロックは貼り付けず、その下の本体だけをコードペインに貼り付けて既存の内容を置き換えてください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ================================================================
' ThisWorkbook 特殊モジュール (管理.xlsm 専用)
' 配置:   ThisWorkbook.cls として管理.xlsm の VBE に手作業 Import (VBE 仕様で Import 不可、コピペのみ)
' 関連:   clsSetupOrchestrator.bas v2, modConfigLoader.bas, modConfigHolder.bas
' Version: v2.1 (2026-05-16 EOD、Q1-Q57 反映)
' Phase: 7
' ADR:   ADR-0053 §2.1
' v2.1:  管理.xlsm 画面構成 = M-02/M-03/M-04/M-10/M-11/M-12/M-13/M-14 (8 画面) + LOG
'        Q44: 起動時 ActiveSheet = M-02
'        Q19: format_dir への write/delete は管理.xlsm 限定 (modFormatLoader 内で ThisWorkbook.Name enforce)
'        Q39: config.txt はテキストエディタ直接編集、Workbook_SheetChange ハンドラ不要（M-11 設定画面は生存・debugLevel 編集 GUI、SSOT-Q22）
' ================================================================
Option Explicit

Private Const XLSM_NAME As String = "管理"
Private Const STARTUP_SHEET As String = "M-02"  ' v2.1 Q44 確定 (M-01 メニュー削除、起動時 = 業務画面直行)

' ================================================================
' Workbook_Open
' 概要:   xlsm 起動時に setup を実行 (管理.xlsm 用)
' 手順:   1. modConfigLoader で xlsm 名対応の config.txt を read し modConfigHolder にセット (Q8)
'         2. clsLogger.Init (ログシート + debugLevel ERROR 既定、Q7)
'         3. modKnowledgeFileIO.CleanupOldBackups で 90 日超 backup 自動削除 (Q34)
'         4. clsSetupOrchestrator.RunFullSetup("管理")
'            - シート確保 (M-02/M-03/M-04/M-12/M-13/M-10/M-11/M-14/LOG)
'            - UI スタンザ適用 (modUILoader.ApplyUiToSheet、Q20)
'            - タブ色 (ADR-0053 §2.1.1)
'            - format_dir 初期化 (既存 .txt が無ければ空フォルダ seed)
'            - Workbook.Protect Structure + シート保護
'            - ActiveSheet = M-02 (Q44)
' ================================================================
Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' v2.1 Q34: 起動時に 90 日超バックアップを自動削除 (管理.xlsm のみ実施)
    On Error Resume Next
    Call modKnowledgeFileIO.CleanupOldBackups
    On Error GoTo ErrHandler

    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    Call orch.RunFullSetup(XLSM_NAME)

    ' S5-LOG-02: SAVE-EXIT-OK-II-003 (Workbook_Open success exit, screen 管理)
    On Error Resume Next
    Dim oLogger003 As clsLogger
    Set oLogger003 = New clsLogger
    oLogger003.Init ThisWorkbook.Worksheets("LOG")
    oLogger003.LogInfo "ThisWorkbook_管理", "Workbook_Open", "Workbook_Open 正常完了: " & XLSM_NAME, "", "SAVE-EXIT-OK-II-003"
    On Error GoTo 0

    ' Q44: startup ActiveSheet = STARTUP_SHEET (production spec mirror)
    On Error Resume Next
    ThisWorkbook.Worksheets(STARTUP_SHEET).Activate
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Dim msg As String
    msg = "管理.xlsm 起動エラー: " & Err.Description & vbCrLf & _
          "config.txt の存在 / debugLevel 値 / シート構成 / format_dir 書込権限 を確認してください"
    Debug.Print msg
    ' S5-LOG: BACKTOMAIN-ERR-EE-031 (Workbook_Open failure, 管理画面到達不可 = back-to-main 相当遷移エラー)
    On Error Resume Next
    Dim oLogger031 As clsLogger
    Set oLogger031 = New clsLogger
    oLogger031.Init ThisWorkbook.Worksheets("LOG")
    oLogger031.LogError "ThisWorkbook_管理", "Workbook_Open", "Workbook_Open 失敗: " & Err.Description, "", "BACKTOMAIN-ERR-EE-031"
    On Error GoTo 0
    MsgBox msg, vbCritical, "Workbook_Open"
End Sub

' ================================================================
' Workbook_BeforeClose
' ================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("LOG")
    If Not ws Is Nothing Then
        Dim r As Long
        r = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1
        If r < 9 Then r = 9
        ws.Cells(r, 1).Value = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
        ws.Cells(r, 2).Value = "ThisWorkbook"
        ws.Cells(r, 3).Value = "BeforeClose"
        ws.Cells(r, 4).Value = "INFO"
        ws.Cells(r, 5).Value = "xlsm 終了: " & XLSM_NAME
    End If
    On Error GoTo 0
End Sub
```
