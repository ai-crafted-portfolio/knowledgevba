---
title: ThisWorkbook.cls
description: ThisWorkbook.cls のソースコード（コピペ用）
---

# ThisWorkbook.cls

**配置先**: `登録修正.xlsm` 用の VBA モジュール  
**種類**: クラス モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\register\`
- ファイル名: `ThisWorkbook.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

---

## ソースコード

```vb
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
' ThisWorkbook 特殊モジュール (登録修正.xlsm 専用)
' 配置:   ThisWorkbook.cls として登録修正.xlsm の VBE に手作業 Import (VBE 仕様で Import 不可、コピペのみ)
' 関連:   clsSetupOrchestrator.bas v2, modConfigLoader.bas, modConfigHolder.bas
' Version: v2.1 (2026-05-16 EOD、Q1-Q57 反映)
' Phase: 7
' ADR:   ADR-0053 §2.1
' v2.1:  登録修正.xlsm 画面構成 = M-05/M-06 (2 画面) + LOG
'        Q44: 起動時 ActiveSheet = ナレッジ登録 (M-05)
'        Q19: format_dir への write/delete は管理.xlsm 限定 (登録修正.xlsm からは read のみ)
'        Q39: config.txt はテキストエディタ直接編集、Workbook_SheetChange ハンドラ不要（M-11 設定画面は生存・debugLevel 編集 GUI、SSOT-Q22）
' ================================================================
Option Explicit

Private Const XLSM_NAME As String = "登録修正"
Private Const STARTUP_SHEET As String = "ナレッジ登録"  ' v2.1 Q44 確定 (M-01 メニュー削除、起動時 = 業務画面直行)

' ================================================================
' Workbook_Open
' 概要:   xlsm 起動時に setup を実行 (登録修正.xlsm 用)
' 手順:   1. modConfigLoader で xlsm 名対応の config.txt を read し modConfigHolder にセット (Q8)
'         2. clsLogger.Init (ログシート + debugLevel ERROR 既定、Q7)
'         3. modKnowledgeFileIO.CleanupOldBackups で 90 日超 backup 自動削除 (Q34)
'         4. clsSetupOrchestrator.RunForRegister  (= Setup_register 経由)
'            - シート確保 (M-05/M-06/LOG)
'            - UI スタンザ適用 (modUILoader.ApplyUiToSheet、Q20)
'            - タブ色 (ADR-0053 §2.1.1)
'            - format_dir 初期化 (既存 .txt が無ければ空フォルダ seed)
'            - Workbook.Protect Structure + シート保護
'            - ActiveSheet = ナレッジ登録 (M-05) (Q44)
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
    oLogger003.LogInfo "ThisWorkbook_touroku", "Workbook_Open", "Workbook_Open 正常完了: " & XLSM_NAME, "", "SAVE-EXIT-OK-II-003"
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
    msg = "登録修正.xlsm 起動エラー: " & Err.Description & vbCrLf & _
          "config.txt の存在 / debugLevel 値 / シート構成 / format_dir 書込権限 を確認してください"
    Debug.Print msg
    ' S5-LOG: BACKTOMAIN-ERR-EE-031 (Workbook_Open failure, 管理画面到達不可 = back-to-main 相当遷移エラー)
    On Error Resume Next
    Dim oLogger031 As clsLogger
    Set oLogger031 = New clsLogger
    oLogger031.Init ThisWorkbook.Worksheets("LOG")
    oLogger031.LogError "ThisWorkbook_touroku", "Workbook_Open", "Workbook_Open 失敗: " & Err.Description, "", "BACKTOMAIN-ERR-EE-031"
    On Error GoTo 0
    ' v2.3 install_admin.bat の headless 実機テストで Setup_admin が
    ' 裏で MsgBox を出してハングする問題への恒久対策。
    ' modCommon.IsHeadless() で COM 自動実行を検出し、その場合は
    ' MsgBox を出さず clsLogger / Debug.Print のみで通知する。
    ' Application.Run "Setup_admin" 経路では Workbook_Open は起動
    ' しないはずだが、何らかの再有効化経路で起動した場合の安全網。
    If Not modCommon.IsHeadless() Then
        MsgBox msg, vbCritical, "Workbook_Open"
    Else
        Debug.Print "[HEADLESS] suppressed MsgBox: " & msg
        modCommon.AppendProgressLog modCommon.ProgressTs() & "ThisWorkbook(管理).Workbook_Open ErrHandler suppressed MsgBox: " & Err.Description
    End If
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
