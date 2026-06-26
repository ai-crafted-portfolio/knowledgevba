---
title: ThisWorkbook.cls
description: ThisWorkbook.cls のソースコード（コピペ用）
---

# ThisWorkbook.cls

**配置先**: `管理.xlsm` 用の VBA モジュール
**種類**: クラスモジュール
**更新日**: 2026-06-11 19:15 JST

---

## 取り込み方法（ThisWorkbook は特別）

**【重要】`ThisWorkbook` は Excel のブック（.xlsm）に最初から存在する document module です。ファイルとして Import できません。**

1. VBE（`Alt + F11`）を開き、プロジェクトツリーの `ThisWorkbook` を **ダブルクリック**してコードペインを開きます。
2. コードペインの**既存コードをすべて削除**します。
3. 下のコード全文をコピーして**貼り付け**ます。

> VB クラスモジュールヘッダ（`VERSION 1.0 CLASS` / `BEGIN` / `END` / `Attribute` 行）は不要です。本体は `Option Explicit` から始まります。
> 文字コード・改行はコードペインに直接貼り付けるため、ファイル保存時の ANSI/CRLF 指定は不要です。

---

## ソースコード

```vb
' ================================================================
' ThisWorkbook 特殊モジュール (管理.xlsm 専用)
' 配置:   ThisWorkbook.cls として管理.xlsm の VBE に手作業 Import (VBE 仕様で Import 不可、コピペのみ)
' 関連:   clsSetupOrchestrator.bas v2, modConfigLoader.bas, modConfigHolder.bas
' Version: v2.1 (2026-05-16 EOD、Q1-Q57 反映)
' Phase: 7
' ADR:   ADR-0053 §2.1
' v2.1:  管理.xlsm 画面構成 = M-02/M-03/M-04/M-10/M-11/M-12/M-14 (7 画面) + LOG
'        Q44: 起動時 ActiveSheet = M-02
'        Q19: format_dir への write/delete は管理.xlsm 限定 (modFormatLoader 内で ThisWorkbook.Name enforce)
'        Q39: config.txt はテキストエディタ直接編集、Workbook_SheetChange ハンドラ不要（M-11 設定画面は生存・debugLevel 編集 GUI、SSOT-Q22）
' ================================================================
Option Explicit

Private Const XLSM_NAME As String = "管理"
Private Const STARTUP_SHEET As String = "フォーマット一覧"  ' v2.1 Q44 確定 (M-01 メニュー削除、起動時 = 業務画面直行)

' ================================================================
' Workbook_Open
' 概要:   xlsm 起動時に setup を実行 (管理.xlsm 用)
' 手順:   1. modConfigLoader で xlsm 名対応の config.txt を read し modConfigHolder にセット (Q8)
'         2. clsLogger.Init (ログシート + debugLevel ERROR 既定、Q7)
'         3. modKnowledgeFileIO.CleanupOldBackups で 90 日超 backup 自動削除 (Q34)
'         4. clsSetupOrchestrator.RunFullSetup("管理")
'            - シート確保 (M-02/M-03/M-04/M-12/M-10/M-11/M-14/LOG)
'            - UI スタンザ適用 (modUILoader.ApplyUiToSheet、Q20)
'            - タブ色 (ADR-0053 §2.1.1)
'            - format_dir 初期化 (既存 .txt が無ければ空フォルダ seed)
'            - Workbook.Protect Structure + シート保護
'            - ActiveSheet = M-02 (Q44)
' ================================================================
Private Sub Workbook_Open()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0001] ThisWorkbook.Workbook_Open ENTER"  ' [ADR-0100]
    ' [ADR-0100][gDebugLevel] init from config with safe fallback
    On Error Resume Next
    modCommon.gDebugLevel = modConfigHolder.GetDebugLevel()
    If Err.Number <> 0 Or modCommon.gDebugLevel < 0 Or modCommon.gDebugLevel > 5 Then
        modCommon.gDebugLevel = DEFAULT_DEBUG_LEVEL_FALLBACK
    End If
    Err.Clear
    On Error GoTo 0
    Debug.Print "[D-INIT] " & XLSM_NAME & " gDebugLevel=" & modCommon.gDebugLevel & " ts=" & Format$(Now, "hh:nn:ss")
    Debug.Print "[WBOPEN-ENTER] " & XLSM_NAME & " ts=" & Format$(Now, "hh:nn:ss")  ' [WBOPEN-DEBUG-PRINT-INJECTED]
    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' v2.1 Q34: 起動時に 90 日超バックアップを自動削除 (管理.xlsm のみ実施)
    On Error Resume Next
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0004] ThisWorkbook.Workbook_Open STEP-1 pre modKnowledgeFileIO.CleanupOldBackups"  ' [ADR-0100]
    Call modKnowledgeFileIO.CleanupOldBackups
    On Error GoTo ErrHandler

    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0005] ThisWorkbook.Workbook_Open STEP-2 pre orch.RunFullSetup"  ' [ADR-0100]
    Call orch.RunFullSetup(XLSM_NAME)
    ' 2026-06-07: Seed C4 with next FmtId after RunFullSetup so first-open shows FMT-NNN.
    On Error Resume Next
    modEntryFormat.SeedM03FormatIdIfEmpty
    On Error GoTo ErrHandler
    ' 2026-06-06: schedule a delayed re-render so ApplyFreeze runs AFTER
    ' the window becomes fully visible. Skip during install (COM Visible=False)
    ' because OnTime would fire after Excel quits, causing residual error dialog.
    On Error Resume Next
    If Application.UserControl Then
        Application.OnTime Now + TimeValue("0:00:01"), "modRefresh.Btn_RefreshAllSheets"
        ' [BUGFIX 2026-06-06] After re-render finishes, unlock B5/H7/H8/H9
        ' on M-11 so user can toggle CheckBoxes and change debugLevel
        ' dropdown without hitting the sheet-protection dialog.
        Application.OnTime Now + TimeValue("0:00:02"), "modEntrySettings.EnsureSettingsCellsEditable"
    End If
    Err.Clear
    On Error GoTo ErrHandler

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
    Debug.Print "[WBOPEN-EXIT] " & XLSM_NAME & " ts=" & Format$(Now, "hh:nn:ss")  ' [WBOPEN-DEBUG-PRINT-INJECTED]
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0002] ThisWorkbook.Workbook_Open EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0003] ThisWorkbook.Workbook_Open EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[WBOPEN-ERR] " & XLSM_NAME & " err=" & Err.Number & " " & Err.Description  ' [WBOPEN-DEBUG-PRINT-INJECTED]
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
    ' v2.3 install_admin.bat の headless 実機テストで Setup_admin が
    ' 裏で MsgBox を出してハングする問題への恒久対策。
    ' modCommon.IsHeadless() で COM 自動実行を検出し、その場合は
    ' MsgBox を出さず clsLogger / Debug.Print のみで通知する。
    ' Application.Run "Setup_admin" 経路では Workbook_Open は起動
    ' しないはずだが、何らかの再有効化経路で起動した場合の安全網。
    If Not modCommon.IsHeadless() Then
        MsgBox msg, vbCritical, ChrW(&H8D77) & ChrW(&H52D5) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
    Else
        Debug.Print "[HEADLESS] suppressed MsgBox: " & msg
        modCommon.AppendProgressLog modCommon.ProgressTs() & "ThisWorkbook(管理).Workbook_Open ErrHandler suppressed MsgBox: " & Err.Description
    End If
End Sub

' ================================================================
' Workbook_BeforeClose
' ================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0006] ThisWorkbook.Workbook_BeforeClose ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0007] ThisWorkbook.Workbook_BeforeClose EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' Workbook_SheetBeforeDoubleClick (B19 restore, 2026-06-11)
' Role:   M-02 / M-10 row-select UX. Double-click on column A toggles
'         the selection mark.
'         - M-02 (format list)  rows 17..60: check ChrW(2713) <-> empty
'         - M-10 (storage)      rows 11..15: single-select radio
'           ChrW(25CF)=selected / ChrW(25CB)=not selected
' Why:    The 2026-06-08 build had this handler (see
'         evidence_summary_2026-06-08.md sec 2.1) but it was never
'         backported to canonical, so every reinstall (ThisWorkbook
'         body injection) silently dropped it = regression B19.
' Note:   Unprotect -> write -> Protect(UserInterfaceOnly) pattern is
'         required because UserInterfaceOnly does NOT survive a
'         save/reopen cycle (same root cause as B14).
' ================================================================
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrHandler
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0008] ThisWorkbook.Workbook_SheetBeforeDoubleClick ENTER"  ' [ADR-0100]
    ' Sheet names via ChrW (CP932-safe): M-02 / M-10
    Dim nmM02 As String
    nmM02 = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H4E00) & ChrW(&H89A7)
    Dim nmM10 As String
    nmM10 = ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & ChrW(&H8A2D) & ChrW(&H5B9A)
    If Sh.Name = nmM02 Then
        Call ToggleM02RowCheck(Sh, Target, Cancel)
    ElseIf Sh.Name = nmM10 Then
        Call ToggleM10Radio(Sh, Target, Cancel)
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0009] ThisWorkbook.Workbook_SheetBeforeDoubleClick EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] SheetBeforeDoubleClick toggle: " & Err.Number & " " & Err.Description
End Sub

' ================================================================
' ToggleM02RowCheck (B19)
' Role: M-02 col A rows 17..60 - toggle check mark on double-click.
' ================================================================
Private Sub ToggleM02RowCheck(ByVal Sh As Object, ByVal Target As Range, ByRef outCancel As Boolean)
    Const COL_CHECK As Long = 1
    Const FIRST_ROW As Long = 17   ' = modEntryFormat M02_DATA_FIRST_ROW
    Const LAST_ROW As Long = 60    ' = modEntryFormat M02_DATA_LAST_ROW
    If Target.Column <> COL_CHECK Then Exit Sub
    If Target.Row < FIRST_ROW Or Target.Row > LAST_ROW Then Exit Sub
    outCancel = True
    On Error Resume Next
    Sh.Unprotect
    If CStr(Target.Value) = ChrW(&H2713) Then
        Target.Value = ""
    Else
        Target.Value = ChrW(&H2713)
    End If
    Sh.Protect Password:="", UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

' ================================================================
' ToggleM10Radio (B19)
' Role: M-10 col A rows 11..15 - single-select radio on double-click.
'       Selected row -> deselect; other row -> clear all + select it.
'       2026-06-11: rows 11..15 (5th row = config_dir was added after
'       the original 06-08 handler which covered 11..14 only).
' ================================================================
Private Sub ToggleM10Radio(ByVal Sh As Object, ByVal Target As Range, ByRef outCancel As Boolean)
    Const COL_CHECK As Long = 1
    Const FIRST_ROW As Long = 11
    Const LAST_ROW As Long = 15
    If Target.Column <> COL_CHECK Then Exit Sub
    If Target.Row < FIRST_ROW Or Target.Row > LAST_ROW Then Exit Sub
    outCancel = True
    On Error Resume Next
    Sh.Unprotect
    Dim ri As Long
    If CStr(Target.Value) = ChrW(&H25CF) Then
        Target.Value = ChrW(&H25CB)
    Else
        For ri = FIRST_ROW To LAST_ROW
            Sh.Cells(ri, COL_CHECK).Value = ChrW(&H25CB)
        Next ri
        Target.Value = ChrW(&H25CF)
    End If
    Sh.Protect Password:="", UserInterfaceOnly:=True
    On Error GoTo 0
End Sub
```
