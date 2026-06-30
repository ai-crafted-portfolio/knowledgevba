---
title: modRefresh.bas
description: modRefresh.bas のソースコード（コピペ用）
---

# modRefresh.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-30 15:25 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modRefresh.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modRefresh"
' ============================================================
' modRefresh (Phase R-3-psi-Refresh, 2026-05-29)
'   ui_seed/<role>/M-NN.txt を edit したあと、再 install せずに sheet を
'   再描画するための entry point。button OnClick / Alt+F8 macro から呼ぶ。
'   実体は clsSetupOrchestrator.Reapply* (RunFullSetup step6 ApplyUiStanzas
'   と同一 pipeline: BindSheet -> ClearScreen -> ApplyFromStanza -> ProtectSheet)。
'   本モジュールは ScreenUpdating/EnableEvents 抑止 + ActiveSheet 復元の
'   薄いラッパのみ。ASCII-only (CP932/UTF-8 round-trip 安全, ADR-0006)。
' ============================================================
Option Explicit

' 全 sheet (LOG 以外) を ui_seed から再描画。button / Alt+F8 用。
Public Sub Btn_RefreshAllSheets()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1318] modRefresh.Btn_RefreshAllSheets ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_RefreshAllSheets"
    Dim XLSM As String
    XLSM = Replace$(ThisWorkbook.Name, ".xlsm", "")
    modBtnGuard.LogEnter BTN, XLSM
    modCommon.AppendProgressLog "[BTN-REFRESH-ENTER] " & Now
    ' [Fix-6 / ADR-0133-followup 2026-06-19] Refresh is a pure seed re-render
    ' (RunFullSetup loads + auto-heals config itself). It MUST NOT require
    ' "config" as a prerequisite: a migrated workbook with an empty settings
    ' cell used to abort here, which also blocked the Workbook_Open auto-
    ' re-render that recreates orphaned button shapes. Empty requirement = OK.
    If Not modBtnGuard.CheckPrereq(BTN, "", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1319] modRefresh.Btn_RefreshAllSheets EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    ' [Change-3] storage-access fallback: keep the previous screen and
    ' skip this re-render when a configured storage folder is unreachable.
    Dim badDir As String
    badDir = modConfigHolder.EnsureStorageAccessible(XLSM)
    If Len(badDir) > 0 Then
        modConfigHolder.NotifyStorageInaccessible badDir
        modBtnGuard.LogExit BTN, XLSM, False
        Exit Sub
    End If
    ' [P5d fix 2026-06-04] ScreenUpdating=False 削除 (shape 描画 silent fail 回避)
    ' [Fix-6 followup 2026-06-19] Snapshot the 9 settings cells before the
    ' re-render so ApplyUiStanzas' Cells.Clear cannot lose user-entered
    ' values; written back verbatim after the pipeline (RestoreSettings).
    Dim settingsBackup As Object
    Set settingsBackup = modConfigLoader.SnapshotSettings()
    Dim prevEv As Boolean
    Dim activeWs As Object
    On Error Resume Next
    Set activeWs = ThisWorkbook.ActiveSheet
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo 0

    modCommon.AppendProgressLog "[BTN-REFRESH-PASS-CHECKPREREQ]"
    ' [P5e fix 2026-06-04] install と同じ RunFullSetup を invoke (ReapplyAllSheets による画面破壊回避)
    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    Dim baseName As String
    baseName = Replace(ThisWorkbook.Name, ".xlsm", "")
    modCommon.AppendProgressLog "[BTN-REFRESH-DISPATCH] baseName=" & baseName
    Select Case baseName
        Case ChrW(&H7BA1) & ChrW(&H7406): orch.RunForAdmin
        Case ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63): orch.RunForRegister
        Case ChrW(&H691C) & ChrW(&H7D22): orch.RunForSearch
        Case Else: orch.RunForAdmin
    End Select

    On Error Resume Next
    If Not activeWs Is Nothing Then activeWs.Activate
    ' [B35 2026-06-12] after a full re-render the sheet can stay in a non-
    ' interactive paint state: cells look frozen and dropdown arrows do not
    ' appear until the user switches sheets. Force a screen + selection
    ' refresh so the active sheet is immediately usable.
    Application.ScreenUpdating = True
    Application.Interactive = True
    If Not activeWs Is Nothing Then
        activeWs.Range(activeWs.Cells(1, 1), activeWs.Cells(1, 1)).Select
    End If
    Application.EnableEvents = prevEv
    Application.StatusBar = RefreshedMsg()
    On Error GoTo 0
    ' 2026-06-07: Re-seed M-03 C4 + reload M-10 storage paths after refresh
    If XLSM = ChrW(&H7BA1) & ChrW(&H7406) Then
        On Error Resume Next
        Application.Run ThisWorkbook.Name & "!modEntryFormat.SeedM03FormatIdIfEmpty"
        Dim wsStor As Worksheet
        Set wsStor = ThisWorkbook.Worksheets(ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & ChrW(&H8A2D) & ChrW(&H5B9A))
        Application.Run ThisWorkbook.Name & "!modEntrySettings.LoadStorageToSheet", wsStor
        On Error GoTo 0
    End If
    ' [Fix-6 followup 2026-06-19] Write the snapshot back so the 9 M-10
    ' cells survive the re-render regardless of holder/config state.
    On Error Resume Next
    modConfigLoader.RestoreSettings settingsBackup
    On Error GoTo 0
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1320] modRefresh.Btn_RefreshAllSheets EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1321] modRefresh.Btn_RefreshAllSheets EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    ' [P2 fix] ScreenUpdating/EnableEvents restore before HandleButtonError
    On Error Resume Next
    If Not activeWs Is Nothing Then activeWs.Activate
    Application.EnableEvents = prevEv
    On Error GoTo 0
    modBtnGuard.HandleButtonError BTN, Err, XLSM
    modBtnGuard.LogExit BTN, XLSM, False
End Sub

' ActiveSheet 1 枚だけ再描画。Alt+F8 / button 用。
Public Sub Btn_RefreshSheet_Active()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1322] modRefresh.Btn_RefreshSheet_Active ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_RefreshSheet_Active"
    Dim XLSM As String
    XLSM = Replace$(ThisWorkbook.Name, ".xlsm", "")
    modBtnGuard.LogEnter BTN, XLSM
    ' [Fix-6 / ADR-0133-followup 2026-06-19] Refresh is a pure seed re-render
    ' (RunFullSetup loads + auto-heals config itself). It MUST NOT require
    ' "config" as a prerequisite: a migrated workbook with an empty settings
    ' cell used to abort here, which also blocked the Workbook_Open auto-
    ' re-render that recreates orphaned button shapes. Empty requirement = OK.
    If Not modBtnGuard.CheckPrereq(BTN, "", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1323] modRefresh.Btn_RefreshSheet_Active EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    ' [Change-3] storage-access fallback: keep the previous screen and
    ' skip this re-render when a configured storage folder is unreachable.
    Dim badDir As String
    badDir = modConfigHolder.EnsureStorageAccessible(XLSM)
    If Len(badDir) > 0 Then
        modConfigHolder.NotifyStorageInaccessible badDir
        modBtnGuard.LogExit BTN, XLSM, False
        Exit Sub
    End If
    Dim prevSU As Boolean, prevEv As Boolean
    Dim activeWs As Object
    On Error Resume Next
    Set activeWs = ThisWorkbook.ActiveSheet
    prevSU = Application.ScreenUpdating
    prevEv = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo 0

    ' [P5e horizontal expand 2026-06-04] install と同じ RunFullSetup 経路 (画面破壊回避)
    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    Dim baseName As String
    baseName = Replace(ThisWorkbook.Name, ".xlsm", "")
    modCommon.AppendProgressLog "[BTN-REFRESH-ACTIVE-DISPATCH] baseName=" & baseName
    Select Case baseName
        Case ChrW(&H7BA1) & ChrW(&H7406): orch.RunForAdmin
        Case ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63): orch.RunForRegister
        Case ChrW(&H691C) & ChrW(&H7D22): orch.RunForSearch
        Case Else: orch.RunForAdmin
    End Select

    On Error Resume Next
    If Not activeWs Is Nothing Then activeWs.Activate
    Application.EnableEvents = prevEv
    Application.ScreenUpdating = prevSU
    Application.StatusBar = RefreshedMsg()
    On Error GoTo 0
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1324] modRefresh.Btn_RefreshSheet_Active EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1325] modRefresh.Btn_RefreshSheet_Active EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.HandleButtonError BTN, Err, XLSM
    modBtnGuard.LogExit BTN, XLSM, False
End Sub

' 指定 screenId (M-NN) 1 枚を再描画。programmatic 用 (button OnAction 非対象)。
Public Sub Btn_RefreshSheet(ByVal screenId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1326] modRefresh.Btn_RefreshSheet ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_RefreshSheet"
    Dim XLSM As String
    XLSM = Replace$(ThisWorkbook.Name, ".xlsm", "")
    modBtnGuard.LogEnter BTN, XLSM
    ' [Fix-6 / ADR-0133-followup 2026-06-19] Refresh is a pure seed re-render
    ' (RunFullSetup loads + auto-heals config itself). It MUST NOT require
    ' "config" as a prerequisite: a migrated workbook with an empty settings
    ' cell used to abort here, which also blocked the Workbook_Open auto-
    ' re-render that recreates orphaned button shapes. Empty requirement = OK.
    If Not modBtnGuard.CheckPrereq(BTN, "", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1327] modRefresh.Btn_RefreshSheet EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    ' [Change-3] storage-access fallback: keep the previous screen and
    ' skip this re-render when a configured storage folder is unreachable.
    Dim badDir As String
    badDir = modConfigHolder.EnsureStorageAccessible(XLSM)
    If Len(badDir) > 0 Then
        modConfigHolder.NotifyStorageInaccessible badDir
        modBtnGuard.LogExit BTN, XLSM, False
        Exit Sub
    End If
    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    orch.ReapplySheet screenId
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1328] modRefresh.Btn_RefreshSheet EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1329] modRefresh.Btn_RefreshSheet EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.HandleButtonError BTN, Err, XLSM
    modBtnGuard.LogExit BTN, XLSM, False
End Sub

' "更新完了" を ChrW で構築 (本 .bas を ASCII-only に保つ)。
Private Function RefreshedMsg() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1330] modRefresh.RefreshedMsg ENTER"  ' [ADR-0100]
    RefreshedMsg = ChrW(&H66F4) & ChrW(&H65B0) & ChrW(&H5B8C) & ChrW(&H4E86)
End Function
```
