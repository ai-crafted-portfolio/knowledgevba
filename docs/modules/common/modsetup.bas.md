---
title: modSetup.bas
description: modSetup.bas のソースコード（コピペ用）
---

# modSetup.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-30 14:44 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modSetup.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modSetup"
' ============================================================
' modSetup (Installer layer, v2.1 / 3-xlsm + v2.3 Phase A wrappers)
' Role:
'   Common installer entry for the 3 xlsm books (??? / ???? / ?o?^?C??).
'   Each xlsm calls its dedicated Public Sub Setup_xxx which:
'     - ensures required sheets (LOG / M-*) exist (creates if missing)
'     - seeds SystemSettings sheets (???.xlsm only)
'     - emits success/failure logId
'       (SAVE-EXIT-OK-II-009 / BACKTOMAIN-ERR-EE-033)
' Design SSOT:
'   - vba_module_placement_v2.md  1.1 (modSetup = all xlsm, 3 xlsm setup)
'   - installer_3xlsm_v2_main.md  1.4 / 1.5 (Step 5-7 per-xlsm placement)
'   - architecture_v2_3xlsm_stanza.md  7
' Logging:
'   - LogID convention = logging_naming_cheatsheet_v1.md  3 (II / EE suffix)
'   - max SEQ + 1 (audit script verified)
' ASCII-only inside VBA string literals (CP932 mojibake avoidance).
' v2.3 (Phase A): ASCII Setup_admin / Setup_register / Setup_search added
'                 at the bottom of this module; they delegate to
'                 clsSetupOrchestrator (re-use, no logic duplication).
' ============================================================
Option Explicit

Private Const MOD_NAME As String = "modSetup"

' v2.3 Phase A: SETUP_ROLE_* constants moved here (declarations section)
' to fix VBA "variable not defined" compile error caused by them being
' placed after Public Sub bodies (lines moved from below original position).
Private Const SETUP_ROLE_ADMIN    As String = "admin"
Private Const SETUP_ROLE_REGISTER As String = "register"
Private Const SETUP_ROLE_SEARCH   As String = "search"

' v2.3 Phase J (DRIFT-iter2 / ADR-0068, 2026-05-31):
'   Setup_kanri_legacy / Setup_kensaku_legacy / Setup_touroku_legacy
'   removed (legacy v2.1 install entries). The corresponding SHEETS_*
'   constants referenced retired screens (M-13 / M-07 / M-09), so
'   they DRIFTed from clsSetupOrchestrator (the v2.3 SSOT). Deleted
'   along with the helpers (SetupCore / EnsureSheets / SheetExists /
'   SeedSystemSettings / SeedStorage / SeedLabel / SeedValueIfEmpty /
'   SafeGetConfig) that only the legacy entries used.
'   v2.3 install bat -> Setup_admin / Setup_register / Setup_search
'   -> clsSetupOrchestrator.RunFor* is the only live install path.

' ============================================================
' v2.3 Phase A: ASCII Setup_* entries
' ----------------------------------------------------------------
' Adds Setup_admin / Setup_register / Setup_search Public Subs so
' the v2.3 install bat (migration_plan_v1_to_v2_rev2 sect.16) can
' drive setup via Application.Run "Setup_admin" without depending
' on the CP932-encoded names Setup_??? / etc. These wrappers
' delegate to clsSetupOrchestrator (re-use, no logic duplication)
' v2.3 Phase F: MsgBox removed; failures are logged only.
' v2.3 Phase J (DRIFT-iter2, 2026-05-31): legacy v2.1 Setup_*_legacy
'   and SetupCore path were removed (see top header). Only the v2.3
'   ASCII Setup_* path below is live.
' ============================================================

' (SETUP_ROLE_* moved to declarations section near top of module)

' ============================================================
' Public Sub: Setup_admin
' Role: v2.3 install entry for ???.xlsm. Drives the full setup
'       sequence via clsSetupOrchestrator.RunForAdmin:
'         - log sheet bootstrap
'         - config.txt reload
'         - sheet ensure (M-02/M-03/M-04/M-10/M-11/M-12/M-14)
'         - tab color, UI stanza apply, protection
'         - active sheet = M-02
'       NOTE: Sheet list (incl. removal of deprecated M-13 / M-07 /
'             M-09) is governed by clsSetupOrchestrator (the v2.3 SSOT
'             after legacy SHEETS_* constants were removed in Phase J).
'             This wrapper is the stable bat entry regardless.
' Errors: caught -> clsLogger.LogError only (v2.3 Phase F: no MsgBox).
' ============================================================
Public Sub Setup_admin()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1341] modSetup.Setup_admin ENTER"  ' [ADR-0100]
    Debug.Print "[SETUP-ENTER] Setup_admin ts=" & Format$(Now, "hh:nn:ss")  ' [SETUP-DEBUG-PRINT-INJECTED]
    InvokeSetupForRole SETUP_ROLE_ADMIN
    ' 2026-06-07: Seed M-03 C4 with next FmtId if empty (post-Setup hook).
    Debug.Print "[SETUP-SEED] before SeedM03FormatIdIfEmpty"  ' [SETUP-DEBUG-PRINT-INJECTED]
    On Error Resume Next
    Application.Run ThisWorkbook.Name & "!modEntryFormat.SeedM03FormatIdIfEmpty"
    Debug.Print "[SETUP-SEED] after SeedM03FormatIdIfEmpty err=" & Err.Number
    On Error GoTo 0

    ' [Fix-6 follow-up] M-10 storage dirs are now plain cell inputs; the
    ' legacy radios are gone. Clean up leftover radio shapes (idempotent).
    On Error Resume Next
    RemoveLegacyM10Radios
    On Error GoTo 0

    ' [USER-REQ 2026-06-09] Apply TRUE/FALSE dropdown to M-03 searchTarget column (I19:I60)
    On Error Resume Next
    SetupM03SearchTargetValidation
    On Error GoTo 0

    Debug.Print "[SETUP-EXIT] Setup_admin ts=" & Format$(Now, "hh:nn:ss")  ' [SETUP-DEBUG-PRINT-INJECTED]
End Sub

' ============================================================
' Public Sub: RemoveLegacyM10Radios
' Role: Remove leftover legacy M10_RadioDir_* OptionButton shapes from
'       older installs. Storage dirs are now plain cell inputs (Fix-6).
' ============================================================
Public Sub RemoveLegacyM10Radios()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-9001] modSetup.RemoveLegacyM10Radios ENTER"
    On Error GoTo ErrHandler

    Dim wsName As String
    wsName = ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & ChrW(&H8A2D) & ChrW(&H5B9A)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo ErrHandler
    If ws Is Nothing Then Exit Sub

    Dim wasProt As Boolean
    wasProt = ws.ProtectContents
    If wasProt Then ws.Unprotect

    ' Remove existing M10_RadioDir_* shapes (idempotent for re-runs)
    Dim shp As Object
    Dim toDelete As Collection
    Set toDelete = New Collection
    For Each shp In ws.Shapes
        If shp.Name Like "M10_RadioDir_*" Then
            toDelete.Add shp.Name
        End If
    Next shp
    Dim nm As Variant
    For Each nm In toDelete
        ws.Shapes(CStr(nm)).Delete
    Next nm

    If wasProt Then ws.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True

    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-9002] modSetup.RemoveLegacyM10Radios EXIT-OK"
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] RemoveLegacyM10Radios: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Setup_register
' Role: v2.3 install entry for ?o?^?C??.xlsm (skeleton; full
'       behavior is delegated to clsSetupOrchestrator.RunForRegister).
'       Body details are tracked under v2.3 Phase B.
' ============================================================
Public Sub Setup_register()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1342] modSetup.Setup_register ENTER"  ' [ADR-0100]
    Debug.Print "[SETUP-ENTER] Setup_register ts=" & Format$(Now, "hh:nn:ss")  ' [SETUP-DEBUG-PRINT-INJECTED]
    InvokeSetupForRole SETUP_ROLE_REGISTER
    Debug.Print "[SETUP-EXIT] Setup_register ts=" & Format$(Now, "hh:nn:ss")  ' [SETUP-DEBUG-PRINT-INJECTED]
End Sub

' ============================================================
' Public Sub: Setup_search
' Role: v2.3 install entry for ????.xlsm (skeleton; full
'       behavior is delegated to clsSetupOrchestrator.RunForSearch).
'       Body details are tracked under v2.3 Phase B.
' ============================================================
Public Sub Setup_search()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1343] modSetup.Setup_search ENTER"  ' [ADR-0100]
    Debug.Print "[SETUP-ENTER] Setup_search ts=" & Format$(Now, "hh:nn:ss")  ' [SETUP-DEBUG-PRINT-INJECTED]
    InvokeSetupForRole SETUP_ROLE_SEARCH
    Debug.Print "[SETUP-EXIT] Setup_search ts=" & Format$(Now, "hh:nn:ss")  ' [SETUP-DEBUG-PRINT-INJECTED]
End Sub

' ============================================================
' Private Sub: InvokeSetupForRole
' Role: Common dispatcher used by Setup_admin/Setup_register/
'       Setup_search. Instantiates clsSetupOrchestrator once,
'       dispatches by role. v2.3 Phase F: MsgBox removed for
'       headless install safety. Failures are logged only.
' Args:
'   roleName - "admin" | "register" | "search"
' ============================================================
Private Sub InvokeSetupForRole(ByVal roleName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1344] modSetup.InvokeSetupForRole ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    ' v2.3 install hang fix (2026-05-26): ???X?e?b?v??n???O???????
    ' ?? install_admin.bat ?o?H???????????ADebug.Print ??
    ' modCommon.AppendProgressLog ??i?????O???c???Bc:\temp\setup_admin_progress.txt
    ' ?? append ?????B
    Dim stepTs As String
    stepTs = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print stepTs & "InvokeSetupForRole step 0 : enter role=" & roleName
    modCommon.AppendProgressLog stepTs & "InvokeSetupForRole step 0 : enter role=" & roleName

    Debug.Print stepTs & "InvokeSetupForRole step 1 : new clsSetupOrchestrator"
    modCommon.AppendProgressLog stepTs & "InvokeSetupForRole step 1 : new clsSetupOrchestrator"
    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    Debug.Print stepTs & "InvokeSetupForRole step 2 : clsSetupOrchestrator instantiated OK"
    modCommon.AppendProgressLog stepTs & "InvokeSetupForRole step 2 : clsSetupOrchestrator instantiated OK"

    Debug.Print stepTs & "InvokeSetupForRole step 3 : dispatch to RunFor* role=" & roleName
    modCommon.AppendProgressLog stepTs & "InvokeSetupForRole step 3 : dispatch to RunFor* role=" & roleName
    Select Case roleName
        Case SETUP_ROLE_ADMIN
            orch.RunForAdmin
        Case SETUP_ROLE_REGISTER
            orch.RunForRegister
        Case SETUP_ROLE_SEARCH
            orch.RunForSearch
        Case Else
            Err.Raise vbObjectError + 5100, _
                "modSetup.InvokeSetupForRole", _
                "Unknown role: " & roleName
    End Select
    Debug.Print stepTs & "InvokeSetupForRole step 4 : RunFor* returned OK role=" & roleName
    modCommon.AppendProgressLog stepTs & "InvokeSetupForRole step 4 : RunFor* returned OK role=" & roleName
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1345] modSetup.InvokeSetupForRole EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1346] modSetup.InvokeSetupForRole EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[SETUP-ERR] role=" & roleName & " err=" & Err.Number & " " & Err.Description  ' [SETUP-DEBUG-PRINT-INJECTED]
    ' v2.3 Phase F: headless run safety. MsgBox removed; error
    '              notification is done via LOG sheet (clsLogger)
    '              and Debug.Print only. Setup_xxx returns void
    '              but failure is observable in the LOG sheet.
    Dim errTs As String
    errTs = "[setup] [" & Format$(Now(), "hh:nn:ss") & "] "
    Debug.Print errTs & "InvokeSetupForRole ErrHandler : role=" & roleName & " errNum=" & Err.Number & " errDesc=" & Err.Description
    modCommon.AppendProgressLog errTs & "InvokeSetupForRole ErrHandler : role=" & roleName & " errNum=" & Err.Number & " errDesc=" & Err.Description
    LogSetupAdminError roleName, Err.Number, Err.Description
    On Error GoTo 0
End Sub

' ============================================================
' Private Sub: LogSetupAdminError
' Role: Best-effort LogError via clsLogger on LOG sheet.
'       Swallows any logging error (best-effort logging only).
' ============================================================
Private Sub LogSetupAdminError(ByVal roleName As String, _
                                ByVal errNum As Long, _
                                ByVal errDesc As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1347] modSetup.LogSetupAdminError ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim logSheet As Object
    Set logSheet = ThisWorkbook.Worksheets("LOG")
    If Not logSheet Is Nothing Then
        Dim oLogger As clsLogger
        Set oLogger = New clsLogger
        oLogger.Init logSheet
        oLogger.LogError "modSetup", "Setup_" & roleName, _
            "Setup wrapper failed: " & errNum & " " & errDesc, _
            roleName, "SETUPV23-ERR-EE-001"
    End If
    Debug.Print "[ERR] modSetup.Setup_" & roleName & ": " & _
                errNum & " " & errDesc
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1348] modSetup.LogSetupAdminError EXIT-OK"  ' [ADR-0100]
End Sub


' [USER-REQ 2026-06-09] Public Sub: SetupM03SearchTargetValidation
' Applies TRUE/FALSE dropdown validation to M-03 grid column I (searchTarget)
' so the dropdown is visible even before user clicks 読込 / フィールド追加.
Public Sub SetupM03SearchTargetValidation()
    On Error Resume Next
    Dim wsName As String
    wsName = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H8A2D) & ChrW(&H8A08)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(wsName)
    If ws Is Nothing Then Exit Sub
    Dim wasProt As Boolean
    wasProt = ws.ProtectContents
    If wasProt Then ws.Unprotect
    Dim r As Range
    Set r = ws.Range("I19:I60")
    r.Validation.Delete
    r.Validation.Add Type:=xlValidateList, Formula1:="TRUE,FALSE"
    r.Validation.InCellDropdown = True
    r.Validation.IgnoreBlank = True
    ' Also unlock for direct edit
    r.Locked = False
    If wasProt Then ws.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True
End Sub
```
