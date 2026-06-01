---
title: modSetup.bas
description: modSetup.bas のソースコード（コピペ用）
---

# modSetup.bas

**配置先**: `共通モジュール (3 ブック全て)` 用の VBA モジュール  
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\`
- ファイル名: `modSetup.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

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
    InvokeSetupForRole SETUP_ROLE_ADMIN
End Sub

' ============================================================
' Public Sub: Setup_register
' Role: v2.3 install entry for ?o?^?C??.xlsm (skeleton; full
'       behavior is delegated to clsSetupOrchestrator.RunForRegister).
'       Body details are tracked under v2.3 Phase B.
' ============================================================
Public Sub Setup_register()
    InvokeSetupForRole SETUP_ROLE_REGISTER
End Sub

' ============================================================
' Public Sub: Setup_search
' Role: v2.3 install entry for ????.xlsm (skeleton; full
'       behavior is delegated to clsSetupOrchestrator.RunForSearch).
'       Body details are tracked under v2.3 Phase B.
' ============================================================
Public Sub Setup_search()
    InvokeSetupForRole SETUP_ROLE_SEARCH
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
    Exit Sub

ErrHandler:
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
End Sub
```
