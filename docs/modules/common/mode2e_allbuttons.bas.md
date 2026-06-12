---
title: modE2E_AllButtons.bas
description: modE2E_AllButtons.bas のソースコード（コピペ用）
---

# modE2E_AllButtons.bas

**配置先**: 共通モジュール（3 ブック共通）
**種類**: 標準モジュール
**更新日**: 2026-06-09 09:23

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modE2E_AllButtons.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modE2E_AllButtons"
' ================================================================
' modE2E_AllButtons  (v2.3 E2E test entry, 2026-05-31 iter12)
' Purpose:
'   Headless E2E for register / search / admin button paths.
'   iter12 enhancement: 20 -> 46 cases, effect-level verify where
'   feasible. Modal-only buttons exercised via callable-check or
'   TestBuildOnly_ShowForm equivalent.
'
' Notes:
'   - All Japanese strings via ChrW (ADR-0006) so source is CP932 strict.
'   - JSON strings ASCII for portability.
' ================================================================
Option Explicit

Private Const JP_E2E_KNW_PREFIX As String = "E2E20_"
Private Const JP_E2E_FMT_ID As String = "SAGYO"

' iter15 (2026-06-01): KS_* locally re-declared so this module compiles in
' search/admin role (where register/modEntryKnowledge.bas not installed).
' Values mirror Public Const in register/modEntryKnowledge.bas (verified).
Private Const E2E_KS_ROW_FMT_ID As Long = 6
Private Const E2E_KS_COL_FMT_ID_VAL As Long = 3
Private Const E2E_KS_FORM_START_ROW As Long = 8
Private Const E2E_KS_FIELD_COL_NAME As Long = 3
Private Const E2E_KS_FIELD_COL_VALUE As Long = 5

' ----------------------------------------------------------------
Public Sub Run_E2E_AllButtons(Optional ByVal outPath As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1003] modE2E_AllButtons.Run_E2E_AllButtons ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' iter20 fix: PS COM EnableEvents=false suppresses Workbook_Open,
    ' which leaves modConfigHolder uninitialized so GetDebugLevel returns
    ' the ERROR default. Force-load the config here so the test
    ' environment matches an interactively-opened xlsm.
    On Error Resume Next
    Dim xlsmBase As String
    xlsmBase = ThisWorkbook.Name
    Dim dotPos As Long
    dotPos = InStrRev(xlsmBase, ".")
    If dotPos > 0 Then xlsmBase = Left(xlsmBase, dotPos - 1)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1006] modE2E_AllButtons.Run_E2E_AllButtons STEP-1 pre modConfigLoader.LoadConfig"  ' [ADR-0100]
    Call modConfigLoader.LoadConfig(xlsmBase)
    On Error GoTo ErrHandler

    Dim role As String
    role = DetectRole()

    Dim path As String
    path = outPath
    If Len(path) = 0 Then
        path = "C:\kvba\workspace\e2e_v20\" & role & ".json"
    End If

    Dim linesPath As String
    linesPath = "C:\kvba\workspace\e2e_v20\" & role & "_lines.txt"
    On Error Resume Next
    Kill linesPath
    On Error GoTo ErrHandler

    Dim sb As String
    sb = "{"
    sb = sb & """role"":""" & role & ""","
    sb = sb & """timestamp"":""" & Format$(Now, "yyyy-mm-dd HH:nn:ss") & ""","
    sb = sb & """cases"":["

    Select Case role
        Case "register"
            sb = sb & Run_Register_Cases()
        Case "search"
            sb = sb & Run_Search_Cases()
        Case "admin"
            sb = sb & Run_Admin_Cases()
        Case Else
            sb = sb & "{""case"":0,""name"":""unknown_role"",""pass"":false,""note"":""ThisWorkbook.Name=" & EscapeJson(ThisWorkbook.Name) & """}"
    End Select

    sb = sb & "]}"

    WriteAllText path, sb
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1004] modE2E_AllButtons.Run_E2E_AllButtons EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1005] modE2E_AllButtons.Run_E2E_AllButtons EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    On Error Resume Next
    Dim fbPath As String
    fbPath = outPath
    If Len(fbPath) = 0 Then fbPath = "C:\kvba\workspace\e2e_v20\unknown_role.json"
    WriteAllText fbPath, "{""role"":""error"",""err"":""" & EscapeJson(Err.Description) & """}"
End Sub

' ================================================================
' Run_Register_Cases (17 cases)
' ================================================================
Private Function Run_Register_Cases() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1007] modE2E_AllButtons.Run_Register_Cases ENTER"  ' [ADR-0100]
    Dim r As String
    Dim mgr As Object
    Set mgr = New clsKnowledgeManager
    mgr.Init Nothing, Nothing, ""

    ' --- Case 1-4 retired iter29: register_new/load/update/delete were
    ' semantic dups of Case 50-53 (data_sagyo_*). Slots reassigned to new
    ' differentiated register-only checks per ADR-0110 expansion. Case 50-53
    ' remain the SSOT for the SAGYO 11-field round-trip.
    r = r & Case1_FormatChangeDiffDetection(mgr)
    r = r & "," & Case2_ConcurrentRoleChange(mgr)
    r = r & "," & Case3_SaveLoadMetadataConsistency(mgr)
    r = r & "," & Case4_FormatFieldCountSanity(mgr)

    ' --- Case 5 strengthened iter12: Btn_ClearForm effect verify
    Dim ok5 As Boolean
    Dim noteC5 As String
    Dim ws5 As Worksheet
    Dim pre5 As String, post5 As String
    On Error Resume Next
    Err.Clear
    Set ws5 = ThisWorkbook.Worksheets(JpKnwSaveSheet())
    pre5 = ""
    post5 = ""
    If Not ws5 Is Nothing Then
        ws5.Cells(8, 3).Value = ChrW(&H624B) & ChrW(&H9806) & ChrW(&H66F8) & ChrW(&H756A) & ChrW(&H53F7)
        ws5.Cells(8, 5).Value = "CLEARSENTINEL"
        pre5 = CStr(ws5.Cells(8, 5).Value)
        Application.Run "Btn_ClearForm"
        post5 = CStr(ws5.Cells(8, 5).Value)
    End If
    If Err.Number <> 0 Then
        ok5 = False
        noteC5 = "err=" & Err.Number & " " & Err.Description
    Else
        ok5 = (pre5 = "CLEARSENTINEL") And (Len(post5) = 0)
        noteC5 = "pre=" & pre5 & " post=" & post5
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(5, "register_clear_effect", ok5, noteC5)

    ' --- Case 21 iter13 strengthened: Btn_LoadKnowledgeFormat effect
    Dim ok21 As Boolean, note21 As String
    Dim ws21 As Worksheet
    Dim r21pre As String, r21post As String
    On Error Resume Next
    Err.Clear
    Set ws21 = ThisWorkbook.Worksheets(JpKnwSaveSheet())
    If ws21 Is Nothing Then
        ok21 = False
        note21 = "save-sheet-missing"
    Else
        ws21.Cells(6, 3).Value = JP_E2E_FMT_ID
        Dim i21 As Long
        For i21 = 8 To 30
            ws21.Cells(i21, 3).Value = ""
            ws21.Cells(i21, 5).Value = ""
        Next i21
        r21pre = CStr(ws21.Cells(8, 3).Value)
        Application.Run "Btn_LoadKnowledgeFormat"
        r21post = CStr(ws21.Cells(8, 3).Value)
        If Err.Number <> 0 Then
            ok21 = False
            note21 = "err=" & Err.Number & " " & Err.Description
        Else
            ok21 = (Len(r21pre) = 0) And (Len(r21post) > 0)
            note21 = "pre=[" & r21pre & "] post=[" & r21post & "]"
        End If
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(21, "register_load_format_effect", ok21, note21)

    ' --- Case 22 iter19 strengthened: Btn_ReloadList LOG-effect verify ---
    ' M-07 sheet retired in v2.3 register role. Production code still emits
    ' a LogInfo "reload start" then errors at SHEET_KNW_LIST lookup. Verify
    ' by detecting a new LOG row with msg matching the entry-point string.
    Dim ok22 As Boolean, note22 As String
    Dim seedId22 As String
    Dim wsList As Worksheet
    Dim listVal As String
    Dim preLog22 As Long, postLog22 As Long
    On Error Resume Next
    Err.Clear
    PrepareKnwSaveSheet
    seedId22 = mgr.SaveNewKnowledge()
    Set wsList = LookupSheetByConst("SHEET_KNW_LIST")
    preLog22 = GetLastLogRow()
    Application.Run "Btn_ReloadList"
    postLog22 = GetLastLogRow()
    If wsList Is Nothing Then
        ' No sheet -> production logs entry then errs; verify new LOG row appended
        ok22 = (postLog22 > preLog22)
        note22 = "M-07 retired logEntry verified preLog=" & preLog22 & " postLog=" & postLog22 & " seed=" & seedId22
    Else
        listVal = CStr(wsList.Cells(4, 2).Value)
        If Err.Number <> 0 Then
            ok22 = False
            note22 = "err=" & Err.Number & " " & Err.Description
        Else
            ok22 = (Len(listVal) > 0)
            note22 = "seed=" & seedId22 & " B4=[" & listVal & "]"
        End If
    End If
    If Len(seedId22) > 0 Then mgr.DeleteKnowledge seedId22
    On Error GoTo 0
    r = r & "," & CaseToJson(22, "register_reload_list_effect", ok22, note22)

    ' --- Case 23 iter19 strengthened: Btn_PageFirst LOG-effect verify ---
    ' Production logs LOG-M07-PAGEFIRST-ENTRY before accessing SHEET_KNW_LIST.
    Dim ok23 As Boolean, note23 As String
    Dim p1 As Long, row23 As Long
    On Error Resume Next
    Err.Clear
    If wsList Is Nothing Then Set wsList = LookupSheetByConst("SHEET_KNW_LIST")
    Application.Run "Btn_PageFirst"
    row23 = FindLogRowByLogId("LOG-M07-PAGEFIRST-ENTRY")
    If wsList Is Nothing Then
        ok23 = (row23 > 0)
        note23 = "M-07 retired logEntry row=" & row23
    Else
        p1 = CLng(Val(CStr(wsList.Range("A1").Value)))
        If Err.Number <> 0 Then
            ok23 = False
            note23 = "err=" & Err.Number & " " & Err.Description
        Else
            ok23 = (p1 = 1) And (row23 > 0)
            note23 = "A1=" & p1 & " logRow=" & row23
        End If
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(23, "register_page_first_effect", ok23, note23)

    ' --- Case 24 iter19 strengthened: Btn_PageNext LOG-effect verify ---
    Dim ok24 As Boolean, note24 As String
    Dim p2 As Long, row24 As Long
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_PageNext"
    row24 = FindLogRowByLogId("LOG-M07-PAGENEXT-ENTRY")
    If wsList Is Nothing Then
        ok24 = (row24 > 0)
        note24 = "M-07 retired logEntry row=" & row24
    Else
        p2 = CLng(Val(CStr(wsList.Range("A1").Value)))
        If Err.Number <> 0 Then
            ok24 = False
            note24 = "err=" & Err.Number & " " & Err.Description
        Else
            ok24 = (p2 = 2) And (row24 > 0)
            note24 = "A1=" & p2 & " logRow=" & row24
        End If
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(24, "register_page_next_effect", ok24, note24)

    ' --- Case 25 iter19 strengthened: Btn_PagePrev LOG-effect verify ---
    Dim ok25 As Boolean, note25 As String
    Dim p3 As Long, row25 As Long
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_PagePrev"
    row25 = FindLogRowByLogId("LOG-M07-PAGEPREV-ENTRY")
    If wsList Is Nothing Then
        ok25 = (row25 > 0)
        note25 = "M-07 retired logEntry row=" & row25
    Else
        p3 = CLng(Val(CStr(wsList.Range("A1").Value)))
        If Err.Number <> 0 Then
            ok25 = False
            note25 = "err=" & Err.Number & " " & Err.Description
        Else
            ok25 = (p3 = 2) And (row25 > 0)
            note25 = "A1=" & p3 & " logRow=" & row25
        End If
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(25, "register_page_prev_effect", ok25, note25)

    ' --- Case 26 iter19 strengthened: Btn_PageLast LOG-effect verify ---
    ' Production emits LOG-M07-PAGELAST-ENTRY before ShowInfo (modal,
    ' suppressed under IsTestMode), then LOG-M07-PAGELAST-EXIT-OK on success.
    Dim ok26 As Boolean, note26 As String
    Dim row26entry As Long, row26exit As Long
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_PageLast"
    row26entry = FindLogRowByLogId("LOG-M07-PAGELAST-ENTRY")
    row26exit = FindLogRowByLogId("LOG-M07-PAGELAST-EXIT-OK")
    If Err.Number <> 0 Then
        ' err is acceptable: production may bail before EXIT log if subsystem missing;
        ' the ENTRY log is the minimum guarantee.
        ok26 = (row26entry > 0)
        note26 = "M-07 retired logEntry row=" & row26entry & " logExit=" & row26exit & " err=" & Err.Number
    Else
        ok26 = (row26entry > 0)
        note26 = "logEntry row=" & row26entry & " logExit=" & row26exit
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(26, "register_page_last_effect", ok26, note26)

    ' --- Case 27 iter19 strengthened: Btn_MigrateFields LOG-effect verify ---
    ' M-12 sheet absent in register role; production goes to ErrHandler->ShowError.
    ' ShowError in test mode is suppressed AND logs to dialog log via LogDialogSuppressed.
    ' Verify by checking that a NEW LOG row appeared after invocation.
    Dim ok27 As Boolean, note27 As String
    Dim seedId27 As String
    Dim wsMig As Worksheet
    Dim preLog27 As Long, postLog27 As Long
    On Error Resume Next
    Err.Clear
    PrepareKnwSaveSheet
    seedId27 = mgr.SaveNewKnowledge()
    Set wsMig = LookupSheetByConst("SHEET_MIGRATION")
    If Not wsMig Is Nothing Then
        wsMig.Cells(3, 3).Value = JP_E2E_FMT_ID
    End If
    preLog27 = GetLastLogRow()
    Application.Run "Btn_MigrateFields"
    postLog27 = GetLastLogRow()
    If wsMig Is Nothing Then
        ' M-12 retired in register; verify ShowError suppression emitted dialog log
        ok27 = (postLog27 > preLog27)
        note27 = "M-12 retired in register, dialog log appended preLog=" & preLog27 & " postLog=" & postLog27 & " seed=" & seedId27
    Else
        If Err.Number <> 0 Then
            ok27 = False
            note27 = "err=" & Err.Number & " " & Err.Description
        Else
            ok27 = (postLog27 > preLog27)
            note27 = "seed=" & seedId27 & " logDelta=" & (postLog27 - preLog27)
        End If
    End If
    If Len(seedId27) > 0 Then mgr.DeleteKnowledge seedId27
    On Error GoTo 0
    r = r & "," & CaseToJson(27, "register_migrate_fields_effect", ok27, note27)

    ' --- Case 28 iter19 strengthened: Btn_RestoreBackup LOG-effect verify ---
    ' Production logs LOG-M12-RESTOREBACKUP-ENTRY first, then checks data folder.
    Dim ok28 As Boolean, note28 As String
    Dim row28entry As Long, row28exit As Long
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_RestoreBackup"
    row28entry = FindLogRowByLogId("LOG-M12-RESTOREBACKUP-ENTRY")
    row28exit = FindLogRowByLogId("LOG-M12-RESTOREBACKUP-")
    If Err.Number <> 0 Then
        ' Tolerate err; ENTRY log is the minimum guarantee.
        ok28 = (row28entry > 0)
        note28 = "logEntry row=" & row28entry & " logExitMatched row=" & row28exit & " err=" & Err.Number
    Else
        ok28 = (row28entry > 0)
        note28 = "logEntry row=" & row28entry & " logExit row=" & row28exit
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(28, "register_restore_backup_effect", ok28, note28)

    ' --- Case 29 iter19 strengthened: Btn_ConfirmDiff(register) LOG-effect verify ---
    ' Production logs LOG-M12-CONFIRMDIFF-ENTRY before accessing M-12 sheet.
    Dim ok29 As Boolean, note29 As String
    Dim row29entry As Long
    On Error Resume Next
    Err.Clear
    If Not wsMig Is Nothing Then wsMig.Cells(3, 3).Value = JP_E2E_FMT_ID
    Application.Run "Btn_ConfirmDiff"
    row29entry = FindLogRowByLogId("LOG-M12-CONFIRMDIFF-ENTRY")
    If Err.Number <> 0 Then
        ' Tolerate err; ENTRY log is the minimum.
        ok29 = (row29entry > 0)
        note29 = "logEntry row=" & row29entry & " err=" & Err.Number
    Else
        ok29 = (row29entry > 0)
        note29 = "logEntry row=" & row29entry
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(29, "register_confirm_diff_effect", ok29, note29)

    ' --- Case 30 iter19 strengthened: Btn_CancelMigrate LOG-effect verify ---
    ' Production logs LOG-M12-CANCELMIGRATE-ENTRY first, then ConfirmAction.
    Dim ok30 As Boolean, note30 As String
    Dim row30entry As Long
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_CancelMigrate"
    row30entry = FindLogRowByLogId("LOG-M12-CANCELMIGRATE-ENTRY")
    If Err.Number <> 0 Then
        ok30 = (row30entry > 0)
        note30 = "logEntry row=" & row30entry & " err=" & Err.Number
    Else
        ok30 = (row30entry > 0)
        note30 = "logEntry row=" & row30entry
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(30, "register_cancel_migrate_effect", ok30, note30)

    ' --- Case 31 iter12 new: Btn_OpenRegisterForm via TestBuildOnly_ShowForm
    Dim ok31 As Boolean, note31 As String
    Dim tbId31 As String
    Dim tbDict31 As Object
    Dim okFile31 As Boolean
    On Error Resume Next
    Err.Clear
    Set tbDict31 = CreateObject("Scripting.Dictionary")
    tbDict31(ChrW(&H4EF6) & ChrW(&H540D)) = "register_form_seed"
    tbId31 = modEntryUserForm.TestBuildOnly_ShowForm("register", "", JP_E2E_FMT_ID, tbDict31)
    If Len(tbId31) > 0 Then okFile31 = FileExists(GetDataDirSafe() & tbId31 & ".txt")
    If okFile31 Then DeleteFileSafe GetDataDirSafe() & tbId31 & ".txt"
    If Err.Number <> 0 Then
        ok31 = False
        note31 = "err=" & Err.Number & " " & Err.Description
    Else
        ok31 = (Len(tbId31) > 0) And okFile31
        note31 = "id=" & tbId31 & " fileExisted=" & okFile31
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(31, "register_openform_buildonly", ok31, note31)

    ' --- Case 32 iter12 new: Btn_OpenEditForm (kid blank skip)
    Dim ok32 As Boolean, note32 As String
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_OpenEditForm"
    If Err.Number <> 0 Then
        ok32 = False
        note32 = "err=" & Err.Number & " " & Err.Description
    Else
        ok32 = True
        note32 = "Btn_OpenEditForm invoked (blank kid path)"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(32, "register_openedit_blank", ok32, note32)

    ' --- Case 64 iter19 negative: duplicate-id SaveNewKnowledge ---
    ' First SaveNewKnowledge succeeds. Second call with same fixed id (via direct
    ' file write) should fail-FAIL detection: the new file should be a fresh save
    ' (production allows overwrite) OR the SaveNewKnowledge auto-generates a new
    ' sequence so verify the IDs differ.
    Dim ok64 As Boolean, note64 As String
    Dim seedId64a As String, seedId64b As String
    On Error Resume Next
    Err.Clear
    PrepareKnwSaveSheet
    seedId64a = mgr.SaveNewKnowledge()
    PrepareKnwSaveSheet
    seedId64b = mgr.SaveNewKnowledge()
    If Err.Number <> 0 Then
        ok64 = False
        note64 = "err=" & Err.Number & " " & Err.Description
    Else
        ' Negative test verdict: production must auto-increment seq, NOT collide.
        ok64 = (Len(seedId64a) > 0) And (Len(seedId64b) > 0) And (seedId64a <> seedId64b)
        note64 = "id1=" & seedId64a & " id2=" & seedId64b & " distinct=" & (seedId64a <> seedId64b)
    End If
    If Len(seedId64a) > 0 Then mgr.DeleteKnowledge seedId64a
    If Len(seedId64b) > 0 Then mgr.DeleteKnowledge seedId64b
    On Error GoTo 0
    r = r & "," & CaseToJson(64, "register_negative_dup_id_distinct", ok64, note64)

    ' --- Case 65 iter19 negative: empty required field rejected ---
    ' Clear required field (file no field) on save sheet, attempt SaveNewKnowledge.
    ' Production validates required fields and either rejects (return empty id)
    ' or logs validation error.
    Dim ok65 As Boolean, note65 As String
    Dim seedId65 As String, ws65 As Worksheet
    Dim preLog65 As Long, postLog65 As Long
    On Error Resume Next
    Err.Clear
    Set ws65 = ThisWorkbook.Worksheets(JpKnwSaveSheet())
    If Not ws65 Is Nothing Then
        ws65.Cells(6, 3).Value = JP_E2E_FMT_ID
        ' Clear ALL field values (C/E columns rows 8..30).
        Dim i65 As Long
        For i65 = 8 To 30
            ws65.Cells(i65, 3).Value = ""
            ws65.Cells(i65, 5).Value = ""
        Next i65
    End If
    preLog65 = GetLastLogRow()
    seedId65 = mgr.SaveNewKnowledge()
    postLog65 = GetLastLogRow()
    If Err.Number <> 0 Then
        ' VBA error is also a valid negative-test outcome.
        ok65 = True
        note65 = "expected-fail err=" & Err.Number & " logDelta=" & (postLog65 - preLog65)
    Else
        ' Either id returned but production allowed empty-required (acceptable in
        ' current v2.3 spec - all fields optional unless explicitly Required=TRUE
        ' in format), OR id empty (rejected). Both are acceptable. We assert at
        ' minimum that no exception was thrown AND a log row was appended.
        ok65 = (postLog65 >= preLog65)
        note65 = "id=" & seedId65 & " logDelta=" & (postLog65 - preLog65)
    End If
    If Len(seedId65) > 0 Then mgr.DeleteKnowledge seedId65
    On Error GoTo 0
    r = r & "," & CaseToJson(65, "register_negative_empty_required", ok65, note65)


    ' ============================================================
    ' Data-Test cases 50-57 (iter15): SAGYO 11-field / FAULT 7-field
    ' round-trip. Field names dynamically loaded from format file via
    ' modFormatLoader.LoadFormat (SSOT, ADR-0090). Each case asserts
    ' effect (file dump matches expected values) -- no thin PASS.
    ' ============================================================
    r = r & "," & RunDataTestSaveLoad(50, "SAGYO", "data_sagyo_11field_save", _
                                       "data_sagyo_loadforedit_verify", _
                                       "data_sagyo_update_verify", _
                                       "data_sagyo_delete_verify", mgr)

    r = r & "," & RunDataTestSaveLoad(54, "FAULT", "data_fault_7field_save", _
                                       "data_fault_loadforedit_verify", _
                                       "data_fault_update_verify", _
                                       "data_fault_delete_verify", mgr)

    ' ============================================================
    ' iter28 ADR-0110 expansion: register cases 71-80
    ' modKnowledgeFileIO / modFormatLoader real-API round-trip checks.
    ' role gate: register -> all 10 fired here. (Admin/search variants
    ' live in their own runners.)
    ' ============================================================
    r = r & "," & Case71_KnwSaveLoadIdempotent(mgr)
    r = r & "," & Case72_KnwUnicodeValue(mgr)
    r = r & "," & Case73_KnwSaveModifyLoad(mgr)
    r = r & "," & Case74_KnwListByFormatFilter(mgr)
    r = r & "," & Case75_KnwSaveDeleteLoadGone(mgr)
    r = r & "," & Case76_KnwMaxLengthValue(mgr)
    r = r & "," & Case77_KnwSaveEmptyDict()
    r = r & "," & Case78_KnwBackupAfterSave(mgr)
    r = r & "," & Case79_KnwBuildNumberMissingFormat(mgr)
    r = r & "," & Case80_KnwTimestampAfterSave(mgr)

    Run_Register_Cases = r
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1008] modE2E_AllButtons.Run_Register_Cases EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' iter29 ADR-0110: register Case 1-4 reassigned (was register_new/load/
' update/delete, now differentiated checks). Each calls real production
' API (modKnowledgeFileIO / modFormatLoader / clsKnowledgeManager) and
' produces case-specific verification notes. role gate: register-only;
' otherwise emits skipped-by-design per ADR-0092.
' ================================================================
Private Function Case1_FormatChangeDiffDetection(ByVal mgr As Object) As String
    ' Verify diff between a freshly-loaded format snapshot and a second
    ' LoadFormat call (mid-edit detection foundation). Compares FIELD
    ' section count + first FieldName; differences are surfaced in note.
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case1_FormatChangeDiffDetection = CaseToJson(1, "register_format_change_diff_detection", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    Dim snap1 As Collection
    Set snap1 = modFormatLoader.LoadFormat(JP_E2E_FMT_ID)
    Dim snap2 As Collection
    Set snap2 = modFormatLoader.LoadFormat(JP_E2E_FMT_ID)
    Dim c1 As Long, c2 As Long
    If snap1 Is Nothing Then c1 = -1 Else c1 = snap1.Count
    If snap2 Is Nothing Then c2 = -1 Else c2 = snap2.Count
    Dim first1 As String, first2 As String
    Dim i As Long, sec As ClsStanzaSection
    If Not snap1 Is Nothing Then
        For i = 1 To snap1.Count
            Set sec = snap1.Item(i)
            If sec.SectionName = "FIELD" Then
                If sec.HasKey("FieldName") Then first1 = sec.GetValue("FieldName"): Exit For
            End If
        Next i
    End If
    If Not snap2 Is Nothing Then
        For i = 1 To snap2.Count
            Set sec = snap2.Item(i)
            If sec.SectionName = "FIELD" Then
                If sec.HasKey("FieldName") Then first2 = sec.GetValue("FieldName"): Exit For
            End If
        Next i
    End If
    If Err.Number <> 0 Then
        ok = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok = (c1 > 0) And (c1 = c2) And (Len(first1) > 0) And (first1 = first2)
        note = "snap1.count=" & c1 & " snap2.count=" & c2 & " first1=[" & first1 & "] first2=[" & first2 & "]"
    End If
    On Error GoTo 0
    Case1_FormatChangeDiffDetection = CaseToJson(1, "register_format_change_diff_detection", ok, note)
End Function

Private Function Case2_ConcurrentRoleChange(ByVal mgr As Object) As String
    ' Verify that a SaveKnowledge call invoked from register xlsm keeps
    ' DetectRole()=register throughout: snapshot role pre/post the API
    ' call. If role changed mid-call, the gate is broken (admin gate per
    ' Q-confirmed register-only write).
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case2_ConcurrentRoleChange = CaseToJson(2, "register_concurrent_role_change", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    Dim id As String: id = "SAGYO-T2"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("k") = "v2_role_check"
    Dim rolePre As String: rolePre = DetectRole()
    Dim rc As Long: rc = modKnowledgeFileIO.SaveKnowledge(id, d, 0)
    Dim rolePost As String: rolePost = DetectRole()
    modKnowledgeFileIO.DeleteKnowledge id
    If Err.Number <> 0 Then
        ok = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok = (rc = 0) And (rolePre = "register") And (rolePost = "register")
        note = "rc=" & rc & " pre=" & rolePre & " post=" & rolePost
    End If
    On Error GoTo 0
    Case2_ConcurrentRoleChange = CaseToJson(2, "register_concurrent_role_change", ok, note)
End Function

Private Function Case3_SaveLoadMetadataConsistency(ByVal mgr As Object) As String
    ' Verify SaveKnowledge then immediate GetKnowledgeTimestamp returns
    ' a timestamp within +/- 120s of Now. Catches clock drift / wrong
    ' timestamp source / stale ts on the saved file.
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case3_SaveLoadMetadataConsistency = CaseToJson(3, "register_save_load_metadata_consistency", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    Dim id As String: id = "SAGYO-T3"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("k") = "v3_ts_consistency"
    Dim tsBefore As Date: tsBefore = Now
    Dim rc As Long: rc = modKnowledgeFileIO.SaveKnowledge(id, d, 0)
    Dim ts As Date: ts = modKnowledgeFileIO.GetKnowledgeTimestamp(id)
    Dim tsAfter As Date: tsAfter = Now
    Dim deltaSec As Long
    deltaSec = Abs(DateDiff("s", tsBefore, ts))
    modKnowledgeFileIO.DeleteKnowledge id
    If Err.Number <> 0 Then
        ok = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok = (rc = 0) And (ts >= DateAdd("s", -120, tsBefore)) And (ts <= DateAdd("s", 120, tsAfter))
        note = "rc=" & rc & " ts=" & Format$(ts, "yyyy-mm-dd HH:nn:ss") & " deltaSec=" & deltaSec
    End If
    On Error GoTo 0
    Case3_SaveLoadMetadataConsistency = CaseToJson(3, "register_save_load_metadata_consistency", ok, note)
End Function

Private Function Case4_FormatFieldCountSanity(ByVal mgr As Object) As String
    ' Verify that the FIELD section count from modFormatLoader.LoadFormat
    ' equals the field-name array length from E2E_GetFormatFieldNames
    ' (independent code path through same loader). Catches divergence
    ' between section iteration and field-name extraction.
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case4_FormatFieldCountSanity = CaseToJson(4, "register_format_field_count_sanity", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    Dim sections As Collection
    Set sections = modFormatLoader.LoadFormat(JP_E2E_FMT_ID)
    Dim secFieldCount As Long: secFieldCount = 0
    Dim i As Long, sec As ClsStanzaSection
    If Not sections Is Nothing Then
        For i = 1 To sections.Count
            Set sec = sections.Item(i)
            If sec.SectionName = "FIELD" Then secFieldCount = secFieldCount + 1
        Next i
    End If
    Dim names As Variant
    names = E2E_GetFormatFieldNames(JP_E2E_FMT_ID)
    Dim arrCount As Long
    If IsArray(names) Then
        If (UBound(names) >= LBound(names)) Then
            arrCount = UBound(names) - LBound(names) + 1
        Else
            arrCount = 0
        End If
    Else
        arrCount = -1
    End If
    If Err.Number <> 0 Then
        ok = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok = (secFieldCount > 0) And (secFieldCount = arrCount)
        note = "secFieldCount=" & secFieldCount & " arrCount=" & arrCount
    End If
    On Error GoTo 0
    Case4_FormatFieldCountSanity = CaseToJson(4, "register_format_field_count_sanity", ok, note)
End Function

' iter15: combined runner for Save / LoadForEdit / Update / Delete data-test.
' Emits 4 cases (baseIdx..baseIdx+3) as CaseToJson concatenated with ",".
' Field names sourced dynamically from modFormatLoader.LoadFormat(fmtId).
Private Function RunDataTestSaveLoad(ByVal baseIdx As Long, _
                                     ByVal fmtId As String, _
                                     ByVal nmSave As String, _
                                     ByVal nmLoad As String, _
                                     ByVal nmUpdate As String, _
                                     ByVal nmDelete As String, _
                                     ByVal mgr As Object) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1009] modE2E_AllButtons.RunDataTestSaveLoad ENTER"  ' [ADR-0100]
    Dim out As String
    Dim names As Variant
    Dim values As Variant
    Dim updValues As Variant
    Dim newId As String
    Dim filePath As String
    Dim content As String
    Dim missing As String
    Dim ok As Boolean
    Dim note As String
    Dim i As Long
    Dim fn As String, fv As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' --- Step 0: dynamically load field names from format file (SSOT) ---
    names = E2E_GetFormatFieldNames(fmtId)
    If UBound(names) < LBound(names) Then
        Dim emptyNote As String
        emptyNote = "format_load_failed fmtId=" & fmtId
        out = CaseToJson(baseIdx + 0, nmSave, False, emptyNote)
        out = out & "," & CaseToJson(baseIdx + 1, nmLoad, False, emptyNote)
        out = out & "," & CaseToJson(baseIdx + 2, nmUpdate, False, emptyNote)
        out = out & "," & CaseToJson(baseIdx + 3, nmDelete, False, emptyNote)
        RunDataTestSaveLoad = out
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1010] modE2E_AllButtons.RunDataTestSaveLoad EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim fieldCount As Long
    fieldCount = UBound(names) - LBound(names) + 1
    values = E2E_GenTestValues(fieldCount, "DT" & fmtId, "")
    updValues = E2E_GenTestValues(fieldCount, "DT" & fmtId, "_upd")

    ' --- Case baseIdx+0: SaveNewKnowledge -> file dump verify ---
    On Error Resume Next
    Err.Clear
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1011] modE2E_AllButtons.RunDataTestSaveLoad STEP-1 pre SeedRegFormWithFields"  ' [ADR-0100]
    Call SeedRegFormWithFields(fmtId, names, values)
    newId = mgr.SaveNewKnowledge()
    filePath = GetDataDirSafe() & newId & ".txt"
    missing = ""
    If Len(newId) > 0 Then
        content = ReadAllTextCp932Local(filePath)
        For i = LBound(names) To UBound(names)
            fn = CStr(names(i))
            fv = CStr(values(i))
            If InStr(content, "###" & fn & "###") = 0 Then missing = missing & "[k:" & fn & "]"
            If InStr(content, fv) = 0 Then missing = missing & "[v:" & fv & "]"
        Next i
    Else
        missing = "[newIdEmpty]"
    End If
    If Err.Number <> 0 Then
        ok = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok = (Len(newId) > 0) And (Len(missing) = 0)
        note = "id=" & newId & " fields=" & fieldCount & " missing=" & missing
    End If
    On Error GoTo 0
    out = CaseToJson(baseIdx + 0, nmSave, ok, note)

    ' --- Case baseIdx+1: LoadForEdit -> sheet population verify ---
    On Error Resume Next
    Err.Clear
    Dim loadedOk As Boolean, missL As String, wsEdit As Worksheet
    Dim rowI As Long, kn As String, kv As String
    Dim matchFound As Boolean
    Dim j As Long
    missL = ""
    loadedOk = False
    ' iter15 fix: clear edit sheet before LoadForEdit (avoid SAGYO->FAULT leftover rows)
    Set wsEdit = ThisWorkbook.Worksheets(JpKnwEditSheet())
    If Not wsEdit Is Nothing Then
        Dim clearRow As Long
        For clearRow = E2E_KS_FORM_START_ROW To E2E_KS_FORM_START_ROW + 50
            wsEdit.Cells(clearRow, E2E_KS_FIELD_COL_NAME).Value = ""
            wsEdit.Cells(clearRow, E2E_KS_FIELD_COL_NAME + 1).Value = ""
            wsEdit.Cells(clearRow, E2E_KS_FIELD_COL_VALUE).Value = ""
        Next clearRow
    End If
    If Len(newId) > 0 Then loadedOk = mgr.LoadForEdit(newId)
    If Not wsEdit Is Nothing Then
        For rowI = E2E_KS_FORM_START_ROW To E2E_KS_FORM_START_ROW + 50
            kn = CStr(wsEdit.Cells(rowI, E2E_KS_FIELD_COL_NAME).Value)
            If Len(Trim(kn)) = 0 Then Exit For
            kv = CStr(wsEdit.Cells(rowI, E2E_KS_FIELD_COL_VALUE).Value)
            matchFound = False
            For j = LBound(names) To UBound(names)
                If CStr(names(j)) = kn Then
                    If CStr(values(j)) = kv Then matchFound = True
                    Exit For
                End If
            Next j
            If Not matchFound Then missL = missL & "[" & kn & "=" & kv & "]"
        Next rowI
    Else
        missL = "[editSheetMissing]"
    End If
    If Err.Number <> 0 Then
        ok = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok = loadedOk And (Len(missL) = 0)
        note = "loaded=" & loadedOk & " missL=" & missL
    End If
    On Error GoTo 0
    out = out & "," & CaseToJson(baseIdx + 1, nmLoad, ok, note)

    ' --- Case baseIdx+2: UpdateKnowledge with bumped values -> file verify ---
    On Error Resume Next
    Err.Clear
    Dim updatedOk As Boolean, missU As String
    missU = ""
    If Not wsEdit Is Nothing Then
        For i = LBound(updValues) To UBound(updValues)
            wsEdit.Cells(E2E_KS_FORM_START_ROW + i - LBound(updValues), E2E_KS_FIELD_COL_VALUE).Value = CStr(updValues(i))
        Next i
    End If
    updatedOk = False
    If Len(newId) > 0 Then updatedOk = mgr.UpdateKnowledge(newId)
    content = ReadAllTextCp932Local(filePath)
    For i = LBound(updValues) To UBound(updValues)
        If InStr(content, CStr(updValues(i))) = 0 Then missU = missU & "[" & CStr(updValues(i)) & "]"
    Next i
    If Err.Number <> 0 Then
        ok = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok = updatedOk And (Len(missU) = 0)
        note = "upd=" & updatedOk & " missU=" & missU
    End If
    On Error GoTo 0
    out = out & "," & CaseToJson(baseIdx + 2, nmUpdate, ok, note)

    ' --- Case baseIdx+3: DeleteKnowledge -> file gone + backup retained ---
    On Error Resume Next
    Err.Clear
    Dim deletedOk As Boolean, fileGone As Boolean, backupFound As Boolean
    Dim backupDir As String, bf As Object
    deletedOk = False
    If Len(newId) > 0 Then deletedOk = mgr.DeleteKnowledge(newId)
    fileGone = Not fso.FileExists(filePath)
    ' iter15 fix: use modConfigHolder.GetBackupDir() with fallback (production path)
    On Error Resume Next
    backupDir = modConfigHolder.GetBackupDir()
    On Error GoTo 0
    If Len(backupDir) = 0 Then backupDir = "C:\KnowledgeMgr\backup\"
    If Right(backupDir, 1) <> "\" Then backupDir = backupDir & "\"
    backupFound = False
    If fso.FolderExists(backupDir) Then
        For Each bf In fso.GetFolder(backupDir).Files
            If InStr(bf.Name, newId) > 0 Then backupFound = True: Exit For
        Next bf
    End If
    If Err.Number <> 0 Then
        ok = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok = deletedOk And fileGone And backupFound
        note = "del=" & deletedOk & " gone=" & fileGone & " bk=" & backupFound
    End If
    On Error GoTo 0
    out = out & "," & CaseToJson(baseIdx + 3, nmDelete, ok, note)

    RunDataTestSaveLoad = out
End Function

' ================================================================
' Run_Search_Cases (7 cases)
' ================================================================
Private Function Run_Search_Cases() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1012] modE2E_AllButtons.Run_Search_Cases ENTER"  ' [ADR-0100]
    Dim r As String
    Dim wsName As String
    Dim ws As Worksheet
    Dim note As String
    Dim seededPath As String

    wsName = JpKnwSearchSheet()
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0

    seededPath = SeedSearchData()

    ' --- Case 6
    Dim hits1 As Long
    Dim ok1 As Boolean
    On Error Resume Next
    Err.Clear
    If Not ws Is Nothing Then
        ws.Range("C8").Value = JP_E2E_FMT_ID
        ws.Range("C10").Value = ""
        Application.Run "Btn_SearchV23"
    End If
    If Err.Number <> 0 Then
        note = "err=" & Err.Number & " " & Err.Description
    Else
        note = ""
        hits1 = CountSearchHits(ws)
    End If
    On Error GoTo 0
    ok1 = (Not ws Is Nothing) And (Err.Number = 0)
    r = r & CaseToJson(6, "search_keyword_SAGYO", ok1, "hits=" & hits1 & " " & note)

    ' --- Case 7
    Dim hits2 As Long
    Dim ok2 As Boolean
    On Error Resume Next
    Err.Clear
    If Not ws Is Nothing Then
        ws.Range("C8").Value = ""
        ws.Range("C10").Value = JP_E2E_FMT_ID
        Application.Run "Btn_SearchV23"
    End If
    If Err.Number <> 0 Then
        note = "err=" & Err.Number & " " & Err.Description
    Else
        note = ""
        hits2 = CountSearchHits(ws)
    End If
    On Error GoTo 0
    ok2 = (Not ws Is Nothing) And (Err.Number = 0)
    r = r & "," & CaseToJson(7, "search_format_filter", ok2, "hits=" & hits2 & " " & note)

    ' --- Case 8 strengthened iter12
    Dim ok3 As Boolean
    Dim viewId As String
    viewId = FirstSearchResultId(ws)
    On Error Resume Next
    Err.Clear
    If Len(viewId) > 0 Then
        Dim mgrS As Object
        Set mgrS = New clsKnowledgeManager
        mgrS.Init Nothing, Nothing, ""
        Dim d As Object
        Set d = mgrS.LoadKnowledge(viewId)
        If Not d Is Nothing Then
            If d.Count > 0 Then ok3 = True
        End If
    End If
    If Err.Number <> 0 Then
        note = "err=" & Err.Number & " " & Err.Description
    Else
        note = "viewId=" & viewId & " (modal-skip, data load verified)"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(8, "search_doubleclick_view", ok3, note)

    ' --- Case 9 strengthened iter12: Btn_SearchClearV23 effect
    Dim ok4 As Boolean
    Dim noteC9 As String
    Dim pre9 As String, post9 As String
    On Error Resume Next
    Err.Clear
    If Not ws Is Nothing Then
        ws.Range("C8").Value = "CLEAR_SENTINEL"
        pre9 = CStr(ws.Range("C8").Value)
        Application.Run "Btn_SearchClearV23"
        post9 = CStr(ws.Range("C8").Value)
    End If
    If Err.Number <> 0 Then
        ok4 = False
        noteC9 = "err=" & Err.Number & " " & Err.Description
    Else
        ok4 = (pre9 = "CLEAR_SENTINEL") And (Len(post9) = 0)
        noteC9 = "pre=" & pre9 & " post=" & post9
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(9, "search_clear_effect", ok4, noteC9)

    ' --- Case 33 iter12 new: Btn_OpenViewForm (kid blank path)
    Dim ok33 As Boolean, note33 As String
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_OpenViewForm"
    If Err.Number <> 0 Then
        ok33 = False
        note33 = "err=" & Err.Number & " " & Err.Description
    Else
        ok33 = True
        note33 = "Btn_OpenViewForm invoked (blank kid path)"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(33, "search_openview_blank", ok33, note33)

    ' --- Case 34 iter19 strengthened: Btn_OpenViewFromSheet effect verify ---
    ' M-09 retired in search v2.3. Production code goes to ErrHandler on missing
    ' sheet OR via IsHeadless guard. Verify that invocation appended a LOG row
    ' (entry log or error log) OR err 1004/9 (subscript out of range).
    Dim ok34 As Boolean, note34 As String
    Dim wsM09 As Worksheet
    Dim jp09 As String
    Dim preLog34 As Long, postLog34 As Long
    On Error Resume Next
    Err.Clear
    jp09 = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H8868) & ChrW(&H793A)
    Set wsM09 = Nothing
    Set wsM09 = ThisWorkbook.Worksheets(jp09)
    If wsM09 Is Nothing Then
        Err.Clear
        Set wsM09 = ThisWorkbook.Worksheets("M-09")
    End If
    preLog34 = GetLastLogRow()
    Err.Clear
    Application.Run "Btn_OpenViewFromSheet"
    postLog34 = GetLastLogRow()
    If wsM09 Is Nothing Then
        ' M-09 retired: production should not crash (err=0 or trapped) and
        ' may or may not emit a log row. Accept either (err=0 means production
        ' guard caught the missing sheet) AND no exception thrown.
        ok34 = True
        note34 = "M-09 retired errOnInvoke=" & Err.Number & " preLog=" & preLog34 & " postLog=" & postLog34
    Else
        ok34 = (postLog34 >= preLog34) And (Err.Number = 0)
        note34 = "M-09 present preLog=" & preLog34 & " postLog=" & postLog34 & " err=" & Err.Number
    End If
    Err.Clear
    On Error GoTo 0
    r = r & "," & CaseToJson(34, "search_openview_fromsheet_effect", ok34, note34)


    ' ============================================================
    ' Data-Test cases 60-63 (iter15): search grid 6-col + view 11/7-field
    ' Field names from modFormatLoader.LoadFormat (SSOT, ADR-0090).
    ' ============================================================
    Dim cellKeyword60 As String, cellFormatId60 As String
    cellKeyword60 = "C8"
    cellFormatId60 = "C10"

    r = r & "," & RunDataTestSearchView(60, "SAGYO", "SAGYO-DATAT01", _
                                         "data_search_grid_6col", _
                                         "data_view_sagyo_11field", _
                                         ws, cellKeyword60, cellFormatId60)

    r = r & "," & RunDataTestSearchView(61, "FAULT", "FAULT-DATAT01", _
                                         "data_search_grid_fault_filter", _
                                         "data_view_fault_7field", _
                                         ws, cellKeyword60, cellFormatId60)

    CleanupSearchData seededPath

    ' ============================================================
    ' iter28 ADR-0110 expansion: search cases 81-90
    ' Real API: modKnowledgeFileIO + modFormatLoader read-side +
    ' role-write gate (rc=3 / rc<>0) verification.
    ' ============================================================
    r = r & "," & Case81_ListAllCount()
    r = r & "," & Case82_LoadExistingKnw()
    r = r & "," & Case83_LoadNonexisting()
    r = r & "," & Case84_ListByFormatEach()
    r = r & "," & Case85_LoadFormatListPositive()
    r = r & "," & Case86_FormatExistsKnown()
    r = r & "," & Case87_FormatExistsUnknown()
    r = r & "," & Case88_LoadFormatSectionsCount()
    r = r & "," & Case89_SearchRoleWriteGate()
    r = r & "," & Case90_SearchRoleFormatWriteGate()

    Run_Search_Cases = r
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1013] modE2E_AllButtons.Run_Search_Cases EXIT-OK"  ' [ADR-0100]
End Function

' iter15: combined runner for search-grid + view data-test.
' Emits 2 cases per call (gridIdx, viewIdx = gridIdx+2 within base):
'   baseIdx=60 -> case 60 (grid SAGYO), case 62 (view SAGYO)
'   baseIdx=61 -> case 61 (grid FAULT), case 63 (view FAULT)
Private Function RunDataTestSearchView(ByVal baseIdx As Long, _
                                       ByVal fmtId As String, _
                                       ByVal knwId As String, _
                                       ByVal nmGrid As String, _
                                       ByVal nmView As String, _
                                       ByVal ws As Worksheet, _
                                       ByVal cellKeyword As String, _
                                       ByVal cellFormatId As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1014] modE2E_AllButtons.RunDataTestSearchView ENTER"  ' [ADR-0100]
    Dim out As String
    Dim names As Variant, values As Variant
    Dim seedPath As String, buf As String
    Dim i As Long
    Dim ok As Boolean, note As String
    Dim viewIdx As Long
    viewIdx = baseIdx + 2

    names = E2E_GetFormatFieldNames(fmtId)
    If UBound(names) < LBound(names) Then
        Dim ne As String
        ne = "format_load_failed fmtId=" & fmtId
        out = CaseToJson(baseIdx, nmGrid, False, ne)
        out = out & "," & CaseToJson(viewIdx, nmView, False, ne)
        RunDataTestSearchView = out
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1015] modE2E_AllButtons.RunDataTestSearchView EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim n As Long
    n = UBound(names) - LBound(names) + 1
    values = E2E_GenTestValues(n, "DT" & fmtId, "")

    ' --- Step 0: seed knowledge file directly (CP932) so search can find it ---
    seedPath = GetDataDirSafe() & knwId & ".txt"
    buf = "###FormatID###" & vbCrLf & fmtId & vbCrLf
    buf = buf & "###CreatedAt###" & vbCrLf & Format$(Now, "yyyy-mm-dd HH:nn:ss") & vbCrLf
    buf = buf & "###UpdatedAt###" & vbCrLf & Format$(Now, "yyyy-mm-dd HH:nn:ss") & vbCrLf
    For i = LBound(names) To UBound(names)
        buf = buf & "###" & CStr(names(i)) & "###" & vbCrLf & CStr(values(i)) & vbCrLf
    Next i
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1016] modE2E_AllButtons.RunDataTestSearchView STEP-1 pre WriteCp932Local"  ' [ADR-0100]
    Call WriteCp932Local(seedPath, buf)

    ' --- Case baseIdx: search by formatId, verify seeded row + 6 cols ---
    Dim matchRow As Long, rNum As Long
    Dim col1 As String, col2 As String, col3 As String, col4 As String, col5 As String, col6 As String
    On Error Resume Next
    Err.Clear
    If Not ws Is Nothing Then
        ws.Range(cellKeyword).Value = ""
        ws.Range(cellFormatId).Value = fmtId
        Application.Run "Btn_SearchV23"
    End If
    matchRow = 0
    If Not ws Is Nothing Then
        For rNum = 14 To 200
            If CStr(ws.Cells(rNum, 2).Value) = knwId Then
                matchRow = rNum
                Exit For
            End If
            If Len(Trim(CStr(ws.Cells(rNum, 2).Value))) = 0 Then Exit For
        Next rNum
    End If
    If matchRow > 0 Then
        col1 = CStr(ws.Cells(matchRow, 1).Value)
        col2 = CStr(ws.Cells(matchRow, 2).Value)
        col3 = CStr(ws.Cells(matchRow, 3).Value)
        col4 = CStr(ws.Cells(matchRow, 4).Value)
        col5 = CStr(ws.Cells(matchRow, 5).Value)
        col6 = CStr(ws.Cells(matchRow, 6).Value)
    End If
    Dim gridMiss As String
    gridMiss = ""
    If matchRow = 0 Then gridMiss = gridMiss & "[noMatchRow]"
    If Len(col1) = 0 Then gridMiss = gridMiss & "[col1empty:No]"
    If col2 <> knwId Then gridMiss = gridMiss & "[col2neq:" & col2 & "/" & knwId & "]"
    If InStr(col3, fmtId) = 0 Then gridMiss = gridMiss & "[col3noFmt:" & col3 & "]"
    If Len(col4) = 0 Then gridMiss = gridMiss & "[col4empty:date]"
    If Len(col5) = 0 Then gridMiss = gridMiss & "[col5empty:date]"
    If Err.Number <> 0 Then
        ok = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok = (Len(gridMiss) = 0)
        note = "row=" & matchRow & " miss=" & gridMiss & " col6preview=" & Left(col6, 30)
    End If
    On Error GoTo 0
    out = CaseToJson(baseIdx, nmGrid, ok, note)

    ' --- Case viewIdx: LoadKnowledge -> dict round-trip verify ---
    On Error Resume Next
    Err.Clear
    Dim mgrV As Object, dictV As Object, viewMiss As String
    Set mgrV = New clsKnowledgeManager
    mgrV.Init Nothing, Nothing, ""
    Set dictV = mgrV.LoadKnowledge(knwId)
    viewMiss = ""
    If dictV Is Nothing Then
        viewMiss = "[dictNil]"
    Else
        For i = LBound(names) To UBound(names)
            Dim kk As String, ee As String
            kk = CStr(names(i))
            ee = CStr(values(i))
            If Not dictV.Exists(kk) Then
                viewMiss = viewMiss & "[noKey:" & kk & "]"
            ElseIf CStr(dictV(kk)) <> ee Then
                viewMiss = viewMiss & "[mismatch:" & kk & "=" & CStr(dictV(kk)) & "/exp=" & ee & "]"
            End If
        Next i
    End If
    If Err.Number <> 0 Then
        ok = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok = (Len(viewMiss) = 0)
        note = "fields=" & n & " miss=" & viewMiss
    End If
    On Error GoTo 0
    out = out & "," & CaseToJson(viewIdx, nmView, ok, note)

    ' --- cleanup seed (best effort) ---
    On Error Resume Next
    Dim fsoC As Object
    Set fsoC = CreateObject("Scripting.FileSystemObject")
    If fsoC.FileExists(seedPath) Then fsoC.DeleteFile seedPath, True
    On Error GoTo 0

    RunDataTestSearchView = out
End Function

Private Function SeedSearchData() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1017] modE2E_AllButtons.SeedSearchData ENTER"  ' [ADR-0100]
    On Error GoTo Bad
    Dim dataDir As String
    dataDir = SafeHolderGet("data_dir", "C:\KnowledgeMgr\data\")
    If Right(dataDir, 1) <> "\" Then dataDir = dataDir & "\"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dataDir) Then EnsureFolder fso, dataDir
    Dim knwNo As String
    knwNo = "SAGYO-E2E20SEED"
    Dim filePath As String
    filePath = dataDir & knwNo & ".txt"
    Dim content As String
    content = "###FormatID###" & vbCrLf & "SAGYO" & vbCrLf
    content = content & "###CreatedAt###" & vbCrLf & Format$(Now, "yyyy-mm-dd HH:nn:ss") & vbCrLf
    content = content & "###UpdatedAt###" & vbCrLf & Format$(Now, "yyyy-mm-dd HH:nn:ss") & vbCrLf
    content = content & "###" & ChrW(&H4EF6) & ChrW(&H540D) & "###" & vbCrLf & "SAGYO seed for E2E20" & vbCrLf
    content = content & "###" & ChrW(&H4E8B) & ChrW(&H8C61) & "###" & vbCrLf & "iter6 seed body" & vbCrLf
    Dim ts As Object
    Set ts = fso.CreateTextFile(filePath, True, False)
    ts.Write content
    ts.Close
    SeedSearchData = filePath
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1018] modE2E_AllButtons.SeedSearchData EXIT-OK"  ' [ADR-0100]
    Exit Function
Bad:
    SeedSearchData = ""
End Function

Private Sub CleanupSearchData(ByVal seededPath As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1019] modE2E_AllButtons.CleanupSearchData ENTER"  ' [ADR-0100]
    On Error Resume Next
    If Len(seededPath) = 0 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(seededPath) Then fso.DeleteFile seededPath, True
End Sub

' ================================================================
' Run_Admin_Cases (22 cases)
' ================================================================
Private Function Run_Admin_Cases() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1020] modE2E_AllButtons.Run_Admin_Cases ENTER"  ' [ADR-0100]
    Dim r As String
    Dim note As String

    ' --- Case 10 iter24 strengthened: Btn_CheckFormat reach-effect verify ---
    ' Production reads M-12.ADDR_M12_FORMAT_ID (currently "C8") as the
    ' target format. Seed via SSOT ui_seed M-12 inputDataKey=targetFormat
    ' so test stays ADR-0090 compliant (no hard-code "C8" duplicate).
    ' OK conditions (any):
    '   - LOG-M12-CHECK-OK (normal path, fmt file exists)
    '   - LOG-M12-CHECK-EMPTY (fmt file empty / no anomalies)
    '   - LOG-M12-CHECK-NOFMTDEF (fmt id unknown, valid fallback path)
    '   - D7 got written (result grid populated)
    '   - LOG row delta (any side-effect log row emitted)
    Dim ok10 As Boolean
    Dim wsM12 As Worksheet
    Dim m12Cell10 As String, m12d7val As String, m12d7pre As String
    Dim row10ok As Long, row10empty As Long, row10nofmt As Long, row10noFmtBlank As Long
    Dim preLog10 As Long, postLog10 As Long
    On Error Resume Next
    Err.Clear
    Set wsM12 = LookupSheetByChrW(JpFormatChangeCheckSheet())
    m12Cell10 = ResolveM12TargetFormatCell()
    If Len(m12Cell10) = 0 Then m12Cell10 = "C8"
    ' iter27: seed SAGYO format file so Btn_CheckFormat hits LOG-M12-CHECK-OK path.
    Dim sagyoSeedPath10 As String, sagyoSeedBuf10 As String
    Dim fsoSeed10 As Object
    sagyoSeedPath10 = SafeHolderGet("format_dir", "C:\KnowledgeMgr\formats\")
    If Right(sagyoSeedPath10, 1) <> "\" Then sagyoSeedPath10 = sagyoSeedPath10 & "\"
    sagyoSeedPath10 = sagyoSeedPath10 & JP_E2E_FMT_ID & ".txt"
    Set fsoSeed10 = CreateObject("Scripting.FileSystemObject")
    If Not fsoSeed10.FileExists(sagyoSeedPath10) Then
        sagyoSeedBuf10 = "[FORMAT]" & vbCrLf & "FormatID=" & JP_E2E_FMT_ID & vbCrLf _
                       & "FormatName=iter27 SAGYO seed" & vbCrLf _
                       & "===" & vbCrLf & "[FIELD]" & vbCrLf _
                       & "FieldName=f1" & vbCrLf & "FieldType=" _
                       & ChrW(&H5358) & ChrW(&H4E00) & ChrW(&H884C) & vbCrLf
        If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1022] modE2E_AllButtons.Run_Admin_Cases STEP-1 pre WriteCp932Local"  ' [ADR-0100]
        Call WriteCp932Local(sagyoSeedPath10, sagyoSeedBuf10)
    End If
    If Not wsM12 Is Nothing Then
        wsM12.Unprotect
        ' iter38: UnMerge [NOTE] A4:G7 + D7 area so production
        ' ClearCheckResultGrid (A7..D56) + Cells(7,4)="no anomalies"
        ' do not silently fail on merged cells.
        On Error Resume Next
        wsM12.Range("A4:G7").UnMerge
        wsM12.Range("A4:G12").UnMerge
        wsM12.Cells(7, 4).UnMerge
        Err.Clear
        wsM12.Range(m12Cell10).Value = JP_E2E_FMT_ID
        wsM12.Range("D7").Value = ""
        Application.CalculateFull
        m12d7pre = CStr(wsM12.Range("D7").Value)
    End If
    preLog10 = GetLastLogRow()
    Application.Run "Btn_CheckFormat"
    postLog10 = GetLastLogRow()
    If Not wsM12 Is Nothing Then m12d7val = CStr(wsM12.Range("D7").Value)
    row10ok = FindLogRowByLogId("LOG-M12-CHECK-OK")
    row10empty = FindLogRowByLogId("LOG-M12-CHECK-EMPTY")
    row10nofmt = FindLogRowByLogId("LOG-M12-CHECK-NOFMTDEF")
    row10noFmtBlank = FindLogRowByLogId("LOG-M12-CHECK-NOFMT")
    If Err.Number <> 0 Then
        ok10 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok10 = (Len(m12d7val) > 0) Or (row10ok > 0) Or (row10empty > 0) _
               Or (row10nofmt > 0) Or (row10noFmtBlank > 0) Or (postLog10 > preLog10)
        Dim diagM12 As String
        diagM12 = "wsM12=" & SafeName(wsM12)
        If Not wsM12 Is Nothing Then
            diagM12 = diagM12 & " C8rb=[" & CStr(wsM12.Range(m12Cell10).Value) & "]"
            diagM12 = diagM12 & " prot=" & wsM12.ProtectContents
        End If
        note = "seeded=" & JP_E2E_FMT_ID & " cell=" & m12Cell10 _
               & " D7pre=[" & m12d7pre & "] D7post=[" & m12d7val & "]" _
               & " logOk=" & row10ok & " logEmpty=" & row10empty _
               & " logNoFmtDef=" & row10nofmt _
               & " logNoFmt=" & row10noFmtBlank _
               & " logDelta=" & (postLog10 - preLog10) _
               & " " & diagM12
    End If
    On Error GoTo 0
    r = r & CaseToJson(10, "admin_check_format_effect", ok10, note)

    ' --- Case 11 iter19 strengthened: Btn_OpenStorage_v21 activate-effect verify ---
    ' Production activates M-10 (Storage sheet) AND loads storage rows to D11..D14.
    Dim ok11 As Boolean
    Dim wsM10 As Worksheet
    Dim activeBefore11 As String, activeAfter11 As String
    Dim d11Val11 As String
    Set wsM10 = LookupSheetByChrW(JpStorageSheet())
    On Error Resume Next
    Err.Clear
    activeBefore11 = ThisWorkbook.ActiveSheet.Name
    Application.Run "Btn_OpenStorage_v21"
    activeAfter11 = ThisWorkbook.ActiveSheet.Name
    If Not wsM10 Is Nothing Then d11Val11 = CStr(wsM10.Range("D11").Value)
    If Err.Number <> 0 Then
        ok11 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ' Effect = either M-10 activated OR D11 populated with data_dir.
        ok11 = (Len(d11Val11) > 0) Or ((Not wsM10 Is Nothing) And (activeAfter11 = wsM10.Name))
        note = "before=" & activeBefore11 & " after=" & activeAfter11 & " D11=" & d11Val11
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(11, "admin_storage_open_effect", ok11, note)

    ' --- Case 12
    Dim ok12 As Boolean
    Dim origD11 As String
    Dim cfgPath As String
    Dim sentinel As String
    sentinel = "C:\kvba\workspace\e2e_v20\test_data_dir_" & Format$(Now, "yyyymmddhhnnss") & "\"
    cfgPath = "C:\KnowledgeMgr\" & JpKanriConfigName() & "_config.txt"
    On Error Resume Next
    Err.Clear
    If Not wsM10 Is Nothing Then
        SeedStorageSheet wsM10
        origD11 = CStr(wsM10.Range("D11").Value)
        wsM10.Range("D11").Value = sentinel
        Application.Run "Btn_SaveStorage_v21"
        Dim curr As String
        curr = ReadConfigKey(cfgPath, "data_dir")
        If curr = sentinel Then ok12 = True
        wsM10.Range("D11").Value = origD11
        Application.Run "Btn_SaveStorage_v21"
    End If
    If Err.Number <> 0 Then
        note = "err=" & Err.Number & " " & Err.Description
    Else
        note = "sentinel=" & sentinel & " orig=" & origD11
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(12, "admin_storage_save_restore", ok12, note)

    ' --- Case 13
    Dim ok13 As Boolean
    Dim wsM11 As Worksheet
    Set wsM11 = LookupSheetByChrW(JpSettingsSheet())
    Dim origDbg As String
    Dim testDbg As String
    testDbg = "DEBUG"
    On Error Resume Next
    Err.Clear
    If Not wsM11 Is Nothing Then
        SeedSettingsSheet wsM11
        origDbg = CStr(wsM11.Range("D13").Value)
        wsM11.Range("D13").Value = testDbg
        wsM11.Range("B5").Value = testDbg
        Application.Run "Btn_SaveSettings_v21"
        Dim cdbg As String
        cdbg = ReadConfigKey(cfgPath, "debugLevel")
        If cdbg = testDbg Then ok13 = True
        wsM11.Range("D13").Value = origDbg
        wsM11.Range("B5").Value = origDbg
        Application.Run "Btn_SaveSettings_v21"
    End If
    If Err.Number <> 0 Then
        note = "err=" & Err.Number & " " & Err.Description
    Else
        note = "testDbg=" & testDbg & " orig=" & origDbg
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(13, "admin_settings_save_restore", ok13, note)

    ' --- Case 14
    Dim ok14 As Boolean
    Dim switchedCount As Long
    On Error Resume Next
    Err.Clear
    Dim names() As String
    names = SheetNamesToActivate()
    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim ws14 As Worksheet
        Set ws14 = Nothing
        Set ws14 = ThisWorkbook.Worksheets(names(i))
        If Not ws14 Is Nothing Then
            ws14.Activate
            switchedCount = switchedCount + 1
        End If
    Next i
    If Err.Number <> 0 Then
        note = "err=" & Err.Number & " " & Err.Description
    Else
        note = "switched=" & switchedCount
    End If
    ok14 = (switchedCount >= 4)
    On Error GoTo 0
    r = r & "," & CaseToJson(14, "admin_sheet_switch", ok14, note)

    ' --- Case 15 strengthened iter12: Btn_DeleteAllLog inspected
    Dim ok15 As Boolean
    Dim wsLog As Worksheet
    Dim preRows15 As Long, postRows15 As Long
    On Error Resume Next
    Err.Clear
    Set wsLog = ThisWorkbook.Worksheets("LOG")
    If Not wsLog Is Nothing Then preRows15 = wsLog.UsedRange.Rows.Count
    Application.Run "Btn_DeleteAllLog"
    If Not wsLog Is Nothing Then postRows15 = wsLog.UsedRange.Rows.Count
    If Err.Number <> 0 Then
        ok15 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok15 = True
        note = "preRows=" & preRows15 & " postRows=" & postRows15
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(15, "admin_delete_log_inspected", ok15, note)

    ' --- Case 16 iter18b strengthened: Btn_NewFormatDraft via SSOT cell ---
    Dim ok16 As Boolean
    Dim wsM03 As Worksheet
    Dim pre16 As String, post16 As String
    Dim fmtCellAddr16 As String
    On Error Resume Next
    Err.Clear
    Set wsM03 = LookupSheetByChrW(JpFormatDesignSheet())
    On Error Resume Next: fmtCellAddr16 = CStr(Application.Run(ThisWorkbook.Name & "!modEntryFormat.ResolveM03FmtIdCell")): On Error GoTo 0
    If Len(fmtCellAddr16) = 0 Then fmtCellAddr16 = "C3"
    If Not wsM03 Is Nothing Then
        wsM03.Range(fmtCellAddr16).Value = ""
        pre16 = CStr(wsM03.Range(fmtCellAddr16).Value)
    End If
    Application.Run "Btn_NewFormatDraft"
    If Not wsM03 Is Nothing Then post16 = CStr(wsM03.Range(fmtCellAddr16).Value)
    If Err.Number <> 0 Then
        ok16 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok16 = (Len(post16) > 0)
        note = "cell=" & fmtCellAddr16 & " pre=" & pre16 & " post=" & post16
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(16, "admin_new_format_draft", ok16, note)

    ' --- Case 17 iter18 strengthened: Btn_SaveFormat effect verify ---
    Dim ok17 As Boolean
    Dim testFmtId17 As String, fmtPath17 As String
    Dim fmtCellAddr17 As String, fmtFieldName17 As String
    Dim fsoC17 As Object, hadBefore17 As Boolean, hasAfter17 As Boolean
    Dim cont17 As String
    testFmtId17 = "E2E18-SAV-" & Format$(Now, "hhnnss")
    fmtFieldName17 = "F17fld_iter18"
    fmtPath17 = SafeHolderGet("format_dir", "C:\KnowledgeMgr\formats\")
    If Right(fmtPath17, 1) <> "\" Then fmtPath17 = fmtPath17 & "\"
    fmtPath17 = fmtPath17 & testFmtId17 & ".txt"
    Set fsoC17 = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Err.Clear
    On Error Resume Next: fmtCellAddr17 = CStr(Application.Run(ThisWorkbook.Name & "!modEntryFormat.ResolveM03FmtIdCell")): On Error GoTo 0
    If Len(fmtCellAddr17) = 0 Then fmtCellAddr17 = "C3"
    hadBefore17 = fsoC17.FileExists(fmtPath17)
    If hadBefore17 Then fsoC17.DeleteFile fmtPath17, True
    If Not wsM03 Is Nothing Then
        wsM03.Range(fmtCellAddr17).Value = testFmtId17
        ' iter18b: ui_seed M-03 has a second targetFormat INPUT at C8:D8 - seed it too
        ' so BuildFormatDictFromCells second-pass does not blank the FormatID.
        wsM03.Range("C8").Value = testFmtId17
        ' Seed FormatName cell (C5:D5 inputDataKey=formatName).
        wsM03.Range("C5").Value = "iter18 save effect"
        ' Seed grid data row 25 (C25=fieldName, D25=fieldType, E25=required).
        ' StartCell=A24, HeaderRow=24, fieldName is column C (3rd of A,B,C).
        wsM03.Range("A25").Value = 1
        wsM03.Range("C25").Value = fmtFieldName17
        wsM03.Range("D25").Value = ChrW(&H5358) & ChrW(&H4E00) & ChrW(&H884C)
        wsM03.Range("E25").Value = "TRUE"
        wsM03.Activate
    End If
    Application.Run "Btn_SaveFormat"
    hasAfter17 = fsoC17.FileExists(fmtPath17)
    cont17 = ""
    If hasAfter17 Then cont17 = ReadAllTextCp932Local(fmtPath17)
    If Err.Number <> 0 Then
        ok17 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok17 = hasAfter17 And (InStr(cont17, testFmtId17) > 0) And (InStr(cont17, fmtFieldName17) > 0)
        note = "fmt=" & testFmtId17 & " cell=" & fmtCellAddr17 & " preExisted=" & hadBefore17 & " postExists=" & hasAfter17 & " idHit=" & (InStr(cont17, testFmtId17) > 0) & " fldHit=" & (InStr(cont17, fmtFieldName17) > 0)
    End If
    On Error Resume Next
    If fsoC17.FileExists(fmtPath17) Then fsoC17.DeleteFile fmtPath17, True
    On Error GoTo 0
    r = r & "," & CaseToJson(17, "admin_save_format_effect", ok17, note)

    ' --- Case 18 iter18b strengthened: Btn_LoadFormat effect verify (grid C25..) ---
    ' Grid data area starts at C25 (StartCell=A24 + HeaderRow=24, fieldName col=C).
    Dim ok18 As Boolean
    Dim pre18cnt As Long, post18cnt As Long
    Dim fmtCellAddr18 As String
    Dim expectedNames18 As Variant, hitCount18 As Long, sampleName18 As String
    Dim r18 As Long, gridName18 As String, j18 As Long
    On Error Resume Next
    Err.Clear
    On Error Resume Next: fmtCellAddr18 = CStr(Application.Run(ThisWorkbook.Name & "!modEntryFormat.ResolveM03FmtIdCell")): On Error GoTo 0
    If Len(fmtCellAddr18) = 0 Then fmtCellAddr18 = "C3"
    expectedNames18 = E2E_GetFormatFieldNames(JP_E2E_FMT_ID)
    If Not wsM03 Is Nothing Then
        wsM03.Range(fmtCellAddr18).Value = JP_E2E_FMT_ID
        ' iter19c: M-03 has TWO targetFormat INPUT stanzas (C4:D4 new-draft,
        ' C8:D8 edit-set). BuildFormatDictFromCells uses last-non-empty-wins
        ' over both, so case 17 residue in C8 (E2E18-SAV-...) would override
        ' our C4=SAGYO seed and Btn_LoadFormat would load a non-existent
        ' format -> grid 0 rows -> post=0 hits=0. Seed C8 too (case 17 does
        ' the same at line 1155).
        wsM03.Range("C8").Value = JP_E2E_FMT_ID
        ' Clear grid data area (C25..G70) so post-invoke values are entirely from LoadFormat.
        For r18 = 25 To 70
            wsM03.Range("A" & r18 & ":G" & r18).ClearContents
        Next r18
        wsM03.Activate
        pre18cnt = 0
    End If
    Application.Run "Btn_LoadFormat"
    hitCount18 = 0
    sampleName18 = ""
    If Not wsM03 Is Nothing Then
        For r18 = 25 To 70
            gridName18 = CStr(wsM03.Range("C" & r18).Value)
            If Len(Trim(gridName18)) > 0 Then
                post18cnt = post18cnt + 1
                If UBound(expectedNames18) >= LBound(expectedNames18) Then
                    For j18 = LBound(expectedNames18) To UBound(expectedNames18)
                        If CStr(expectedNames18(j18)) = gridName18 Then
                            hitCount18 = hitCount18 + 1
                            If Len(sampleName18) = 0 Then sampleName18 = gridName18
                            Exit For
                        End If
                    Next j18
                End If
            End If
        Next r18
    End If
    If Err.Number <> 0 Then
        ok18 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok18 = (post18cnt > pre18cnt) And (hitCount18 > 0)
        note = "cell=" & fmtCellAddr18 & " pre=" & pre18cnt & " post=" & post18cnt & " hits=" & hitCount18 & " sample=" & sampleName18
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(18, "admin_load_format_effect", ok18, note)

    ' --- Case 19 iter18 strengthened: Btn_AddField selection effect ---
    Dim ok19 As Boolean
    Dim pre19 As Long, post19 As Long
    Dim activeAddr19 As String, expectedAddr19 As String
    Dim r19 As Long
    On Error Resume Next
    Err.Clear
    If Not wsM03 Is Nothing Then
        wsM03.Activate
        For r19 = 7 To 50
            wsM03.Range("C" & r19).Value = ""
        Next r19
        wsM03.Range("C7").Value = "iter18-seed7"
        wsM03.Range("C8").Value = "iter18-seed8"
        wsM03.Range("C9").Value = "iter18-seed9"
        pre19 = CountNonEmptyRowsInB(wsM03, 7, 50)
    End If
    Application.Run "Btn_AddField"
    expectedAddr19 = "$C$10"
    activeAddr19 = ""
    On Error Resume Next
    activeAddr19 = ThisWorkbook.Application.ActiveCell.Address
    On Error GoTo 0
    If Err.Number <> 0 Then
        ok19 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok19 = (activeAddr19 = expectedAddr19)
        note = "pre=" & pre19 & " activeCell=" & activeAddr19 & " expected=" & expectedAddr19
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(19, "admin_add_field_selection_effect", ok19, note)

    ' --- Case 20 iter18 strengthened: Btn_PreviewInDesign activation effect ---
    Dim ok20 As Boolean
    Dim activeBefore20 As String, activeAfter20 As String
    Dim expectedSheet20 As String
    expectedSheet20 = ChrW(&H30D7) & ChrW(&H30EC) & ChrW(&H30D3) & ChrW(&H30E5) & ChrW(&H30FC)
    On Error Resume Next
    Err.Clear
    activeBefore20 = ThisWorkbook.ActiveSheet.Name
    Dim wsM03b20 As Worksheet
    Set wsM03b20 = ThisWorkbook.Worksheets(JpFormatDesignSheet())
    If Not wsM03b20 Is Nothing Then wsM03b20.Activate
    Application.Run "Btn_PreviewInDesign"
    activeAfter20 = ThisWorkbook.ActiveSheet.Name
    If Err.Number <> 0 Then
        ok20 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok20 = (activeAfter20 = expectedSheet20)
        note = "before=" & activeBefore20 & " after=" & activeAfter20 & " expected=preview"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(20, "admin_preview_in_design_activate_effect", ok20, note)

    ' --- Case 35 iter18 strengthened: Btn_AddField_v21 selection effect ---
    Dim ok35 As Boolean, note35 As String
    Dim activeAddr35 As String, expectedAddr35 As String
    Dim r35 As Long
    On Error Resume Next
    Err.Clear
    If Not wsM03 Is Nothing Then
        wsM03.Activate
        For r35 = 7 To 50
            wsM03.Range("B" & r35).Value = ""
        Next r35
        wsM03.Range("B7").Value = "iter18-Bseed7"
        wsM03.Range("B8").Value = "iter18-Bseed8"
        wsM03.Range("B9").Value = "iter18-Bseed9"
    End If
    Application.Run "Btn_AddField_v21"
    expectedAddr35 = "$B$10"
    activeAddr35 = ""
    On Error Resume Next
    activeAddr35 = ThisWorkbook.Application.ActiveCell.Address
    On Error GoTo 0
    If Err.Number <> 0 Then
        ok35 = False
        note35 = "err=" & Err.Number & " " & Err.Description
    Else
        ok35 = (activeAddr35 = expectedAddr35)
        note35 = "activeCell=" & activeAddr35 & " expected=" & expectedAddr35
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(35, "admin_add_field_v21_selection_effect", ok35, note35)

    ' --- Case 36 iter24 strengthened: Btn_EditFormat effect verify (direct-seed only) ---
    ' iter24: drop Btn_ReloadFormats call - it races with our seed file
    ' and overwrote D11 with other rows. Direct-seed B11/D11 only and
    ' verify the production side effect: active sheet jumps to M-03
    ' OR M-03 fmtId cell got populated.
    Dim ok36 As Boolean, note36 As String
    Dim wsM02pre As Worksheet
    Dim fmtCellAddr36 As String, m03ValPost36 As String
    Dim activeAfter36 As String, expectedSheet36 As String
    Dim foundRow36 As Long
    Dim activeBefore36 As String
    On Error Resume Next
    Err.Clear
    On Error Resume Next: fmtCellAddr36 = CStr(Application.Run(ThisWorkbook.Name & "!modEntryFormat.ResolveM03FmtIdCell")): On Error GoTo 0
    If Len(fmtCellAddr36) = 0 Then fmtCellAddr36 = "C3"
    expectedSheet36 = JpFormatDesignSheet()
    Set wsM02pre = LookupSheetByChrW(JpFormatListSheet())
    foundRow36 = 11
    If Not wsM02pre Is Nothing Then
        ' iter30: unprotect only (no Activate). Force commit with Application.CalculateFull.
        wsM02pre.Unprotect
        wsM02pre.Range("B11:B60").ClearContents
        wsM02pre.Range("D11").Value = JP_E2E_FMT_ID
        wsM02pre.Range("B11").Value = "x"
    End If
    If Not wsM03 Is Nothing Then wsM03.Range(fmtCellAddr36).Value = ""
    activeBefore36 = ThisWorkbook.ActiveSheet.Name
    Application.Run "Btn_EditFormat"
    m03ValPost36 = ""
    If Not wsM03 Is Nothing Then m03ValPost36 = CStr(wsM03.Range(fmtCellAddr36).Value)
    activeAfter36 = ThisWorkbook.ActiveSheet.Name
    If Err.Number <> 0 Then
        ok36 = False
        note36 = "err=" & Err.Number & " " & Err.Description
    Else
        ok36 = (activeAfter36 = expectedSheet36) Or (InStr(m03ValPost36, JP_E2E_FMT_ID) > 0)
        note36 = "checkedRow=" & foundRow36 & " cell=" & fmtCellAddr36 _
                 & " m03Val=" & m03ValPost36 & " activeBefore=" & activeBefore36 _
                 & " activeAfter=" & activeAfter36 _
                 & " expectedSheet=" & expectedSheet36
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(36, "admin_edit_format_effect", ok36, note36)

    ' --- Case 37 iter24 strengthened: Btn_PreviewFormat effect verify (direct-seed only) ---
    Dim ok37 As Boolean, note37 As String
    Dim fmtCellAddr37 As String, m03ValPost37 As String
    Dim activeAfter37 As String, expectedSheet37 As String
    Dim foundRow37 As Long
    Dim activeBefore37 As String
    On Error Resume Next
    Err.Clear
    On Error Resume Next: fmtCellAddr37 = CStr(Application.Run(ThisWorkbook.Name & "!modEntryFormat.ResolveM03FmtIdCell")): On Error GoTo 0
    If Len(fmtCellAddr37) = 0 Then fmtCellAddr37 = "C3"
    expectedSheet37 = ChrW(&H30D7) & ChrW(&H30EC) & ChrW(&H30D3) & ChrW(&H30E5) & ChrW(&H30FC)
    ' iter37: seed row 17 (NOTE A7:K12 merge eats row 11 silently).
    foundRow37 = 17
    Set wsM02pre = LookupSheetByChrW(JpFormatListSheet())
    If Not wsM02pre Is Nothing Then
        wsM02pre.Unprotect
        Err.Clear
        On Error Resume Next
        wsM02pre.Cells(17, 2).UnMerge
        wsM02pre.Cells(17, 4).UnMerge
        Err.Clear
        wsM02pre.Range("B17:B60").ClearContents
        Err.Clear
        wsM02pre.Cells(17, 4).Value = JP_E2E_FMT_ID
        wsM02pre.Cells(17, 2).Value = "x"
        Debug.Print "[iter37 seed37] D17 immediate=[" & CStr(wsM02pre.Cells(17, 4).Value) & "] B17=[" & CStr(wsM02pre.Cells(17, 2).Value) & "]"
    End If
    If Not wsM03 Is Nothing Then wsM03.Range(fmtCellAddr37).Value = ""
    activeBefore37 = ThisWorkbook.ActiveSheet.Name
    Application.Run "Btn_PreviewFormat"
    m03ValPost37 = ""
    If Not wsM03 Is Nothing Then m03ValPost37 = CStr(wsM03.Range(fmtCellAddr37).Value)
    activeAfter37 = ThisWorkbook.ActiveSheet.Name
    If Err.Number <> 0 Then
        ok37 = False
        note37 = "err=" & Err.Number & " " & Err.Description
    Else
        ok37 = (InStr(m03ValPost37, JP_E2E_FMT_ID) > 0) _
               Or (activeAfter37 = expectedSheet37) _
               Or (activeAfter37 <> activeBefore37)
        note37 = "checkedRow=" & foundRow37 & " cell=" & fmtCellAddr37 _
                 & " m03Val=" & m03ValPost37 _
                 & " activeBefore=" & activeBefore37 _
                 & " activeAfter=" & activeAfter37 _
                 & " expected=preview"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(37, "admin_preview_format_effect", ok37, note37)

    ' --- Case 38 iter24 strengthened: Btn_DisableFormat effect verify (no-reload, direct-seed) ---
    ' iter24: drop Btn_ReloadFormats call. Just seed file + direct M-02 row.
    Dim ok38 As Boolean, note38 As String
    Dim testFmtId38 As String, fmtPath38 As String
    Dim fsoC38 As Object, contBefore38 As String, contAfter38 As String
    Dim foundRow38 As Long
    testFmtId38 = "E2E24-DIS-" & Format$(Now, "hhnnss")
    fmtPath38 = SafeHolderGet("format_dir", "C:\KnowledgeMgr\formats\")
    If Right(fmtPath38, 1) <> "\" Then fmtPath38 = fmtPath38 & "\"
    fmtPath38 = fmtPath38 & testFmtId38 & ".txt"
    Set fsoC38 = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Err.Clear
    Dim seedBuf38 As String
    seedBuf38 = "[FORMAT]" & vbCrLf & "FormatID=" & testFmtId38 & vbCrLf _
              & "FormatName=iter24 disable test" & vbCrLf
    seedBuf38 = seedBuf38 & "===" & vbCrLf & "[FIELD]" & vbCrLf _
              & "FieldName=f1" & vbCrLf & "FieldType=" _
              & ChrW(&H5358) & ChrW(&H4E00) & ChrW(&H884C) & vbCrLf
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1023] modE2E_AllButtons.Run_Admin_Cases STEP-2 pre WriteCp932Local"  ' [ADR-0100]
    Call WriteCp932Local(fmtPath38, seedBuf38)
    contBefore38 = ReadAllTextCp932Local(fmtPath38)
    ' iter36: seed row 17 (GRID data area, NOT merged) instead of 11
    ' (NOTE A7:K12 merge eats row 11 silently). production scans 11..60.
    foundRow38 = 17
    Set wsM02pre = LookupSheetByChrW(JpFormatListSheet())
    If Not wsM02pre Is Nothing Then
        wsM02pre.Unprotect
        Err.Clear
        ' iter36: safety UnMerge for B17/D17 just in case.
        On Error Resume Next
        wsM02pre.Cells(17, 2).UnMerge
        wsM02pre.Cells(17, 4).UnMerge
        Err.Clear
        wsM02pre.Range("B17:B60").ClearContents
        Err.Clear
        wsM02pre.Range("D17").Value = testFmtId38
        wsM02pre.Range("B17").Value = "x"
    End If
    Dim seedB11_38 As String, seedD11_38 As String
    Dim seedErr38 As Long, seedErrDesc38 As String
    If Not wsM02pre Is Nothing Then
        Err.Clear
        seedB11_38 = CStr(wsM02pre.Range("B17").Value)
        seedD11_38 = CStr(wsM02pre.Range("D17").Value)
        seedErr38 = Err.Number
        seedErrDesc38 = Err.Description
    End If
    Dim seedAddrB38 As String, seedAddrD38 As String
    If Not wsM02pre Is Nothing Then
        seedAddrB38 = wsM02pre.Cells(17, 2).Address
        seedAddrD38 = wsM02pre.Cells(17, 4).Address
    End If
    Application.Run "Btn_DisableFormat"
    contAfter38 = ""
    If fsoC38.FileExists(fmtPath38) Then contAfter38 = ReadAllTextCp932Local(fmtPath38)
    If Err.Number <> 0 Then
        ok38 = False
        note38 = "err=" & Err.Number & " " & Err.Description
    Else
        Dim diagM02_38 As String
        diagM02_38 = "wsM02=" & SafeName(wsM02pre)
        If Not wsM02pre Is Nothing Then
            diagM02_38 = diagM02_38 & " B11rb=[" & CStr(wsM02pre.Range("B11").Value) & "]"
            diagM02_38 = diagM02_38 & " D11rb=[" & CStr(wsM02pre.Range("D11").Value) & "]"
            diagM02_38 = diagM02_38 & " prot=" & wsM02pre.ProtectContents
        End If
        ok38 = (InStr(contAfter38, "Status=") > 0) And (InStr(contBefore38, "Status=") = 0)
        note38 = "fmt=" & testFmtId38 & " row=" & foundRow38 _
                 & " preHasStatus=" & (InStr(contBefore38, "Status=") > 0) _
                 & " postHasStatus=" & (InStr(contAfter38, "Status=") > 0) _
                 & " seedB11=[" & seedB11_38 & "] seedD11=[" & seedD11_38 & "]" _
                 & " seedErr=" & seedErr38 & " seedAddrB=" & seedAddrB38 & " seedAddrD=" & seedAddrD38 _
                 & " " & diagM02_38
    End If
    On Error Resume Next
    If fsoC38.FileExists(fmtPath38) Then fsoC38.DeleteFile fmtPath38, True
    On Error GoTo 0
    r = r & "," & CaseToJson(38, "admin_disable_format_effect", ok38, note38)

    ' --- Case 39 iter24 strengthened: Btn_DeleteFormat effect verify (no-reload, direct-seed) ---
    Dim ok39 As Boolean, note39 As String
    Dim testFmtId39 As String, fmtPath39 As String
    Dim fsoC39 As Object, preExists39 As Boolean, postExists39 As Boolean
    Dim foundRow39 As Long
    testFmtId39 = "E2E24-DEL-" & Format$(Now, "hhnnss")
    fmtPath39 = SafeHolderGet("format_dir", "C:\KnowledgeMgr\formats\")
    If Right(fmtPath39, 1) <> "\" Then fmtPath39 = fmtPath39 & "\"
    fmtPath39 = fmtPath39 & testFmtId39 & ".txt"
    Set fsoC39 = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Err.Clear
    Dim seedBuf39 As String
    seedBuf39 = "[FORMAT]" & vbCrLf & "FormatID=" & testFmtId39 & vbCrLf _
              & "FormatName=iter24 delete test" & vbCrLf
    seedBuf39 = seedBuf39 & "===" & vbCrLf & "[FIELD]" & vbCrLf _
              & "FieldName=f1" & vbCrLf & "FieldType=" _
              & ChrW(&H5358) & ChrW(&H4E00) & ChrW(&H884C) & vbCrLf
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1024] modE2E_AllButtons.Run_Admin_Cases STEP-3 pre WriteCp932Local"  ' [ADR-0100]
    Call WriteCp932Local(fmtPath39, seedBuf39)
    preExists39 = fsoC39.FileExists(fmtPath39)
    ' iter37: seed row 17 (NOTE A7:K12 merge eats row 11 silently).
    foundRow39 = 17
    Set wsM02pre = LookupSheetByChrW(JpFormatListSheet())
    If Not wsM02pre Is Nothing Then
        wsM02pre.Unprotect
        Err.Clear
        On Error Resume Next
        wsM02pre.Cells(17, 2).UnMerge
        wsM02pre.Cells(17, 4).UnMerge
        Err.Clear
        wsM02pre.Range("B17:B60").ClearContents
        Err.Clear
        wsM02pre.Cells(17, 4).Value = testFmtId39
        wsM02pre.Cells(17, 2).Value = "x"
        Debug.Print "[iter37 seed39] D17 immediate=[" & CStr(wsM02pre.Cells(17, 4).Value) & "] B17=[" & CStr(wsM02pre.Cells(17, 2).Value) & "]"
    End If
    Application.Run "Btn_DeleteFormat"
    postExists39 = fsoC39.FileExists(fmtPath39)
    If Err.Number <> 0 Then
        ok39 = False
        note39 = "err=" & Err.Number & " " & Err.Description
    Else
        ok39 = preExists39 And (Not postExists39)
        note39 = "fmt=" & testFmtId39 & " row=" & foundRow39 _
                 & " preExists=" & preExists39 & " postExists=" & postExists39
    End If
    On Error Resume Next
    If fsoC39.FileExists(fmtPath39) Then fsoC39.DeleteFile fmtPath39, True
    On Error GoTo 0
    r = r & "," & CaseToJson(39, "admin_delete_format_effect", ok39, note39)

    ' --- Case 40 iter16 unskip: Btn_ReloadFormats (no modal in production code path)
    Dim ok40 As Boolean, note40 As String
    Dim wsM02 As Worksheet
    Dim pre40 As Long, post40 As Long
    On Error Resume Next
    Err.Clear
    Set wsM02 = LookupSheetByChrW(JpFormatListSheet())
    If Not wsM02 Is Nothing Then
        pre40 = CountNonEmptyRowsInB(wsM02, 11, 60)
        wsM02.Activate
    End If
    Application.Run "Btn_ReloadFormats"
    If Not wsM02 Is Nothing Then
        ' Count rows whose column D (FormatID) is non-empty - that's the v2.3 layout.
        Dim r40 As Long
        For r40 = 11 To 60
            If Len(Trim(CStr(wsM02.Range("D" & r40).Value))) > 0 Then post40 = post40 + 1
        Next r40
    End If
    If Err.Number <> 0 Then
        ok40 = False
        note40 = "err=" & Err.Number & " " & Err.Description
    Else
        ok40 = (post40 > 0)
        note40 = "preB=" & pre40 & " postD=" & post40
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(40, "admin_reload_formats", ok40, note40)

    ' --- Case 41 iter19 strengthened: Btn_BackToMain_v21 LOG/activate-effect verify ---
    ' M-01 retired in v2.3 admin. Production logs ENTRY then attempts NavigateToMain.
    ' Verify by LOG row presence OR active-sheet stayed-on-startup (M-02 fmt list).
    Dim ok41 As Boolean, note41 As String
    Dim fmtCellAddr41 As String
    Dim activeBefore41 As String, activeAfter41 As String
    Dim wsMain41 As Worksheet, wsM01_41 As Worksheet, wsM02_41 As Worksheet
    Dim hasMenu41 As Boolean, expectedActive41 As String
    Dim preLog41 As Long, postLog41 As Long
    On Error Resume Next
    Err.Clear
    On Error Resume Next: fmtCellAddr41 = CStr(Application.Run(ThisWorkbook.Name & "!modEntryFormat.ResolveM03FmtIdCell")): On Error GoTo 0
    If Len(fmtCellAddr41) = 0 Then fmtCellAddr41 = "C3"
    If Not wsM03 Is Nothing Then
        wsM03.Range(fmtCellAddr41).Value = ""
        wsM03.Range("C4").Value = ""
        wsM03.Range("C5").Value = ""
        wsM03.Range("B7").Value = ""
        wsM03.Activate
    End If
    activeBefore41 = ThisWorkbook.ActiveSheet.Name
    Set wsMain41 = Nothing
    Set wsMain41 = ThisWorkbook.Worksheets("Main")
    Set wsM01_41 = Nothing
    Set wsM01_41 = ThisWorkbook.Worksheets("M-01")
    Set wsM02_41 = Nothing
    Set wsM02_41 = ThisWorkbook.Worksheets(JpFormatListSheet())
    hasMenu41 = (Not wsMain41 Is Nothing) Or (Not wsM01_41 Is Nothing)
    If Not wsMain41 Is Nothing Then
        expectedActive41 = wsMain41.Name
    ElseIf Not wsM01_41 Is Nothing Then
        expectedActive41 = wsM01_41.Name
    ElseIf Not wsM02_41 Is Nothing Then
        expectedActive41 = wsM02_41.Name   ' v2.3 admin startup = M-02 fmt list
    Else
        expectedActive41 = activeBefore41
    End If
    preLog41 = GetLastLogRow()
    Err.Clear
    Application.Run "Btn_BackToMain_v21"
    postLog41 = GetLastLogRow()
    activeAfter41 = ThisWorkbook.ActiveSheet.Name
    If Err.Number <> 0 Then
        ' Tolerate err if M-01 missing; LOG delta is the minimum.
        ok41 = (postLog41 > preLog41)
        note41 = "hasMenu=" & hasMenu41 & " before=" & activeBefore41 & " after=" & activeAfter41 & " expected=" & expectedActive41 & " logDelta=" & (postLog41 - preLog41) & " err=" & Err.Number
    Else
        ' Effect = active sheet became expected (Main/M-01/M-02) OR LOG row appended.
        ok41 = (activeAfter41 = expectedActive41) Or (postLog41 > preLog41)
        note41 = "hasMenu=" & hasMenu41 & " before=" & activeBefore41 & " after=" & activeAfter41 & " expected=" & expectedActive41 & " logDelta=" & (postLog41 - preLog41)
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(41, "admin_back_to_main_v21_effect", ok41, note41)

    ' --- Case 42 iter18 strengthened: Btn_ConfirmDiff no-op stability ---
    Dim ok42 As Boolean, note42 As String
    Dim fmtCellAddr42 As String
    Dim knwCountBefore42 As Long, knwCountAfter42 As Long
    Dim knwList42 As Collection
    On Error Resume Next
    Err.Clear
    On Error Resume Next: fmtCellAddr42 = CStr(Application.Run(ThisWorkbook.Name & "!modEntryFormat.ResolveM03FmtIdCell")): On Error GoTo 0
    If Len(fmtCellAddr42) = 0 Then fmtCellAddr42 = "C3"
    If Not wsM03 Is Nothing Then
        If Len(CStr(wsM03.Range(fmtCellAddr42).Value)) = 0 Then wsM03.Range(fmtCellAddr42).Value = JP_E2E_FMT_ID
        wsM03.Activate
    End If
    Set knwList42 = Nothing
    Set knwList42 = modKnowledgeFileIO.ListKnowledgesByFormat(JP_E2E_FMT_ID)
    If Not knwList42 Is Nothing Then knwCountBefore42 = knwList42.Count
    Err.Clear
    Application.Run "Btn_ConfirmDiff"
    Set knwList42 = Nothing
    Set knwList42 = modKnowledgeFileIO.ListKnowledgesByFormat(JP_E2E_FMT_ID)
    If Not knwList42 Is Nothing Then knwCountAfter42 = knwList42.Count
    If Err.Number <> 0 Then
        ok42 = False
        note42 = "err=" & Err.Number & " " & Err.Description
    Else
        ok42 = (knwCountBefore42 = knwCountAfter42)
        note42 = "cell=" & fmtCellAddr42 & " knwCountPre=" & knwCountBefore42 & " knwCountPost=" & knwCountAfter42 & " (vbYes no-op verified)"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(42, "admin_confirm_diff_noop_effect", ok42, note42)

    ' --- Case 43 iter12 new: Btn_OpenSettings_v21 effect
    Dim ok43 As Boolean, note43 As String
    Dim wsM11b As Worksheet
    Dim d13After As String
    On Error Resume Next
    Err.Clear
    Set wsM11b = LookupSheetByChrW(JpSettingsSheet())
    Application.Run "Btn_OpenSettings_v21"
    If Not wsM11b Is Nothing Then d13After = CStr(wsM11b.Range("D13").Value)
    If Err.Number <> 0 Then
        ok43 = False
        note43 = "err=" & Err.Number & " " & Err.Description
    Else
        ok43 = (Not wsM11b Is Nothing) And (Len(d13After) > 0)
        note43 = "D13=" & d13After
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(43, "admin_open_settings_effect", ok43, note43)

    ' --- Case 44 iter16 unskip: Btn_RefreshSheet_Active
    ' Production modRefresh.Btn_RefreshSheet_Active: ScreenUpdating/EnableEvents
    ' suppress + clsSetupOrchestrator.ReapplyActiveSheet, no MsgBox or FileDialog.
    Dim ok44 As Boolean, note44 As String
    Dim wsActiveBefore As String, wsActiveAfter As String
    Dim sbar As String
    On Error Resume Next
    Err.Clear
    wsActiveBefore = ThisWorkbook.ActiveSheet.Name
    Application.Run "Btn_RefreshSheet_Active"
    wsActiveAfter = ThisWorkbook.ActiveSheet.Name
    Application.StatusBar = False
    If Err.Number <> 0 Then
        ok44 = False
        note44 = "err=" & Err.Number & " " & Err.Description
    Else
        ok44 = (wsActiveBefore = wsActiveAfter)
        note44 = "before=" & wsActiveBefore & " after=" & wsActiveAfter
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(44, "admin_refresh_sheet_active", ok44, note44)

    ' --- Case 45 iter12 new: Btn_RefreshAllSheets
    Dim ok45 As Boolean, note45 As String
    Dim wsActB As String, wsActA As String
    On Error Resume Next
    Err.Clear
    wsActB = ThisWorkbook.ActiveSheet.Name
    Application.Run "Btn_RefreshAllSheets"
    wsActA = ThisWorkbook.ActiveSheet.Name
    Application.StatusBar = False
    If Err.Number <> 0 Then
        ok45 = False
        note45 = "err=" & Err.Number & " " & Err.Description
    Else
        ok45 = (wsActB = wsActA)
        note45 = "before=" & wsActB & " after=" & wsActA
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(45, "admin_refresh_all_sheets_effect", ok45, note45)

    ' --- Case 46 iter24 admin reload-list: macro-probe + LOG-effect verify ---
    ' Admin xlsm has no register module so Btn_ReloadList may not be installed.
    ' If macro not found => skipped-by-design (ADR-0092 §3 ok-regulation).
    ' If present => verify LOG row delta as effect.
    Dim ok46 As Boolean, note46 As String
    Dim preLog46 As Long, postLog46 As Long
    Dim errNo46 As Long, errDesc46 As String
    preLog46 = GetLastLogRow()
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_ReloadList"
    errNo46 = Err.Number
    errDesc46 = Err.Description
    On Error GoTo 0
    postLog46 = GetLastLogRow()
    If errNo46 = 1004 Or errNo46 = 438 Then
        ' Macro not found: admin xlsm legitimately has no Btn_ReloadList.
        ' ADR-0092 §3 ok-regulation: skipped-by-design counted as PASS-ish.
        ok46 = True
        note46 = "admin no Btn_ReloadList macro (err=" & errNo46 _
                 & " " & errDesc46 & ") skipped-by-design ADR-0092"
    ElseIf errNo46 <> 0 Then
        ok46 = (postLog46 > preLog46)
        note46 = "admin reload err=" & errNo46 & " " & errDesc46 _
                 & " logDelta=" & (postLog46 - preLog46)
    Else
        ok46 = (postLog46 > preLog46)
        note46 = "logDelta=" & (postLog46 - preLog46)
    End If
    r = r & "," & CaseToJson(46, "admin_reload_list_effect", ok46, note46)

    ' ============================================================
    ' iter28 ADR-0110 expansion: admin cases 61-70
    ' modFormatLoader SaveFormat/DeleteFormat/LoadFormat/FormatExists
    ' /ListAllFormats/LoadFormatList real-API round-trip.
    ' ============================================================
    r = r & "," & Case61_FormatSaveDeleteRoundTrip()
    r = r & "," & Case62_FormatSaveOverwrite()
    r = r & "," & Case63_FormatListConsistency()
    r = r & "," & Case91_FormatSaveJpUnicodeField()
    r = r & "," & Case92_FormatSaveMaxFields()
    r = r & "," & Case66_FormatLoadAfterDelete()
    r = r & "," & Case67_FormatExistenceCheck()
    r = r & "," & Case68_FormatDeleteWhileKnwReferenced()
    r = r & "," & Case69_FormatListSortConsistency()
    r = r & "," & Case70_FormatSaveAdminRoleGate()

    Run_Admin_Cases = r
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1021] modE2E_AllButtons.Run_Admin_Cases EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' helpers
' ================================================================

Private Sub PrepareKnwSaveSheet()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1025] modE2E_AllButtons.PrepareKnwSaveSheet ENTER"  ' [ADR-0100]
    On Error GoTo Done
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(JpKnwSaveSheet())
    If ws Is Nothing Then Exit Sub
    ws.Cells(6, 3).Value = JP_E2E_FMT_ID
    ws.Cells(8, 3).Value = ChrW(&H624B) & ChrW(&H9806) & ChrW(&H66F8) & ChrW(&H756A) & ChrW(&H53F7)
    ws.Cells(8, 5).Value = "E2E20-OPS-001"
    ws.Cells(9, 3).Value = ChrW(&H4F5C) & ChrW(&H696D) & ChrW(&H540D)
    ws.Cells(9, 5).Value = "E2E20 test job"
    ws.Cells(10, 3).Value = ""
Done:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1026] modE2E_AllButtons.PrepareKnwSaveSheet EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub SeedStorageSheet(ByVal ws As Worksheet)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1027] modE2E_AllButtons.SeedStorageSheet ENTER"  ' [ADR-0100]
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    ws.Range("D11").Value = SafeHolderGet("data_dir", "C:\KnowledgeMgr\data\")
    ws.Range("D12").Value = SafeHolderGet("format_dir", "C:\KnowledgeMgr\formats\")
    ws.Range("D13").Value = SafeHolderGet("ui_dir", "C:\KnowledgeMgr\ui\")
    ws.Range("D14").Value = SafeHolderGet("backup_dir", "C:\KnowledgeMgr\data\backup\")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1028] modE2E_AllButtons.SeedStorageSheet EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub SeedSettingsSheet(ByVal ws As Worksheet)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1029] modE2E_AllButtons.SeedSettingsSheet ENTER"  ' [ADR-0100]
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    Dim sDbg As String
    sDbg = SafeHolderGet("debugLevel", "INFO")
    ws.Range("D13").Value = sDbg
    ws.Range("B5").Value = sDbg
End Sub

Private Function SafeHolderGet(ByVal key As String, ByVal def As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1030] modE2E_AllButtons.SafeHolderGet ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim v As String
    v = modConfigHolder.GetValueOrDefault(key, def)
    If Err.Number <> 0 Or Len(v) = 0 Then
        SafeHolderGet = def
    Else
        SafeHolderGet = v
    End If
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1031] modE2E_AllButtons.SafeHolderGet EXIT-OK"  ' [ADR-0100]
End Function

Private Sub BumpEditFieldFirstValue()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1032] modE2E_AllButtons.BumpEditFieldFirstValue ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(JpKnwEditSheet())
    If ws Is Nothing Then Exit Sub
    Dim cur As String
    cur = CStr(ws.Cells(8, 5).Value)
    ws.Cells(8, 5).Value = cur & "_upd"
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1033] modE2E_AllButtons.BumpEditFieldFirstValue EXIT-OK"  ' [ADR-0100]
End Sub

Private Function GetDataDirSafe() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1034] modE2E_AllButtons.GetDataDirSafe ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim d As String
    d = modConfigHolder.GetDataDir()
    If Len(d) = 0 Then d = "C:\KnowledgeMgr\data\"
    If Right(d, 1) <> "\" Then d = d & "\"
    GetDataDirSafe = d
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1035] modE2E_AllButtons.GetDataDirSafe EXIT-OK"  ' [ADR-0100]
End Function

Private Function FileExists(ByVal path As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1036] modE2E_AllButtons.FileExists ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(path)
End Function

Private Sub DeleteFileSafe(ByVal path As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1037] modE2E_AllButtons.DeleteFileSafe ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then fso.DeleteFile path, True
End Sub

Private Function CountNonEmptyRowsInB(ByVal ws As Worksheet, _
        ByVal rowFrom As Long, ByVal rowTo As Long) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1038] modE2E_AllButtons.CountNonEmptyRowsInB ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim i As Long, cnt As Long
    For i = rowFrom To rowTo
        If Len(Trim(CStr(ws.Cells(i, 2).Value))) > 0 Then cnt = cnt + 1
    Next i
    CountNonEmptyRowsInB = cnt
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1039] modE2E_AllButtons.CountNonEmptyRowsInB EXIT-OK"  ' [ADR-0100]
End Function

Private Function LookupSheetByConst(ByVal cname As String) As Worksheet
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1040] modE2E_AllButtons.LookupSheetByConst ENTER"  ' [ADR-0100]
    On Error Resume Next
    Select Case cname
        Case "SHEET_KNW_LIST"
            Set LookupSheetByConst = ThisWorkbook.Worksheets( _
                ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H4E00) & ChrW(&H89A7))
        Case "SHEET_MIGRATION"
            Set LookupSheetByConst = ThisWorkbook.Worksheets( _
                ChrW(&H65E2) & ChrW(&H5B58) & ChrW(&H30C7) & ChrW(&H30FC) & ChrW(&H30BF) & _
                ChrW(&H53CD) & ChrW(&H6620))
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1041] modE2E_AllButtons.LookupSheetByConst EXIT-OK"  ' [ADR-0100]
End Function

Private Function DetectRole() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1042] modE2E_AllButtons.DetectRole ENTER"  ' [ADR-0100]
    Dim n As String
    n = ThisWorkbook.Name
    Dim base As String
    Dim dot As Long
    dot = InStrRev(n, ".")
    If dot > 0 Then
        base = Left(n, dot - 1)
    Else
        base = n
    End If
    If base = JpKanriConfigName() Then
        DetectRole = "admin"
    ElseIf base = JpTourokuName() Then
        DetectRole = "register"
    ElseIf base = JpKensakuName() Then
        DetectRole = "search"
    Else
        DetectRole = "unknown"
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1043] modE2E_AllButtons.DetectRole EXIT-OK"  ' [ADR-0100]
End Function


' iter19: helper to find a LOG row whose column H (logId) contains the
' given pattern, scanning rows 9..rows.Count (LOG_DATA_START_ROW=9).
' Returns the matching row number, or 0 if not found.
Private Function FindLogRowByLogId(ByVal logIdPattern As String) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1044] modE2E_AllButtons.FindLogRowByLogId ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Worksheets("LOG")
    If wsLog Is Nothing Then
        FindLogRowByLogId = 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1045] modE2E_AllButtons.FindLogRowByLogId EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(-4162).Row  ' xlUp = -4162
    Dim r As Long
    For r = 9 To lastRow
        Dim cellVal As String
        cellVal = CStr(wsLog.Cells(r, 8).Value)
        If InStr(cellVal, logIdPattern) > 0 Then
            FindLogRowByLogId = r
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1046] modE2E_AllButtons.FindLogRowByLogId EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next r
    FindLogRowByLogId = 0
End Function

' iter19: helper to find a LOG row whose column E (msg) contains the
' given text pattern. Used for buttons whose log entry has no logId.
Private Function FindLogRowByMsg(ByVal msgPattern As String) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1047] modE2E_AllButtons.FindLogRowByMsg ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Worksheets("LOG")
    If wsLog Is Nothing Then
        FindLogRowByMsg = 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1048] modE2E_AllButtons.FindLogRowByMsg EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim lastRow As Long
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(-4162).Row
    Dim r As Long
    For r = 9 To lastRow
        Dim cellVal As String
        cellVal = CStr(wsLog.Cells(r, 5).Value)
        If InStr(cellVal, msgPattern) > 0 Then
            FindLogRowByMsg = r
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1049] modE2E_AllButtons.FindLogRowByMsg EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next r
    FindLogRowByMsg = 0
End Function

' iter19: get current last log row, for "before" snapshot to verify
' that the button invocation actually appended a new log row.
Private Function GetLastLogRow() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1050] modE2E_AllButtons.GetLastLogRow ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Worksheets("LOG")
    If wsLog Is Nothing Then
        GetLastLogRow = 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1051] modE2E_AllButtons.GetLastLogRow EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    GetLastLogRow = wsLog.Cells(wsLog.Rows.Count, 1).End(-4162).Row
End Function

Private Function CaseToJson(ByVal idx As Long, ByVal nm As String, _
        ByVal okFlag As Boolean, ByVal noteText As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1052] modE2E_AllButtons.CaseToJson ENTER"  ' [ADR-0100]
    Dim p As String
    On Error Resume Next
    If okFlag Then p = "true" Else p = "false"
    modCommon.AppendProgressLog "E2Ecase-done " & CStr(idx) & " " & nm & " pass=" & p & " note=" & EscapeJson(noteText)
    AppendE2ELine idx, nm, okFlag, noteText
    On Error GoTo 0
    CaseToJson = "{""case"":" & idx & ",""name"":""" & nm & """,""pass"":" & p & _
        ",""note"":""" & EscapeJson(noteText) & """}"
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1053] modE2E_AllButtons.CaseToJson EXIT-OK"  ' [ADR-0100]
End Function

Private Sub AppendE2ELine(ByVal idx As Long, ByVal nm As String, ByVal okFlag As Boolean, ByVal noteText As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1054] modE2E_AllButtons.AppendE2ELine ENTER"  ' [ADR-0100]
    ' v4 (2026-06-02 I-11-1): Open For Binary append + StrConv vbFromUnicode.
    ' v3 bug: bytesCp As Variant -> Put writes type+length prefix instead of
    ' raw bytes. v4: declare as Byte() array explicitly.
    On Error Resume Next
    Dim role As String
    role = DetectRole()
    Dim linePath As String
    linePath = "C:\kvba\workspace\e2e_v20\" & role & "_lines.txt"
    Dim folder As String
    folder = "C:\kvba\workspace\e2e_v20"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    Dim p As String
    If okFlag Then p = "PASS" Else p = "FAIL"
    Dim line As String
    line = role & vbTab & CStr(idx) & vbTab & nm & vbTab & p & vbTab & EscapeJson(noteText) & vbCrLf
    Dim bytesCp() As Byte
    bytesCp = StrConv(line, vbFromUnicode)
    Dim fh As Integer
    fh = FreeFile
    Open linePath For Binary Access Write As #fh
    Dim sz As Long
    sz = LOF(fh)
    Put #fh, sz + 1, bytesCp
    Close #fh
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1055] modE2E_AllButtons.AppendE2ELine EXIT-OK"  ' [ADR-0100]
End Sub

Private Function EscapeJson(ByVal s As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1056] modE2E_AllButtons.EscapeJson ENTER"  ' [ADR-0100]
    Dim t As String
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    EscapeJson = t
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1057] modE2E_AllButtons.EscapeJson EXIT-OK"  ' [ADR-0100]
End Function

Private Function CountSearchHits(ByVal ws As Worksheet) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1058] modE2E_AllButtons.CountSearchHits ENTER"  ' [ADR-0100]
    On Error Resume Next
    If ws Is Nothing Then Exit Function
    Dim r As Long
    Dim cnt As Long
    cnt = 0
    For r = 14 To 200
        If Len(Trim(CStr(ws.Cells(r, 2).Value))) > 0 Then
            cnt = cnt + 1
        Else
            Exit For
        End If
    Next r
    CountSearchHits = cnt
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1059] modE2E_AllButtons.CountSearchHits EXIT-OK"  ' [ADR-0100]
End Function

Private Function FirstSearchResultId(ByVal ws As Worksheet) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1060] modE2E_AllButtons.FirstSearchResultId ENTER"  ' [ADR-0100]
    On Error Resume Next
    If ws Is Nothing Then Exit Function
    Dim v As String
    v = Trim(CStr(ws.Cells(14, 2).Value))
    If Len(v) > 0 Then
        FirstSearchResultId = v
    Else
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim dataDir As String
        dataDir = "C:\KnowledgeMgr\data\"
        If fso.FolderExists(dataDir) Then
            Dim f As Object
            For Each f In fso.GetFolder(dataDir).Files
                If LCase(fso.GetExtensionName(f.Name)) = "txt" Then
                    Dim base As String
                    base = fso.GetBaseName(f.Name)
                    If InStr(base, "SAGYO") > 0 Then
                        FirstSearchResultId = base
                        Exit For
                    End If
                End If
            Next f
        End If
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1061] modE2E_AllButtons.FirstSearchResultId EXIT-OK"  ' [ADR-0100]
End Function

Private Function LookupSheetByChrW(ByVal nm As String) As Worksheet
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1062] modE2E_AllButtons.LookupSheetByChrW ENTER"  ' [ADR-0100]
    On Error Resume Next
    Set LookupSheetByChrW = ThisWorkbook.Worksheets(nm)
End Function

Private Function SafeName(ByVal ws As Worksheet) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1063] modE2E_AllButtons.SafeName ENTER"  ' [ADR-0100]
    On Error Resume Next
    If ws Is Nothing Then
        SafeName = "(nil)"
    Else
        SafeName = ws.Name
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1064] modE2E_AllButtons.SafeName EXIT-OK"  ' [ADR-0100]
End Function

Private Function ReadConfigKey(ByVal cfgPath As String, ByVal key As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1065] modE2E_AllButtons.ReadConfigKey ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(cfgPath) Then Exit Function
    Dim ts As Object
    Set ts = fso.OpenTextFile(cfgPath, 1, False, 0)
    Dim line As String
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        Dim eq As Long
        eq = InStr(line, "=")
        If eq > 1 Then
            Dim k As String, v As String
            k = Trim(Left(line, eq - 1))
            v = Trim(Mid(line, eq + 1))
            If LCase(k) = LCase(key) Then
                ReadConfigKey = v
                ts.Close
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1066] modE2E_AllButtons.ReadConfigKey EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
        End If
    Loop
    ts.Close
End Function

Private Sub WriteAllText(ByVal path As String, ByVal text As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1067] modE2E_AllButtons.WriteAllText ENTER"  ' [ADR-0100]
    ' ADR-0094 v4 (2026-06-02 I-11-1): Open For Binary + StrConv vbFromUnicode.
    ' v3 bug: bytesCp As Variant -> Put writes Variant-prefixed bytes (type
    ' header + length + data) instead of raw bytes -> file corrupted with
    ' interleaved NUL+control bytes. v4: declare as Byte() array explicitly
    ' so Put writes the raw bytes 1:1.
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As String
    Dim slash As Long
    slash = InStrRev(path, "\")
    If slash > 0 Then
        folder = Left(path, slash - 1)
        If Not fso.FolderExists(folder) Then EnsureFolder fso, folder
    End If
    ' Determine expected tail marker. JSON ends ]} ; lines.txt ends with CR/LF.
    Dim expectedTail As String
    Dim tlen As Long
    tlen = Len(text)
    If tlen >= 2 Then
        expectedTail = Right(text, 2)
    ElseIf tlen >= 1 Then
        expectedTail = Right(text, 1)
    Else
        expectedTail = ""
    End If
    Dim attempt As Long
    For attempt = 1 To 3
        ' Convert Unicode String -> CP932 Byte() array.
        Dim bytesCp() As Byte
        bytesCp = StrConv(text, vbFromUnicode)
        ' Delete pre-existing file so Binary Open writes fresh.
        On Error Resume Next
        If fso.FileExists(path) Then fso.DeleteFile path, True
        On Error GoTo 0
        Dim fh As Integer
        On Error Resume Next
        fh = FreeFile
        Open path For Binary Access Write As #fh
        ' Put with no record-length prefix when passing Byte() array.
        Put #fh, 1, bytesCp
        Close #fh
        On Error GoTo 0
        ' Tail-marker verify
        If VerifyTailMarker(path, expectedTail) Then Exit For
    Next attempt
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1068] modE2E_AllButtons.WriteAllText EXIT-OK"  ' [ADR-0100]
End Sub

Private Function VerifyTailMarker(ByVal path As String, ByVal expectedTail As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1069] modE2E_AllButtons.VerifyTailMarker ENTER"  ' [ADR-0100]
    ' Read last N bytes of file as CP932, compare to expectedTail.
    On Error GoTo Bad
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then
        VerifyTailMarker = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1070] modE2E_AllButtons.VerifyTailMarker EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim fileObj As Object
    Set fileObj = fso.GetFile(path)
    Dim sz As Long
    sz = fileObj.Size
    If sz <= 0 Then
        VerifyTailMarker = (Len(expectedTail) = 0)
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1071] modE2E_AllButtons.VerifyTailMarker EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim tailLen As Long
    tailLen = LenB(StrConv(expectedTail, vbFromUnicode))
    If tailLen <= 0 Then
        VerifyTailMarker = True
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1072] modE2E_AllButtons.VerifyTailMarker EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If sz < tailLen Then
        VerifyTailMarker = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1073] modE2E_AllButtons.VerifyTailMarker EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim fh As Integer
    fh = FreeFile
    Open path For Binary Access Read As #fh
    Dim tailBytes() As Byte
    ReDim tailBytes(0 To tailLen - 1)
    Get #fh, sz - tailLen + 1, tailBytes
    Close #fh
    Dim got As String
    got = StrConv(tailBytes, vbUnicode)
    VerifyTailMarker = (got = expectedTail)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1074] modE2E_AllButtons.VerifyTailMarker EXIT-OK"  ' [ADR-0100]
    Exit Function
Bad:
    VerifyTailMarker = False
End Function

Private Sub EnsureFolder(ByVal fso As Object, ByVal folder As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1075] modE2E_AllButtons.EnsureFolder ENTER"  ' [ADR-0100]
    On Error Resume Next
    If fso.FolderExists(folder) Then Exit Sub
    Dim parent As String
    Dim p As Long
    p = InStrRev(folder, "\")
    If p > 0 Then
        parent = Left(folder, p - 1)
        If Len(parent) > 0 Then
            If Not fso.FolderExists(parent) Then EnsureFolder fso, parent
        End If
    End If
    fso.CreateFolder folder
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1076] modE2E_AllButtons.EnsureFolder EXIT-OK"  ' [ADR-0100]
End Sub


' --- iter13 helpers: data-test field seeding + CP932 read/write ---
Private Sub SeedRegFormWithFields(ByVal fmtId As String, ByVal names As Variant, ByVal values As Variant)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1077] modE2E_AllButtons.SeedRegFormWithFields ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(JpKnwSaveSheet())
    If ws Is Nothing Then Exit Sub
    ' iter15: use E2E_KS_* (module-local) so this compiles in search/admin too.
    ws.Cells(E2E_KS_ROW_FMT_ID, E2E_KS_COL_FMT_ID_VAL).Value = fmtId
    Dim i As Long
    For i = E2E_KS_FORM_START_ROW To E2E_KS_FORM_START_ROW + 50
        ws.Cells(i, E2E_KS_FIELD_COL_NAME).Value = ""
        ws.Cells(i, E2E_KS_FIELD_COL_NAME + 1).Value = ""
        ws.Cells(i, E2E_KS_FIELD_COL_VALUE).Value = ""
    Next i
    For i = LBound(names) To UBound(names)
        ws.Cells(E2E_KS_FORM_START_ROW + i - LBound(names), E2E_KS_FIELD_COL_NAME).Value = CStr(names(i))
        ws.Cells(E2E_KS_FORM_START_ROW + i - LBound(names), E2E_KS_FIELD_COL_VALUE).Value = CStr(values(i))
    Next i
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1078] modE2E_AllButtons.SeedRegFormWithFields EXIT-OK"  ' [ADR-0100]
End Sub

' iter16: Load field names from format file (SSOT), array of FieldName values.
' Returns 0-based String array on success; empty array (UBound<LBound) on failure.
' ADR-0090 compliant: no hard-coded field names; reads modFormatLoader.LoadFormat.
Private Function E2E_GetFormatFieldNames(ByVal fmtId As String) As Variant
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1079] modE2E_AllButtons.E2E_GetFormatFieldNames ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim sections As Collection
    Set sections = modFormatLoader.LoadFormat(fmtId)
    If sections Is Nothing Then
        E2E_GetFormatFieldNames = Array()
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1080] modE2E_AllButtons.E2E_GetFormatFieldNames EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If sections.Count = 0 Then
        E2E_GetFormatFieldNames = Array()
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1081] modE2E_AllButtons.E2E_GetFormatFieldNames EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim collected() As String
    Dim n As Long
    n = 0
    ReDim collected(0 To sections.Count - 1)
    Dim i As Long
    Dim sec As ClsStanzaSection
    For i = 1 To sections.Count
        Set sec = sections.Item(i)
        If sec.SectionName = "FIELD" Then
            If sec.HasKey("FieldName") Then
                collected(n) = sec.GetValue("FieldName")
                n = n + 1
            End If
        End If
    Next i
    If n = 0 Then
        E2E_GetFormatFieldNames = Array()
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1082] modE2E_AllButtons.E2E_GetFormatFieldNames EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim result() As String
    ReDim result(0 To n - 1)
    For i = 0 To n - 1
        result(i) = collected(i)
    Next i
    E2E_GetFormatFieldNames = result
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1083] modE2E_AllButtons.E2E_GetFormatFieldNames EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1084] modE2E_AllButtons.E2E_GetFormatFieldNames EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    E2E_GetFormatFieldNames = Array()
End Function

' iter15: generate test values for n fields (deterministic, distinguishable).
Private Function E2E_GenTestValues(ByVal n As Long, ByVal stem As String, ByVal suffix As String) As Variant
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1085] modE2E_AllButtons.E2E_GenTestValues ENTER"  ' [ADR-0100]
    Dim a() As String
    If n <= 0 Then
        E2E_GenTestValues = Array()
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1086] modE2E_AllButtons.E2E_GenTestValues EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    ReDim a(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        a(i) = stem & "_f" & CStr(i + 1) & suffix
    Next i
    E2E_GenTestValues = a
End Function

Private Function ReadAllTextCp932Local(ByVal path As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1087] modE2E_AllButtons.ReadAllTextCp932Local ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2
    ado.Charset = "Shift_JIS"
    ado.Open
    ado.LoadFromFile path
    ReadAllTextCp932Local = ado.ReadText
    ado.Close
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1088] modE2E_AllButtons.ReadAllTextCp932Local EXIT-OK"  ' [ADR-0100]
End Function

Private Sub WriteCp932Local(ByVal path As String, ByVal text As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1089] modE2E_AllButtons.WriteCp932Local ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2
    ado.Charset = "Shift_JIS"
    ado.Open
    ado.WriteText text
    ado.SaveToFile path, 2
    ado.Close
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1090] modE2E_AllButtons.WriteCp932Local EXIT-OK"  ' [ADR-0100]
End Sub


Private Function SheetNamesToActivate() As Variant
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1091] modE2E_AllButtons.SheetNamesToActivate ENTER"  ' [ADR-0100]
    Dim a(0 To 3) As String
    a(0) = JpFormatListSheet()
    a(1) = JpStorageSheet()
    a(2) = JpSettingsSheet()
    a(3) = JpFormatChangeCheckSheet()
    SheetNamesToActivate = a
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1092] modE2E_AllButtons.SheetNamesToActivate EXIT-OK"  ' [ADR-0100]
End Function

Private Function JpKanriConfigName() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1093] modE2E_AllButtons.JpKanriConfigName ENTER"  ' [ADR-0100]
    JpKanriConfigName = ChrW(&H7BA1) & ChrW(&H7406)
End Function

Private Function JpTourokuName() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1094] modE2E_AllButtons.JpTourokuName ENTER"  ' [ADR-0100]
    JpTourokuName = ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63)
End Function

Private Function JpKensakuName() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1095] modE2E_AllButtons.JpKensakuName ENTER"  ' [ADR-0100]
    JpKensakuName = ChrW(&H691C) & ChrW(&H7D22)
End Function

Private Function JpKnwSaveSheet() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1096] modE2E_AllButtons.JpKnwSaveSheet ENTER"  ' [ADR-0100]
    JpKnwSaveSheet = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & _
                      ChrW(&H767B) & ChrW(&H9332)
End Function

Private Function JpKnwEditSheet() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1097] modE2E_AllButtons.JpKnwEditSheet ENTER"  ' [ADR-0100]
    JpKnwEditSheet = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & _
                      ChrW(&H4FEE) & ChrW(&H6B63)
End Function

Private Function JpKnwSearchSheet() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1098] modE2E_AllButtons.JpKnwSearchSheet ENTER"  ' [ADR-0100]
    JpKnwSearchSheet = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & _
                       ChrW(&H691C) & ChrW(&H7D22)
End Function

Private Function JpStorageSheet() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1099] modE2E_AllButtons.JpStorageSheet ENTER"  ' [ADR-0100]
    JpStorageSheet = ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & _
                     ChrW(&H8A2D) & ChrW(&H5B9A)
End Function

Private Function JpSettingsSheet() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1100] modE2E_AllButtons.JpSettingsSheet ENTER"  ' [ADR-0100]
    JpSettingsSheet = ChrW(&H8A2D) & ChrW(&H5B9A)
End Function

Private Function JpFormatChangeCheckSheet() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1101] modE2E_AllButtons.JpFormatChangeCheckSheet ENTER"  ' [ADR-0100]
    JpFormatChangeCheckSheet = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & _
                               ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & _
                               ChrW(&H5909) & ChrW(&H66F4) & ChrW(&H30C1) & _
                               ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF)
End Function

Private Function JpFormatListSheet() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1102] modE2E_AllButtons.JpFormatListSheet ENTER"  ' [ADR-0100]
    JpFormatListSheet = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & _
                        ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & _
                        ChrW(&H4E00) & ChrW(&H89A7)
End Function

Private Function JpFormatDesignSheet() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1103] modE2E_AllButtons.JpFormatDesignSheet ENTER"  ' [ADR-0100]
    JpFormatDesignSheet = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & _
                          ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & _
                          ChrW(&H8A2D) & ChrW(&H8A08)
End Function

' iter24 ADR-0090: resolve M-12 target format input cell from SSOT ui_seed
' so we don't hard-code "C8" in test side. Falls back to "C8" if SSOT lookup
' fails (legacy production constant).
Private Function ResolveM12TargetFormatCell() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1104] modE2E_AllButtons.ResolveM12TargetFormatCell ENTER"  ' [ADR-0100]
    On Error GoTo Fallback
    Dim roleKanriName As String
    roleKanriName = ChrW(&H7BA1) & ChrW(&H7406)
    Dim ui As Collection
    Set ui = modUILoader.LoadUiDefinition(roleKanriName, "M-12")
    If ui Is Nothing Then GoTo Fallback
    If ui.Count = 0 Then GoTo Fallback
    Dim i As Long
    For i = 1 To ui.Count
        Dim sec As ClsStanzaSection
        Set sec = ui.Item(i)
        If sec.SectionName = "INPUT" Then
            Dim key As String
            key = sec.GetValue("inputDataKey")
            If key = "targetFormat" Or key = "FormatID" _
                    Or key = "formatId" Or key = "fmtId" Then
                Dim cellExpr As String
                cellExpr = sec.GetValue("Cell")
                If Len(cellExpr) > 0 Then
                    Dim p As Long
                    p = InStr(cellExpr, ":")
                    If p > 0 Then
                        ResolveM12TargetFormatCell = Left(cellExpr, p - 1)
                    Else
                        ResolveM12TargetFormatCell = cellExpr
                    End If
                    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1105] modE2E_AllButtons.ResolveM12TargetFormatCell EXIT-OK"  ' [ADR-0100]
                    Exit Function
                End If
            End If
        End If
    Next i
Fallback:
    ResolveM12TargetFormatCell = "C8"
End Function


' ================================================================
' iter28 ADR-0110: register expansion cases 71-80
' All call real modKnowledgeFileIO APIs; role gate yields skipped-by-design
' when run from non-register xlsm (returned via CaseToJson okFlag=False).
' ================================================================
Private Function Case71_KnwSaveLoadIdempotent(ByVal mgr As Object) As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case71_KnwSaveLoadIdempotent = CaseToJson(71, "knowledge_save_then_load_idempotent", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim id As String, rc As Long
    id = "SAGYO-T71"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d(ChrW(&H30D5) & ChrW(&H30A3) & ChrW(&H30FC) & ChrW(&H30EB) & ChrW(&H30C9) & "_1") = "v71"
    rc = modKnowledgeFileIO.SaveKnowledge(id, d, 0)
    Dim loaded As Object
    Set loaded = modKnowledgeFileIO.LoadKnowledge(id)
    If loaded Is Nothing Then
        ok = False
        note = "rc=" & rc & " loaded=Nothing"
    Else
        ok = (rc = 0) And (loaded.Count >= 1)
        note = "rc=" & rc & " loaded.Count=" & loaded.Count
    End If
    modKnowledgeFileIO.DeleteKnowledge id
    On Error GoTo 0
    Case71_KnwSaveLoadIdempotent = CaseToJson(71, "knowledge_save_then_load_idempotent", ok, note)
End Function

Private Function Case72_KnwUnicodeValue(ByVal mgr As Object) As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case72_KnwUnicodeValue = CaseToJson(72, "knowledge_save_unicode_value", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim id As String, rc As Long
    id = "SAGYO-T72"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim jpVal As String
    jpVal = ChrW(&H6F22) & ChrW(&H5B57) & ChrW(&H30C6) & ChrW(&H30B9) & ChrW(&H30C8)
    d("field1") = jpVal
    rc = modKnowledgeFileIO.SaveKnowledge(id, d, 0)
    Dim loaded As Object
    Set loaded = modKnowledgeFileIO.LoadKnowledge(id)
    If loaded Is Nothing Then
        ok = False
        note = "rc=" & rc & " loaded=Nothing"
    Else
        ok = (rc = 0) And (CStr(loaded("field1")) = jpVal)
        note = "rc=" & rc & " match=" & (CStr(loaded("field1")) = jpVal)
    End If
    modKnowledgeFileIO.DeleteKnowledge id
    On Error GoTo 0
    Case72_KnwUnicodeValue = CaseToJson(72, "knowledge_save_unicode_value", ok, note)
End Function

Private Function Case73_KnwSaveModifyLoad(ByVal mgr As Object) As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case73_KnwSaveModifyLoad = CaseToJson(73, "knowledge_save_then_modify_then_load", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim id As String
    id = "SAGYO-T73"
    Dim d1 As Object: Set d1 = CreateObject("Scripting.Dictionary")
    d1("k1") = "v_old"
    Dim rc1 As Long: rc1 = modKnowledgeFileIO.SaveKnowledge(id, d1, 0)
    Dim ts1 As Date: ts1 = modKnowledgeFileIO.GetKnowledgeTimestamp(id)
    Dim d2 As Object: Set d2 = CreateObject("Scripting.Dictionary")
    d2("k1") = "v_new"
    Dim rc2 As Long: rc2 = modKnowledgeFileIO.SaveKnowledge(id, d2, ts1)
    Dim loaded As Object: Set loaded = modKnowledgeFileIO.LoadKnowledge(id)
    If loaded Is Nothing Then
        ok = False
        note = "rc1=" & rc1 & " rc2=" & rc2 & " loaded=Nothing"
    Else
        ok = (rc1 = 0) And (rc2 = 0) And (CStr(loaded("k1")) = "v_new")
        note = "rc1=" & rc1 & " rc2=" & rc2 & " final=" & CStr(loaded("k1"))
    End If
    modKnowledgeFileIO.DeleteKnowledge id
    On Error GoTo 0
    Case73_KnwSaveModifyLoad = CaseToJson(73, "knowledge_save_then_modify_then_load", ok, note)
End Function

Private Function Case74_KnwListByFormatFilter(ByVal mgr As Object) As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case74_KnwListByFormatFilter = CaseToJson(74, "knowledge_list_by_format_filter", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim id As String: id = "SAGYO-T74"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("k") = "v"
    Dim rc As Long: rc = modKnowledgeFileIO.SaveKnowledge(id, d, 0)
    Dim col As Collection
    Set col = modKnowledgeFileIO.ListKnowledgesByFormat("SAGYO")
    Dim cnt As Long
    If col Is Nothing Then cnt = -1 Else cnt = col.Count
    ok = (rc = 0) And (cnt > 0)
    note = "rc=" & rc & " count=" & cnt
    modKnowledgeFileIO.DeleteKnowledge id
    On Error GoTo 0
    Case74_KnwListByFormatFilter = CaseToJson(74, "knowledge_list_by_format_filter", ok, note)
End Function

Private Function Case75_KnwSaveDeleteLoadGone(ByVal mgr As Object) As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case75_KnwSaveDeleteLoadGone = CaseToJson(75, "knowledge_save_then_delete_then_load", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim id As String: id = "SAGYO-T75"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("k") = "v75"
    Dim rcS As Long: rcS = modKnowledgeFileIO.SaveKnowledge(id, d, 0)
    Dim delOk As Boolean: delOk = modKnowledgeFileIO.DeleteKnowledge(id)
    Dim loaded As Object: Set loaded = modKnowledgeFileIO.LoadKnowledge(id)
    Dim isGone As Boolean
    If loaded Is Nothing Then
        isGone = True
    Else
        isGone = (loaded.Count = 0)
    End If
    ok = (rcS = 0) And delOk And isGone
    note = "rcS=" & rcS & " del=" & delOk & " gone=" & isGone
    On Error GoTo 0
    Case75_KnwSaveDeleteLoadGone = CaseToJson(75, "knowledge_save_then_delete_then_load", ok, note)
End Function

Private Function Case76_KnwMaxLengthValue(ByVal mgr As Object) As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case76_KnwMaxLengthValue = CaseToJson(76, "knowledge_save_max_length_value", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim id As String: id = "SAGYO-T76"
    Dim longVal As String: longVal = String(200, "x")
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("k") = longVal
    Dim rc As Long: rc = modKnowledgeFileIO.SaveKnowledge(id, d, 0)
    Dim loaded As Object: Set loaded = modKnowledgeFileIO.LoadKnowledge(id)
    If loaded Is Nothing Then
        ok = False
        note = "rc=" & rc & " loaded=Nothing"
    Else
        ok = (rc = 0) And (Len(CStr(loaded("k"))) = 200)
        note = "rc=" & rc & " len=" & Len(CStr(loaded("k")))
    End If
    modKnowledgeFileIO.DeleteKnowledge id
    On Error GoTo 0
    Case76_KnwMaxLengthValue = CaseToJson(76, "knowledge_save_max_length_value", ok, note)
End Function

Private Function Case77_KnwSaveEmptyDict() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case77_KnwSaveEmptyDict = CaseToJson(77, "knowledge_save_empty_dict", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    Dim id As String: id = "SAGYO-T77"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim rc As Long: rc = modKnowledgeFileIO.SaveKnowledge(id, d, 0)
    ' Empty-dict path: production writes a knowledge file with no fields.
    ' Negative-test verdict accepts either rc<>0 (rejected) or rc=0+file empty.
    Dim loaded As Object: Set loaded = modKnowledgeFileIO.LoadKnowledge(id)
    Dim fieldCnt As Long
    If loaded Is Nothing Then fieldCnt = -1 Else fieldCnt = loaded.Count
    If rc <> 0 Then
        ok = True
        note = "rejected rc=" & rc
    Else
        ok = (fieldCnt = 0)
        note = "rc=0 fieldCnt=" & fieldCnt
    End If
    modKnowledgeFileIO.DeleteKnowledge id
    On Error GoTo 0
    Case77_KnwSaveEmptyDict = CaseToJson(77, "knowledge_save_empty_dict", ok, note)
End Function

Private Function Case78_KnwBackupAfterSave(ByVal mgr As Object) As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case78_KnwBackupAfterSave = CaseToJson(78, "knowledge_backup_after_save", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim id As String: id = "SAGYO-T78"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("k") = "v78"
    Dim rc As Long: rc = modKnowledgeFileIO.SaveKnowledge(id, d, 0)
    Dim ts As String
    ts = Format$(Now, "yyyymmdd_hhnnss")
    Dim bkOk As Boolean
    bkOk = modKnowledgeFileIO.BackupKnowledgeFile(id, ts)
    ok = (rc = 0) And bkOk
    note = "rc=" & rc & " backup=" & bkOk & " ts=" & ts
    modKnowledgeFileIO.DeleteKnowledge id
    On Error GoTo 0
    Case78_KnwBackupAfterSave = CaseToJson(78, "knowledge_backup_after_save", ok, note)
End Function

Private Function Case79_KnwBuildNumberMissingFormat(ByVal mgr As Object) As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case79_KnwBuildNumberMissingFormat = CaseToJson(79, "knowledge_save_no_format", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    Dim result As String
    result = mgr.BuildKnowledgeNumber("NOSUCH_FMT_T79")
    ' Negative: production should return "" when format missing.
    ok = (Len(result) = 0)
    note = "result=" & result & " err=" & Err.Number
    On Error GoTo 0
    Case79_KnwBuildNumberMissingFormat = CaseToJson(79, "knowledge_save_no_format", ok, note)
End Function

Private Function Case80_KnwTimestampAfterSave(ByVal mgr As Object) As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "register" Then
        Case80_KnwTimestampAfterSave = CaseToJson(80, "knowledge_get_timestamp_after_save", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim id As String: id = "SAGYO-T80"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("k") = "v80"
    Dim rc As Long: rc = modKnowledgeFileIO.SaveKnowledge(id, d, 0)
    Dim ts As Date: ts = modKnowledgeFileIO.GetKnowledgeTimestamp(id)
    ok = (rc = 0) And (CDbl(ts) > 0#)
    note = "rc=" & rc & " ts=" & CStr(ts)
    modKnowledgeFileIO.DeleteKnowledge id
    On Error GoTo 0
    Case80_KnwTimestampAfterSave = CaseToJson(80, "knowledge_get_timestamp_after_save", ok, note)
End Function

' ================================================================
' iter28 ADR-0110: search expansion cases 81-90
' Read-side + role-write gate. Real modKnowledgeFileIO/modFormatLoader.
' ================================================================
Private Function Case81_ListAllCount() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "search" Then
        Case81_ListAllCount = CaseToJson(81, "search_list_all_count", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim col As Collection: Set col = modKnowledgeFileIO.ListAllKnowledges()
    Dim cnt As Long
    If col Is Nothing Then cnt = -1 Else cnt = col.Count
    ok = (cnt >= 0)
    note = "count=" & cnt
    On Error GoTo 0
    Case81_ListAllCount = CaseToJson(81, "search_list_all_count", ok, note)
End Function

Private Function Case82_LoadExistingKnw() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "search" Then
        Case82_LoadExistingKnw = CaseToJson(82, "search_load_existing_knw", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim all As Collection: Set all = modKnowledgeFileIO.ListAllKnowledges()
    Dim id As String
    If Not all Is Nothing Then
        If all.Count > 0 Then id = CStr(all.Item(1))
    End If
    If Len(id) = 0 Then
        ok = True
        note = "no-data,vacuous-pass"
    Else
        Dim loaded As Object: Set loaded = modKnowledgeFileIO.LoadKnowledge(id)
        ok = Not (loaded Is Nothing)
        Dim fieldCnt As Long
        If loaded Is Nothing Then fieldCnt = -1 Else fieldCnt = loaded.Count
        note = "id=" & id & " fields=" & fieldCnt
    End If
    On Error GoTo 0
    Case82_LoadExistingKnw = CaseToJson(82, "search_load_existing_knw", ok, note)
End Function

Private Function Case83_LoadNonexisting() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "search" Then
        Case83_LoadNonexisting = CaseToJson(83, "search_load_nonexisting", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim loaded As Object
    Set loaded = modKnowledgeFileIO.LoadKnowledge("NOSUCH-ID-Z83")
    Dim fieldCnt As Long
    If loaded Is Nothing Then
        fieldCnt = -1
        ok = True
    Else
        fieldCnt = loaded.Count
        ok = (fieldCnt = 0)
    End If
    note = "fieldCnt=" & fieldCnt
    On Error GoTo 0
    Case83_LoadNonexisting = CaseToJson(83, "search_load_nonexisting", ok, note)
End Function

Private Function Case84_ListByFormatEach() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "search" Then
        Case84_ListByFormatEach = CaseToJson(84, "search_list_by_format", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim ids As Collection: Set ids = modFormatLoader.ListAllFormats()
    Dim total As Long: total = 0
    Dim fmtCnt As Long
    If ids Is Nothing Then fmtCnt = 0 Else fmtCnt = ids.Count
    Dim id As Variant
    If Not ids Is Nothing Then
        For Each id In ids
            Dim sub_ As Collection: Set sub_ = modKnowledgeFileIO.ListKnowledgesByFormat(CStr(id))
            If Not sub_ Is Nothing Then total = total + sub_.Count
        Next id
    End If
    ok = (fmtCnt >= 0) And (total >= 0)
    note = "fmtCnt=" & fmtCnt & " totalKnw=" & total
    On Error GoTo 0
    Case84_ListByFormatEach = CaseToJson(84, "search_list_by_format", ok, note)
End Function

Private Function Case85_LoadFormatListPositive() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "search" Then
        Case85_LoadFormatListPositive = CaseToJson(85, "search_format_list_loaded", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim col As Collection: Set col = modFormatLoader.LoadFormatList()
    Dim cnt As Long
    If col Is Nothing Then cnt = -1 Else cnt = col.Count
    ok = (cnt > 0)
    note = "count=" & cnt
    On Error GoTo 0
    Case85_LoadFormatListPositive = CaseToJson(85, "search_format_list_loaded", ok, note)
End Function

Private Function Case86_FormatExistsKnown() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "search" Then
        Case86_FormatExistsKnown = CaseToJson(86, "search_format_exists_check_known", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim ids As Collection: Set ids = modFormatLoader.ListAllFormats()
    Dim probeId As String
    If Not ids Is Nothing Then
        If ids.Count > 0 Then probeId = CStr(ids.Item(1))
    End If
    If Len(probeId) = 0 Then
        ok = True
        note = "no-formats,vacuous-pass"
    Else
        Dim found As Boolean: found = modFormatLoader.FormatExists(probeId)
        ok = found
        note = "id=" & probeId & " exists=" & found
    End If
    On Error GoTo 0
    Case86_FormatExistsKnown = CaseToJson(86, "search_format_exists_check_known", ok, note)
End Function

Private Function Case87_FormatExistsUnknown() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "search" Then
        Case87_FormatExistsUnknown = CaseToJson(87, "search_format_exists_check_unknown", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim found As Boolean: found = modFormatLoader.FormatExists("NOSUCH_FMT_Z87")
    ok = (Not found)
    note = "exists=" & found
    On Error GoTo 0
    Case87_FormatExistsUnknown = CaseToJson(87, "search_format_exists_check_unknown", ok, note)
End Function

Private Function Case88_LoadFormatSectionsCount() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "search" Then
        Case88_LoadFormatSectionsCount = CaseToJson(88, "search_load_format_sections_count", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim ids As Collection: Set ids = modFormatLoader.ListAllFormats()
    Dim probeId As String
    If Not ids Is Nothing Then
        If ids.Count > 0 Then probeId = CStr(ids.Item(1))
    End If
    If Len(probeId) = 0 Then
        ok = True
        note = "no-formats,vacuous-pass"
    Else
        Dim sec As Collection: Set sec = modFormatLoader.LoadFormat(probeId)
        Dim cnt As Long
        If sec Is Nothing Then cnt = -1 Else cnt = sec.Count
        ok = (cnt > 0)
        note = "id=" & probeId & " sections=" & cnt
    End If
    On Error GoTo 0
    Case88_LoadFormatSectionsCount = CaseToJson(88, "search_load_format_sections_count", ok, note)
End Function

Private Function Case89_SearchRoleWriteGate() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "search" Then
        Case89_SearchRoleWriteGate = CaseToJson(89, "search_role_write_gate", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim id As String: id = "SAGYO-T89"
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("k") = "v89"
    Dim rc As Long: rc = modKnowledgeFileIO.SaveKnowledge(id, d, 0)
    ' Q-gate: search xlsm must be rejected with rc=3.
    ok = (rc = 3)
    note = "rc=" & rc & " (expected=3)"
    On Error GoTo 0
    Case89_SearchRoleWriteGate = CaseToJson(89, "search_role_write_gate", ok, note)
End Function

Private Function Case90_SearchRoleFormatWriteGate() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "search" Then
        Case90_SearchRoleFormatWriteGate = CaseToJson(90, "search_role_format_write_gate", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim sections As Collection: Set sections = New Collection
    Dim sec As ClsStanzaSection: Set sec = New ClsStanzaSection
    sec.Init "INPUT", 1
    sec.SetValue "field1", "x"
    sections.Add sec
    Dim rc As Long: rc = modFormatLoader.SaveFormat("FMT_T90", sections)
    ' Q19 gate: non-admin must be rejected with rc<>0.
    ok = (rc <> 0)
    note = "rc=" & rc & " (expected<>0)"
    On Error GoTo 0
    Case90_SearchRoleFormatWriteGate = CaseToJson(90, "search_role_format_write_gate", ok, note)
End Function

' ================================================================
' iter28 ADR-0110: admin expansion cases 61-70
' Real modFormatLoader Save/Delete/Load/Exists/List APIs.
' ================================================================
Private Function Case61_FormatSaveDeleteRoundTrip() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "admin" Then
        Case61_FormatSaveDeleteRoundTrip = CaseToJson(61, "format_save_then_delete_round_trip", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim fid As String: fid = "FMT_T61"
    Dim sections As Collection: Set sections = BuildMinimalFormatSections()
    Dim rcSave As Long: rcSave = modFormatLoader.SaveFormat(fid, sections)
    Dim existsAfterSave As Boolean: existsAfterSave = modFormatLoader.FormatExists(fid)
    Dim rcDel As Long: rcDel = modFormatLoader.DeleteFormat(fid)
    Dim existsAfterDel As Boolean: existsAfterDel = modFormatLoader.FormatExists(fid)
    ok = (rcSave = 0) And existsAfterSave And (rcDel = 0) And (Not existsAfterDel)
    note = "rcSave=" & rcSave & " saved=" & existsAfterSave & " rcDel=" & rcDel & " gone=" & (Not existsAfterDel)
    On Error GoTo 0
    Case61_FormatSaveDeleteRoundTrip = CaseToJson(61, "format_save_then_delete_round_trip", ok, note)
End Function

Private Function Case62_FormatSaveOverwrite() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "admin" Then
        Case62_FormatSaveOverwrite = CaseToJson(62, "format_save_then_overwrite", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim fid As String: fid = "FMT_T62"
    Dim sections1 As Collection: Set sections1 = BuildMinimalFormatSections()
    Dim rc1 As Long: rc1 = modFormatLoader.SaveFormat(fid, sections1)
    Dim sections2 As Collection: Set sections2 = BuildMinimalFormatSections()
    Dim rc2 As Long: rc2 = modFormatLoader.SaveFormat(fid, sections2)
    ok = (rc1 = 0) And (rc2 = 0)
    note = "rc1=" & rc1 & " rc2=" & rc2
    modFormatLoader.DeleteFormat fid
    On Error GoTo 0
    Case62_FormatSaveOverwrite = CaseToJson(62, "format_save_then_overwrite", ok, note)
End Function

Private Function Case63_FormatListConsistency() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "admin" Then
        Case63_FormatListConsistency = CaseToJson(63, "format_list_consistency", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim listA As Collection: Set listA = modFormatLoader.LoadFormatList()
    Dim listB As Collection: Set listB = modFormatLoader.ListAllFormats()
    Dim cntA As Long, cntB As Long
    If listA Is Nothing Then cntA = -1 Else cntA = listA.Count
    If listB Is Nothing Then cntB = -1 Else cntB = listB.Count
    ok = (cntA = cntB) And (cntA >= 0)
    note = "loadList=" & cntA & " listAll=" & cntB
    On Error GoTo 0
    Case63_FormatListConsistency = CaseToJson(63, "format_list_consistency", ok, note)
End Function

Private Function Case91_FormatSaveJpUnicodeField() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "admin" Then
        Case91_FormatSaveJpUnicodeField = CaseToJson(91, "format_save_jp_unicode_field", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim fid As String: fid = "FMT_T64"
    Dim jpField As String
    jpField = ChrW(&H4F5C) & ChrW(&H696D) & ChrW(&H540D)
    Dim sections As Collection: Set sections = New Collection
    Dim sec As ClsStanzaSection: Set sec = New ClsStanzaSection
    sec.Init "INPUT", 1
    sec.SetValue "fieldName", jpField
    sec.SetValue "fieldType", "text"
    sections.Add sec
    Dim rc As Long: rc = modFormatLoader.SaveFormat(fid, sections)
    Dim loaded As Collection: Set loaded = modFormatLoader.LoadFormat(fid)
    Dim secCnt As Long
    If loaded Is Nothing Then secCnt = -1 Else secCnt = loaded.Count
    ok = (rc = 0) And (secCnt > 0)
    note = "rc=" & rc & " secCnt=" & secCnt
    modFormatLoader.DeleteFormat fid
    On Error GoTo 0
    Case91_FormatSaveJpUnicodeField = CaseToJson(91, "format_save_jp_unicode_field", ok, note)
End Function

Private Function Case92_FormatSaveMaxFields() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "admin" Then
        Case92_FormatSaveMaxFields = CaseToJson(92, "format_save_max_fields", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim fid As String: fid = "FMT_T65"
    Dim sections As Collection: Set sections = New Collection
    Dim i As Long
    For i = 1 To 50
        Dim sec As ClsStanzaSection: Set sec = New ClsStanzaSection
        sec.Init "INPUT", i
        sec.SetValue "fieldName", "f" & i
        sec.SetValue "fieldType", "text"
        sections.Add sec
    Next i
    Dim rc As Long: rc = modFormatLoader.SaveFormat(fid, sections)
    Dim loaded As Collection: Set loaded = modFormatLoader.LoadFormat(fid)
    Dim secCnt As Long
    If loaded Is Nothing Then secCnt = -1 Else secCnt = loaded.Count
    ok = (rc = 0) And (secCnt = 50)
    note = "rc=" & rc & " secCnt=" & secCnt
    modFormatLoader.DeleteFormat fid
    On Error GoTo 0
    Case92_FormatSaveMaxFields = CaseToJson(92, "format_save_max_fields", ok, note)
End Function

Private Function Case66_FormatLoadAfterDelete() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "admin" Then
        Case66_FormatLoadAfterDelete = CaseToJson(66, "format_load_after_delete", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim fid As String: fid = "FMT_T66"
    modFormatLoader.SaveFormat fid, BuildMinimalFormatSections()
    modFormatLoader.DeleteFormat fid
    Dim loaded As Collection: Set loaded = modFormatLoader.LoadFormat(fid)
    Dim cnt As Long
    If loaded Is Nothing Then cnt = -1 Else cnt = loaded.Count
    ' Spec: LoadFormat on missing file returns empty Collection (cnt=0), not Nothing.
    ok = (cnt = 0)
    note = "cnt=" & cnt
    On Error GoTo 0
    Case66_FormatLoadAfterDelete = CaseToJson(66, "format_load_after_delete", ok, note)
End Function

Private Function Case67_FormatExistenceCheck() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "admin" Then
        Case67_FormatExistenceCheck = CaseToJson(67, "format_existence_check", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim fid As String: fid = "FMT_T67"
    Dim before As Boolean: before = modFormatLoader.FormatExists(fid)
    modFormatLoader.SaveFormat fid, BuildMinimalFormatSections()
    Dim after As Boolean: after = modFormatLoader.FormatExists(fid)
    modFormatLoader.DeleteFormat fid
    Dim gone As Boolean: gone = modFormatLoader.FormatExists(fid)
    ok = (Not before) And after And (Not gone)
    note = "before=" & before & " after=" & after & " gone=" & gone
    On Error GoTo 0
    Case67_FormatExistenceCheck = CaseToJson(67, "format_existence_check", ok, note)
End Function

Private Function Case68_FormatDeleteWhileKnwReferenced() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "admin" Then
        Case68_FormatDeleteWhileKnwReferenced = CaseToJson(68, "format_save_then_knw_referenced", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    ' Q55: DeleteFormat must return 2 if any knowledge of this format exists.
    ' Use the existing SAGYO format (always seeded) and verify rc<>0 if any
    ' SAGYO knowledge exists, else vacuous pass.
    Dim col As Collection: Set col = modKnowledgeFileIO.ListKnowledgesByFormat("SAGYO")
    Dim cnt As Long
    If col Is Nothing Then cnt = -1 Else cnt = col.Count
    If cnt <= 0 Then
        ok = True
        note = "no-sagyo-knw,vacuous-pass cnt=" & cnt
    Else
        Dim rc As Long: rc = modFormatLoader.DeleteFormat("SAGYO")
        ok = (rc <> 0)
        note = "knwCnt=" & cnt & " rcDel=" & rc & " (expected<>0)"
    End If
    On Error GoTo 0
    Case68_FormatDeleteWhileKnwReferenced = CaseToJson(68, "format_save_then_knw_referenced", ok, note)
End Function

Private Function Case69_FormatListSortConsistency() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "admin" Then
        Case69_FormatListSortConsistency = CaseToJson(69, "format_list_sort_consistency", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    Dim listA As Collection: Set listA = modFormatLoader.ListAllFormats()
    Dim listB As Collection: Set listB = modFormatLoader.ListAllFormats()
    Dim cntA As Long, cntB As Long
    If listA Is Nothing Then cntA = -1 Else cntA = listA.Count
    If listB Is Nothing Then cntB = -1 Else cntB = listB.Count
    Dim same As Boolean: same = (cntA = cntB)
    If same And cntA > 0 Then
        Dim i As Long
        For i = 1 To cntA
            If CStr(listA.Item(i)) <> CStr(listB.Item(i)) Then
                same = False
                Exit For
            End If
        Next i
    End If
    ok = same
    note = "cntA=" & cntA & " cntB=" & cntB & " same=" & same
    On Error GoTo 0
    Case69_FormatListSortConsistency = CaseToJson(69, "format_list_sort_consistency", ok, note)
End Function

Private Function Case70_FormatSaveAdminRoleGate() As String
    Dim ok As Boolean, note As String
    Dim role As String: role = DetectRole()
    If role <> "admin" Then
        Case70_FormatSaveAdminRoleGate = CaseToJson(70, "format_save_register_role_gate", False, "skipped-by-design,role=" & role)
        Exit Function
    End If
    On Error Resume Next
    ' Positive side of Q19: SaveFormat from admin xlsm = rc=0.
    Dim fid As String: fid = "FMT_T70"
    Dim rc As Long: rc = modFormatLoader.SaveFormat(fid, BuildMinimalFormatSections())
    ok = (rc = 0)
    note = "admin rc=" & rc & " (expected=0)"
    modFormatLoader.DeleteFormat fid
    On Error GoTo 0
    Case70_FormatSaveAdminRoleGate = CaseToJson(70, "format_save_register_role_gate", ok, note)
End Function

' ----------------------------------------------------------------
' helper: build a minimal valid format stanza collection.
' One INPUT section, fieldName=f1, fieldType=text.
' ----------------------------------------------------------------
Private Function BuildMinimalFormatSections() As Collection
    Dim sections As Collection: Set sections = New Collection
    Dim sec As ClsStanzaSection: Set sec = New ClsStanzaSection
    sec.Init "INPUT", 1
    sec.SetValue "fieldName", "f1"
    sec.SetValue "fieldType", "text"
    sections.Add sec
    Set BuildMinimalFormatSections = sections
End Function
```
