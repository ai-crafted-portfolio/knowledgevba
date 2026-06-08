---
title: modE2E_AllButtons.bas
description: modE2E_AllButtons.bas のソースコード（コピペ用）
---

# modE2E_AllButtons.bas

**配置先**: `共通モジュール (3 ブック全て)` 用の VBA モジュール  
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\`
- ファイル名: `modE2E_AllButtons.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

---

## ソースコード

```vb
Attribute VB_Name = "modE2E_AllButtons"
' ================================================================
' modE2E_AllButtons  (v2.3 E2E test entry, 2026-05-30)
' Purpose:
'   15-case headless E2E for register / search / admin button paths.
'   Exposes Run_E2E_AllButtons as the COM-callable entry. The role
'   is auto-detected from ThisWorkbook.Name (via ChrW comparisons
'   so the .bas file stays CP932/UTF-8 portable). Each role runs
'   its own case set, dumps a JSON result file, and the script that
'   drives this module aggregates all three roles.
'
' Notes:
'   - This module is test-only. The Sub names are unique so a user
'     running it in production produces a JSON file only (no side
'     effects beyond test-fixture create/delete inside data_dir).
'   - All Japanese strings are produced by ChrW (ADR-0006) so the
'     source stays CP932 strict without mojibake risk.
'   - JSON strings hold ASCII only by design; field names are ChrW
'     when needed but the JSON values written here are ASCII so the
'     result file is consumable by any reader.
' ================================================================
Option Explicit

Private Const JP_E2E_KNW_PREFIX As String = "E2E20_"
Private Const JP_E2E_FMT_ID As String = "SAGYO"

' ----------------------------------------------------------------
' Public Sub: Run_E2E_AllButtons
'   Entry point. outPath defaults to
'   C:\kvba\workspace\e2e_v20\<role>.json
'   The caller may override outPath via Application.Run.
' ----------------------------------------------------------------
Public Sub Run_E2E_AllButtons(Optional ByVal outPath As String = "")
    On Error GoTo ErrHandler
    Dim role As String
    role = DetectRole()

    Dim path As String
    path = outPath
    If Len(path) = 0 Then
        path = "C:\kvba\workspace\e2e_v20\" & role & ".json"
    End If

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
    Exit Sub
ErrHandler:
    On Error Resume Next
    Dim fbPath As String
    fbPath = outPath
    If Len(fbPath) = 0 Then fbPath = "C:\kvba\workspace\e2e_v20\unknown_role.json"
    WriteAllText fbPath, "{""role"":""error"",""err"":""" & EscapeJson(Err.Description) & """}"
End Sub

' ----------------------------------------------------------------
' Private Sub: Run_Register_Cases (returns JSON case fragments)
' ----------------------------------------------------------------
Private Function Run_Register_Cases() As String
    Dim r As String
    Dim mgr As Object
    Dim newId As String
    Dim ok As Boolean
    Dim note As String

    Set mgr = New clsKnowledgeManager
    mgr.Init Nothing, Nothing, ""

    ' --- Case 1: Btn_SaveKnowledge equivalent path = SaveNewKnowledge
    ' Pre-seed SHEET_KNW_SAVE so the manager can read fields.
    On Error Resume Next
    PrepareKnwSaveSheet
    newId = mgr.SaveNewKnowledge()
    note = ""
    If Err.Number <> 0 Then note = "err=" & Err.Number & " " & Err.Description
    On Error GoTo 0
    ok = (Len(newId) > 0)
    r = r & CaseToJson(1, "register_new", ok, "id=" & newId & " " & note)

    ' --- Case 2: Btn_LoadKnowledge equivalent path = LoadForEdit
    Dim loaded As Boolean
    On Error Resume Next
    loaded = False
    If ok Then loaded = mgr.LoadForEdit(newId)
    note = ""
    If Err.Number <> 0 Then note = "err=" & Err.Number & " " & Err.Description
    On Error GoTo 0
    r = r & "," & CaseToJson(2, "register_load", (ok And loaded), note)

    ' --- Case 3: Btn_UpdateKnowledge equivalent path = UpdateKnowledge
    Dim updated As Boolean
    On Error Resume Next
    updated = False
    If ok And loaded Then
        ' Modify one field cell in EDIT sheet then update.
        BumpEditFieldFirstValue
        updated = mgr.UpdateKnowledge(newId)
    End If
    note = ""
    If Err.Number <> 0 Then note = "err=" & Err.Number & " " & Err.Description
    On Error GoTo 0
    r = r & "," & CaseToJson(3, "register_update", (ok And loaded And updated), note)

    ' --- Case 4: Btn_DeleteKnowledge equivalent path = DeleteKnowledge
    Dim deleted As Boolean
    On Error Resume Next
    deleted = False
    If ok Then deleted = mgr.DeleteKnowledge(newId)
    note = ""
    If Err.Number <> 0 Then note = "err=" & Err.Number & " " & Err.Description
    On Error GoTo 0
    r = r & "," & CaseToJson(4, "register_delete", (ok And deleted), "id=" & newId & " " & note)

    ' --- Case 5: Btn_ClearForm equivalent. UserForm-modal path is
    ' skipped in headless; verify the sheet-level reset by calling
    ' Btn_ClearForm directly (sheet writeback only, no modal).
    Dim cleared As Boolean
    On Error Resume Next
    cleared = True
    Application.Run "Btn_ClearForm"
    If Err.Number <> 0 Then
        cleared = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        note = "Btn_ClearForm invoked"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(5, "register_clear", cleared, note)

    Run_Register_Cases = r
End Function

' ----------------------------------------------------------------
' Private Function: Run_Search_Cases
' ----------------------------------------------------------------
Private Function Run_Search_Cases() As String
    Dim r As String
    Dim wsName As String
    Dim ws As Worksheet
    Dim note As String
    Dim seededPath As String

    wsName = JpKnwSearchSheet()

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0

    ' iter6 fix (2026-05-31): seed a SAGYO knowledge file for the search
    ' role. register-role Case 4 deletes its SAGYO-0001 fixture, so
    ' search-role data_dir is otherwise empty for SAGYO. The seed file is
    ' deleted at the end of Run_Search_Cases (cleanup). Required by Case 8
    ' (doubleclick_view) which depends on Case 6/7 producing >=1 hit.
    seededPath = SeedSearchData()

    ' --- Case 6: Btn_SearchV23 with keyword=SAGYO (format filter)
    Dim hits1 As Long
    Dim ok1 As Boolean
    On Error Resume Next
    Err.Clear
    If Not ws Is Nothing Then
        ws.Range("C8").Value = JP_E2E_FMT_ID  ' search keyword = SAGYO (matches FormatID)
        ws.Range("C10").Value = ""            ' no format filter
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

    ' --- Case 7: Btn_SearchV23 with explicit format filter SAGYO
    Dim hits2 As Long
    Dim ok2 As Boolean
    On Error Resume Next
    Err.Clear
    If Not ws Is Nothing Then
        ws.Range("C8").Value = ""              ' keyword blank
        ws.Range("C10").Value = JP_E2E_FMT_ID  ' format filter = SAGYO
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

    ' --- Case 8: DoubleClick -> view path = OpenViewWithId (UserForm
    ' modal so we cannot click the dynamic form, but we can verify
    ' the entry sub exists and is callable; the UF render path is
    ' covered separately by clsUserFormRenderer test suite).
    ' Verify the public Sub is reachable via a candidate id pick from
    ' search hits. We do NOT call OpenViewWithId because it blocks on
    ' vbModal. Instead we exercise its prerequisite: data load round
    ' trip via clsKnowledgeManager.LoadKnowledge for the same id.
    Dim ok3 As Boolean
    Dim viewId As String
    viewId = FirstSearchResultId(ws)
    On Error Resume Next
    Err.Clear
    If Len(viewId) > 0 Then
        Dim mgr As Object
        Set mgr = New clsKnowledgeManager
        mgr.Init Nothing, Nothing, ""
        Dim d As Object
        Set d = mgr.LoadKnowledge(viewId)
        If Not d Is Nothing Then
            If d.Count > 0 Then ok3 = True
        End If
    End If
    If Err.Number <> 0 Then
        note = "err=" & Err.Number & " " & Err.Description
    Else
        note = "viewId=" & viewId & " (OpenViewWithId modal-skip, data load verified)"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(8, "search_doubleclick_view", ok3, note)

    ' --- Case 9: Close UF path. UF close is owned by the dynamic
    ' UserForm modal which we cannot drive headless. Verify the
    ' related clear button is callable (Btn_SearchClearV23) as the
    ' nearest equivalent for sheet-level reset.
    Dim ok4 As Boolean
    On Error Resume Next
    Err.Clear
    If Not ws Is Nothing Then
        Application.Run "Btn_SearchClearV23"
    End If
    If Err.Number <> 0 Then
        ok4 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok4 = True
        note = "Btn_SearchClearV23 invoked (UserForm-close modal-skip)"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(9, "search_close", ok4, note)

    ' iter6 fix (2026-05-31): cleanup the seed file we created at the
    ' top of Run_Search_Cases so the data_dir stays as it was.
    CleanupSearchData seededPath

    Run_Search_Cases = r
End Function

' iter6 helper (2026-05-31): write a minimal SAGYO knowledge file into
' data_dir so search-role Case 6/7/8 have at least one matching record.
Private Function SeedSearchData() As String
    On Error GoTo Bad
    Dim dataDir As String
    dataDir = SafeHolderGet("data_dir", "C:\KnowledgeMgr\data\")
    If Right$(dataDir, 1) <> "\" Then dataDir = dataDir & "\"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dataDir) Then
        EnsureFolder fso, dataDir
    End If
    Dim knwNo As String
    knwNo = "SAGYO-E2E20SEED"
    Dim filePath As String
    filePath = dataDir & knwNo & ".txt"
    Dim content As String
    content = "###FormatID###" & vbCrLf & "SAGYO" & vbCrLf
    content = content & "###CreatedAt###" & vbCrLf & Format$(Now, "yyyy-mm-dd HH:nn:ss") & vbCrLf
    content = content & "###UpdatedAt###" & vbCrLf & Format$(Now, "yyyy-mm-dd HH:nn:ss") & vbCrLf
    ' subject field name = KENMEI (U+4EF6 U+540D)
    content = content & "###" & ChrW(&H4EF6) & ChrW(&H540D) & "###" & vbCrLf & "SAGYO seed for E2E20" & vbCrLf
    ' excerpt field name = JISHO (U+4E8B U+8C61)
    content = content & "###" & ChrW(&H4E8B) & ChrW(&H8C61) & "###" & vbCrLf & "iter6 seed body" & vbCrLf
    Dim ts As Object
    Set ts = fso.CreateTextFile(filePath, True, False)
    ts.Write content
    ts.Close
    SeedSearchData = filePath
    Exit Function
Bad:
    SeedSearchData = ""
End Function

' iter6 helper (2026-05-31): delete the seed file SeedSearchData created.
Private Sub CleanupSearchData(ByVal seededPath As String)
    On Error Resume Next
    If Len(seededPath) = 0 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(seededPath) Then fso.DeleteFile seededPath, True
End Sub

' ----------------------------------------------------------------
' Private Function: Run_Admin_Cases
' ----------------------------------------------------------------
Private Function Run_Admin_Cases() As String
    Dim r As String
    Dim note As String

    ' --- Case 10: M-12 Btn_CheckFormat with C8 = SAGYO
    Dim ok10 As Boolean
    On Error Resume Next
    Err.Clear
    Dim wsM12 As Worksheet
    Set wsM12 = LookupSheetByChrW(JpFormatChangeCheckSheet())
    If Not wsM12 Is Nothing Then
        wsM12.Range("C8").Value = JP_E2E_FMT_ID
        Application.Run "Btn_CheckFormat"
    End If
    If Err.Number <> 0 Then
        note = "err=" & Err.Number & " " & Err.Description
    Else
        note = "sheet=" & SafeName(wsM12)
    End If
    ok10 = (Not wsM12 Is Nothing) And (Err.Number = 0)
    On Error GoTo 0
    r = r & CaseToJson(10, "admin_check_format", ok10, note)

    ' --- Case 11: M-10 Btn_OpenStorage_v21 modal-skip; instead
    ' verify the prerequisite path: storage sheet exists and has
    ' the 4 input cells D11..D14 reachable from the sheet, AND
    ' Btn_SaveStorage_v21 path can write/restore via SHEET_STORAGE.
    ' Pre-seed D11..D14 from config holder so the sheet view matches
    ' what Btn_OpenStorage_v21 would surface after FolderPicker use.
    Dim ok11 As Boolean
    Dim wsM10 As Worksheet
    Set wsM10 = LookupSheetByChrW(JpStorageSheet())
    Dim cell11 As String
    On Error Resume Next
    If Not wsM10 Is Nothing Then
        SeedStorageSheet wsM10
        cell11 = CStr(wsM10.Range("D11").Value)
    End If
    On Error GoTo 0
    ok11 = (Not wsM10 Is Nothing) And (Len(cell11) > 0)
    note = "sheet=" & SafeName(wsM10) & " D11=" & cell11 & " (FolderPicker modal-skip)"
    r = r & "," & CaseToJson(11, "admin_storage_open_or_check", ok11, note)

    ' --- Case 12: M-10 D11 write -> Btn_SaveStorage_v21 -> restore
    ' (SPEC_DRIFT-4b path: SaveConfigKeys writes config.txt)
    Dim ok12 As Boolean
    Dim origD11 As String
    Dim cfgPath As String
    Dim sentinel As String
    sentinel = "C:\kvba\workspace\e2e_v20\test_data_dir_" & Format$(Now, "yyyymmddhhnnss") & "\"
    cfgPath = "C:\KnowledgeMgr\" & JpKanriConfigName() & "_config.txt"
    On Error Resume Next
    Err.Clear
    If Not wsM10 Is Nothing Then
        ' Ensure D11..D14 are all non-empty so ValidateStorageSheet passes.
        SeedStorageSheet wsM10
        origD11 = CStr(wsM10.Range("D11").Value)
        wsM10.Range("D11").Value = sentinel
        Application.Run "Btn_SaveStorage_v21"
        ' Check config.txt has sentinel
        Dim curr As String
        curr = ReadConfigKey(cfgPath, "data_dir")
        If curr = sentinel Then ok12 = True
        ' Restore
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

    ' --- Case 13: M-11 Btn_SaveSettings_v21 (debugLevel round trip)
    Dim ok13 As Boolean
    Dim wsM11 As Worksheet
    Set wsM11 = LookupSheetByChrW(JpSettingsSheet())
    Dim origDbg As String
    Dim testDbg As String
    testDbg = "DEBUG"
    On Error Resume Next
    Err.Clear
    If Not wsM11 Is Nothing Then
        ' Ensure D13 (debugLevel, D13:E13 merged anchor) is non-empty
        ' before save so the ValidateSettingsSheet check inside
        ' Btn_SaveSettings_v21 passes.
        SeedSettingsSheet wsM11
        origDbg = CStr(wsM11.Range("D13").Value)
        wsM11.Range("D13").Value = testDbg
        Application.Run "Btn_SaveSettings_v21"
        Dim cdbg As String
        cdbg = ReadConfigKey(cfgPath, "debugLevel")
        If cdbg = testDbg Then ok13 = True
        ' Restore
        wsM11.Range("D13").Value = origDbg
        Application.Run "Btn_SaveSettings_v21"
    End If
    If Err.Number <> 0 Then
        note = "err=" & Err.Number & " " & Err.Description
    Else
        note = "testDbg=" & testDbg & " orig=" & origDbg
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(13, "admin_settings_save_restore", ok13, note)

    ' --- Case 14: Sheet switching (Worksheets(...).Activate path)
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

    ' --- Case 15: M-14 Btn_DeleteAllLog (headless guard: clear
    ' skipped + WARN log). Verify the sub is callable and that on
    ' headless it does not crash. The IsHeadless() returns True in
    ' COM mode, so this sub takes the WARN-log path (no destructive
    ' clear). PASS criterion: no exception raised.
    Dim ok15 As Boolean
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_DeleteAllLog"
    If Err.Number <> 0 Then
        ok15 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok15 = True
        note = "Btn_DeleteAllLog invoked (headless guard path)"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(15, "admin_delete_log", ok15, note)

    ' --- DRIFT-iter1-18 (iter2 add): M-03 format-design button wires ---
    ' Cases 16..20 cover Btn_NewFormatDraft / Btn_SaveFormat /
    ' Btn_LoadFormat / Btn_AddField / Btn_PreviewInDesign so a
    ' future wire-break (OnClick literal DRIFT, like the iter1 ui_seed
    ' v21 leftover) becomes detectable in headless E2E.
    ' All cases use Application.Run + On Error Resume Next so a
    ' Btn_* whose sub is missing reports note="err=..." and the
    ' regression bit flips. They do not pre-seed M-03 sheet state
    ' (M-03 button bodies write to ThisWorkbook sheets / format_dir
    ' themselves); the PASS criterion is "Btn_* callable without
    ' compile/runtime err".

    ' --- Case 16: Btn_NewFormatDraft
    Dim ok16 As Boolean
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_NewFormatDraft"
    If Err.Number <> 0 Then
        ok16 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok16 = True
        note = "Btn_NewFormatDraft invoked"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(16, "admin_m03_new_format_draft", ok16, note)

    ' --- Case 17: Btn_SaveFormat
    Dim ok17 As Boolean
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_SaveFormat"
    If Err.Number <> 0 Then
        ok17 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok17 = True
        note = "Btn_SaveFormat invoked"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(17, "admin_m03_save_format", ok17, note)

    ' --- Case 18: Btn_LoadFormat
    Dim ok18 As Boolean
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_LoadFormat"
    If Err.Number <> 0 Then
        ok18 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok18 = True
        note = "Btn_LoadFormat invoked"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(18, "admin_m03_load_format", ok18, note)

    ' --- Case 19: Btn_AddField
    Dim ok19 As Boolean
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_AddField"
    If Err.Number <> 0 Then
        ok19 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok19 = True
        note = "Btn_AddField invoked"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(19, "admin_m03_add_field", ok19, note)

    ' --- Case 20: Btn_PreviewInDesign
    Dim ok20 As Boolean
    On Error Resume Next
    Err.Clear
    Application.Run "Btn_PreviewInDesign"
    If Err.Number <> 0 Then
        ok20 = False
        note = "err=" & Err.Number & " " & Err.Description
    Else
        ok20 = True
        note = "Btn_PreviewInDesign invoked"
    End If
    On Error GoTo 0
    r = r & "," & CaseToJson(20, "admin_m03_preview_in_design", ok20, note)

    Run_Admin_Cases = r
End Function

' ================================================================
' ===== helpers ==================================================
' ================================================================

Private Sub PrepareKnwSaveSheet()
    On Error GoTo Done
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(JpKnwSaveSheet())
    If ws Is Nothing Then Exit Sub

    ' Format ID at (6, 3)
    ws.Cells(6, 3).Value = JP_E2E_FMT_ID

    ' Two minimal fields starting at row 8. Names from SAGYO format.
    ' Field name col = 3 (KS_FIELD_COL_NAME), value col = 5 (KS_FIELD_COL_VALUE)
    ' Use generic placeholder field names; values are ASCII.
    ws.Cells(8, 3).Value = ChrW(&H624B) & ChrW(&H9806) & ChrW(&H66F8) & ChrW(&H756A) & ChrW(&H53F7)  ' jp generic
    ws.Cells(8, 5).Value = "E2E20-OPS-001"
    ws.Cells(9, 3).Value = ChrW(&H4F5C) & ChrW(&H696D) & ChrW(&H540D)  ' generic
    ws.Cells(9, 5).Value = "E2E20 test job"
    ws.Cells(10, 3).Value = ""  ' end marker
Done:
End Sub

Private Sub SeedStorageSheet(ByVal ws As Worksheet)
    ' Mirror the holder->sheet path used by modEntrySettings.LoadStorageToSheet
    ' (which is Private, so test cannot call it directly). The 4 cell
    ' addresses are duplicated here intentionally; if the layout ever
    ' moves, update both this Seed helper and modEntrySettings.
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    ws.Range("D11").Value = SafeHolderGet("data_dir", "C:\KnowledgeMgr\data\")
    ws.Range("D12").Value = SafeHolderGet("format_dir", "C:\KnowledgeMgr\formats\")
    ws.Range("D13").Value = SafeHolderGet("ui_dir", "C:\KnowledgeMgr\ui\")
    ws.Range("D14").Value = SafeHolderGet("backup_dir", "C:\KnowledgeMgr\data\backup\")
End Sub

Private Sub SeedSettingsSheet(ByVal ws As Worksheet)
    ' Mirror modEntrySettings.LoadSettingsToSheet (Private). D13 holds
    ' debugLevel (D13:E13 merged anchor per ui_seed M-11.txt v2.3).
    ' Fall back to "INFO" when holder has nothing.
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    ws.Range("D13").Value = SafeHolderGet("debugLevel", "INFO")
End Sub

Private Function SafeHolderGet(ByVal key As String, ByVal def As String) As String
    On Error Resume Next
    Dim v As String
    v = modConfigHolder.GetValueOrDefault(key, def)
    If Err.Number <> 0 Or Len(v) = 0 Then
        SafeHolderGet = def
    Else
        SafeHolderGet = v
    End If
    On Error GoTo 0
End Function

Private Sub BumpEditFieldFirstValue()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(JpKnwEditSheet())
    If ws Is Nothing Then Exit Sub
    Dim cur As String
    cur = CStr(ws.Cells(8, 5).Value)
    ws.Cells(8, 5).Value = cur & "_upd"
End Sub

' ----------------------------------------------------------------
' Private Function: DetectRole
' Uses ThisWorkbook.Name. Names are produced via ChrW so the .bas
' is encoding-portable.
' ----------------------------------------------------------------
Private Function DetectRole() As String
    Dim n As String
    n = ThisWorkbook.Name
    ' Strip extension
    Dim base As String
    Dim dot As Long
    dot = InStrRev(n, ".")
    If dot > 0 Then
        base = Left$(n, dot - 1)
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
End Function

Private Function CaseToJson(ByVal idx As Long, ByVal nm As String, _
        ByVal okFlag As Boolean, ByVal noteText As String) As String
    Dim p As String
    If okFlag Then p = "true" Else p = "false"
    CaseToJson = "{""case"":" & idx & ",""name"":""" & nm & """,""pass"":" & p & _
        ",""note"":""" & EscapeJson(noteText) & """}"
End Function

Private Function EscapeJson(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    EscapeJson = t
End Function

Private Function CountSearchHits(ByVal ws As Worksheet) As Long
    On Error Resume Next
    If ws Is Nothing Then Exit Function
    Dim r As Long
    Dim cnt As Long
    cnt = 0
    For r = 14 To 200
        If Len(Trim$(CStr(ws.Cells(r, 2).Value))) > 0 Then
            cnt = cnt + 1
        Else
            Exit For
        End If
    Next r
    CountSearchHits = cnt
End Function

Private Function FirstSearchResultId(ByVal ws As Worksheet) As String
    On Error Resume Next
    If ws Is Nothing Then Exit Function
    Dim v As String
    v = Trim$(CStr(ws.Cells(14, 2).Value))
    If Len(v) > 0 Then
        FirstSearchResultId = v
    Else
        ' Fallback: look in data dir for first SAGYO knowledge file
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim dataDir As String
        dataDir = "C:\KnowledgeMgr\data\"
        If fso.FolderExists(dataDir) Then
            Dim f As Object
            For Each f In fso.GetFolder(dataDir).Files
                If LCase$(fso.GetExtensionName(f.Name)) = "txt" Then
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
End Function

Private Function LookupSheetByChrW(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set LookupSheetByChrW = ThisWorkbook.Worksheets(nm)
End Function

Private Function SafeName(ByVal ws As Worksheet) As String
    On Error Resume Next
    If ws Is Nothing Then
        SafeName = "(nil)"
    Else
        SafeName = ws.Name
    End If
End Function

Private Function ReadConfigKey(ByVal cfgPath As String, ByVal key As String) As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(cfgPath) Then Exit Function
    Dim ts As Object
    Set ts = fso.OpenTextFile(cfgPath, 1, False, 0)  ' ASCII
    Dim line As String
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        Dim eq As Long
        eq = InStr(line, "=")
        If eq > 1 Then
            Dim k As String, v As String
            k = Trim$(Left$(line, eq - 1))
            v = Trim$(Mid$(line, eq + 1))
            If LCase$(k) = LCase$(key) Then
                ReadConfigKey = v
                ts.Close
                Exit Function
            End If
        End If
    Loop
    ts.Close
End Function

Private Sub WriteAllText(ByVal path As String, ByVal text As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As String
    Dim slash As Long
    slash = InStrRev(path, "\")
    If slash > 0 Then
        folder = Left$(path, slash - 1)
        If Not fso.FolderExists(folder) Then
            Dim parent As String
            ' Naive recursive mkdir
            EnsureFolder fso, folder
        End If
    End If
    Dim ts As Object
    Set ts = fso.CreateTextFile(path, True, False)
    ts.Write text
    ts.Close
End Sub

Private Sub EnsureFolder(ByVal fso As Object, ByVal folder As String)
    On Error Resume Next
    If fso.FolderExists(folder) Then Exit Sub
    Dim parent As String
    Dim p As Long
    p = InStrRev(folder, "\")
    If p > 0 Then
        parent = Left$(folder, p - 1)
        If Len(parent) > 0 Then
            If Not fso.FolderExists(parent) Then EnsureFolder fso, parent
        End If
    End If
    fso.CreateFolder folder
End Sub

Private Function SheetNamesToActivate() As Variant
    Dim a(0 To 3) As String
    a(0) = JpFormatListSheet()       ' M-02
    a(1) = JpStorageSheet()          ' M-10
    a(2) = JpSettingsSheet()         ' M-11
    a(3) = JpFormatChangeCheckSheet() ' M-12
    SheetNamesToActivate = a
End Function

' ----- JP name functions (ChrW only, CP932 portable) -----

Private Function JpKanriConfigName() As String
    ' U+7BA1 U+7406 = kanri
    JpKanriConfigName = ChrW(&H7BA1) & ChrW(&H7406)
End Function

Private Function JpTourokuName() As String
    ' U+767B U+9332 U+4FEE U+6B63 = touroku-shusei
    JpTourokuName = ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63)
End Function

Private Function JpKensakuName() As String
    ' U+691C U+7D22 = kensaku
    JpKensakuName = ChrW(&H691C) & ChrW(&H7D22)
End Function

Private Function JpKnwSaveSheet() As String
    ' U+30CA U+30EC U+30C3 U+30B8 + U+767B U+9332
    JpKnwSaveSheet = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & _
                      ChrW(&H767B) & ChrW(&H9332)
End Function

Private Function JpKnwEditSheet() As String
    ' nareji shusei
    JpKnwEditSheet = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & _
                      ChrW(&H4FEE) & ChrW(&H6B63)
End Function

Private Function JpKnwSearchSheet() As String
    ' nareji kensaku
    JpKnwSearchSheet = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & _
                       ChrW(&H691C) & ChrW(&H7D22)
End Function

Private Function JpStorageSheet() As String
    ' kakuno-saki settei
    JpStorageSheet = ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & _
                     ChrW(&H8A2D) & ChrW(&H5B9A)
End Function

Private Function JpSettingsSheet() As String
    ' settei
    JpSettingsSheet = ChrW(&H8A2D) & ChrW(&H5B9A)
End Function

Private Function JpFormatChangeCheckSheet() As String
    ' format henkou check
    JpFormatChangeCheckSheet = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & _
                               ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & _
                               ChrW(&H5909) & ChrW(&H66F4) & ChrW(&H30C1) & _
                               ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF)
End Function

Private Function JpFormatListSheet() As String
    ' format ichiran
    JpFormatListSheet = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & _
                        ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & _
                        ChrW(&H4E00) & ChrW(&H89A7)
End Function
```
