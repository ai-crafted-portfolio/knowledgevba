---
title: modEntryFormat.bas
description: modEntryFormat.bas のソースコード（コピペ用）
---

# modEntryFormat.bas

**配置先**: `管理.xlsm` 用の VBA モジュール
**種類**: 標準モジュール

---

## ファイルとして保存

メモ帳（または任意のテキストエディタ）に下のソースコード全文を貼り付け、**`modEntryFormat.bas`** という名前で `installer\vba_modules\admin\` 配下に保存してください。文字コードは ANSI（Shift-JIS）、改行は CRLF にしてください。

---

## ソースコード


```vb
Attribute VB_Name = "modEntryFormat"
' ============================================================
' modEntryFormat (v2.3 Phase C, kanri.xlsm only)
'
' Role:
'   Format-list (M-02) and format-design (M-03) button handlers,
'   plus the shared format-dict bridge / workflow helpers used by
'   Btn_SaveFormat / Btn_LoadFormat.
'
' v2.3 scope (this phase):
'   M-02 (format list, 5 buttons): edit / preview / disable /
'                                  delete / reload
'   M-03 (format design, 5 buttons in 3 sets):
'                                  new-draft / save / load /
'                                  add-field / preview
'
' v2.3 changes vs. v2.2:
'   - New M-02 button: Btn_DisableFormat  (logical disable, sets
'     [FORMAT].Status = "????" in <format_dir>\<fmtId>.txt; the
'     format file is kept and existing knowledge stays untouched).
'   - New M-02 button: Btn_ReloadFormats  (rescans <format_dir>
'     and rewrites the M-02 grid data area).
'   - New M-03 button: Btn_NewFormatDraft (clears M-03 and seeds
'     the FormatID cell with the next auto-numbered fmtId; the
'     user can overwrite it before save).
'   - Row picking on M-02 is now checkbox-based (single check
'     required for edit / preview / disable / delete); zero or
'     multiple checks fall through with a warning.
'   - Removed Btn_CreateNewFormat (M-02 "new" button retired -
'     superseded by Btn_NewFormatDraft on M-03).
'   - Removed Btn_M03_FieldUp / Btn_M03_FieldDown /
'     Btn_M03_FieldEdit (per-row arrow / edit buttons on M-03
'     are dropped in v2.3 - reorder is done by editing the sort
'     column directly per the mock-design-rules).
'   - Removed SwapFieldRows (helper for the dropped per-row
'     buttons).
'   - Renamed Btn_ShowFormatPreview -> Btn_PreviewFormat (M-02
'     row-aware preview button, per v2.3 button table).
'   - Renamed (old) Btn_PreviewFormat -> Btn_PreviewInDesign
'     (M-03 "preview from the design screen" button).
'
' Notes:
'   - ASCII-only inside VBA string literals where practical, to
'     avoid CP932 mojibake when the .bas is round-tripped through
'     Edit. The single exception is the literal "????" which is
'     persisted into [FORMAT].Status and is therefore part of the
'     external contract (consumed by Btn_ReloadFormats as the
'     status text displayed on M-02 column L).
'   - Existing helpers (HasUnsavedInput_M03, NavigateToMain_v21,
'     RenderM04Preview, IsHeadless, NewLogger) are preserved
'     untouched - they are still referenced by Btn_BackToMain_v21
'     and the M-04 preview path.
' ============================================================
Option Explicit

Private Const FIELDS_KEY As String = "Fields"

' ------------------------------------------------------------
' ADR-0090 / ADR-0089 fix (2026-06-01): RoleKanri / M02_SHEET_NAME /
' M03_SHEET_NAME ChrW() helpers are defined at the bottom of the
' module (after the NormalizeFieldType Public Function). Keeping
' the module-level Const block contiguous - no Procedure between
' Const declarations - so the constants stay visible to every
' downstream Sub / Function.
' ------------------------------------------------------------

' ------------------------------------------------------------
' M-02 grid layout (v2.3, per screen_design_v2.md ??2.1 v2.3
' final-mock body).
'   Column B = row-check column (single-checkbox row select)
'   Column C = No
'   Column D = FormatID         <- picked by every M-02 button
'   Column E = FormatName
'   Column F = ID pattern
'   Column G = next number
'   Column H = format version
'   Column I = field count
'   Column J = knowledge count
'   Column K = updated-at
'   Column L = status text (e.g. "????")
'
' Data rows: M02_DATA_FIRST_ROW .. M02_DATA_LAST_ROW
' (The UI seed that physically lays out these cells is updated
' in a separate phase. The constants below are the v2.3 target
' layout used by the handlers in this module.)
' ------------------------------------------------------------
' ADR-0090 fix (2026-06-01): the previous CP932 literal got corrupted to
' "?t?H?[?}?b?g??" through a UTF-8 round-trip. The CJK literal is now
' exposed at the bottom of the module via M02_SHEET_NAME() ChrW helper.
Private Const M02_DATA_FIRST_ROW  As Long = 17
Private Const M02_DATA_LAST_ROW   As Long = 60
' 2026-06-07: M-02 column constants shifted -1 to match ui_seed M-02.txt
' Columns=Selected,No,FormatID,FormatName,IDPattern,NextNumber,FormatVersion,
' FieldCount,KnowledgeCount,UpdatedAt,IsActive (Selected starts at column A).
Private Const M02_COL_ROWCHECK    As String = "A"
Private Const M02_COL_NO          As String = "B"
Private Const M02_COL_FMT_ID      As String = "C"
Private Const M02_COL_FMT_NAME    As String = "D"
Private Const M02_COL_FIELD_COUNT As String = "E"
Private Const M02_COL_KNW_COUNT   As String = "F"
Private Const M02_COL_UPDATED     As String = "G"
' Removed 2026-06-08 (option 1): ID_PATTERN/NEXT_NUM/VERSION/STATUS were displayed
' but not implemented (no IDPattern/NextNumber/FormatVersion/IsActive metadata).

' Status text written into [FORMAT].Status when the user disables
' a format. The matching "enabled" text is intentionally not used
' inside this module - any non-disabled format displays whatever
' value the on-disk [FORMAT].Status carries (typically empty).
Private Const FMT_KEY_STATUS     As String = "Status"
Private Const FMT_STATUS_DISABLED As String = "????"   ' 2026-06-07: kept mojibake (Const cannot use ChrW); display layer maps ???? -> JP

' ------------------------------------------------------------
' M-03 cell layout (v2.3 - reusing the v2.2 UI seed positions
' that Btn_SaveFormat / Btn_LoadFormat already drive through
' modUILoader. The IDPattern slot (I3) is cleared on new-draft
' as a defensive step, even though the v2.3 design doc removes
' the IDPattern input; the seed cleanup keeps a stale value from
' surviving a "draft" reset.)
' ------------------------------------------------------------
' ADR-0090 fix (2026-06-01): CJK literal exposed via M03_SHEET_NAME()
' ChrW helper at the bottom of the module.
'
' ADR-0090 Phase 2 (2026-06-01, iter18): M03_CELL_FMT_ID was a hard-code
' "C3" used both here and in the E2E harness. It is now resolved via
' ResolveM03FmtIdCell() which reads ui_seed M-03.txt INPUT
' inputDataKey="targetFormat" / "FormatID" as SSOT, with a "C3" legacy
' fallback so existing installs continue to work even if the ui_seed
' INPUT key set drifts. ADR-0093 documents the ui_seed-vs-production
' Cell mismatch (ui_seed=C4:D4/C8:D8, production currently writes/reads
' C3) as a separate scope item; the helper preserves the production
' "C3" behaviour by treating the fallback as authoritative when the
' seeded key is not present.
Private Const M03_CELL_FMT_ID_FALLBACK As String = "C3"
' [v2.3 layout 2026-06-06] formatName C5:D5, GRID HeaderRow=18 so data starts row 19.
' (Pre-v2.3 values were E3 / 7 / 56 - old M-03 layout before SubHeader split.)
Private Const M03_CELL_FMT_NAME    As String = "C5"
Private Const M03_CELL_ID_PATTERN  As String = "I3"
Private Const M03_GRID_FIRST_ROW   As Long = 19
Private Const M03_GRID_LAST_ROW    As Long = 56
Private Const M03_GRID_CLEAR_FROM  As String = "A"   ' 2026-06-07: also clear column A (seq numbers)
Private Const M03_GRID_CLEAR_TO    As String = "G"

' ------------------------------------------------------------
' Auto-numbering for Btn_NewFormatDraft. The next fmtId is the
' lowest "FMT-NNN" not already present in <format_dir>. If the
' user has named their formats anything other than FMT-NNN this
' simply produces FMT-001 the first time (and the user can
' overwrite).
' ------------------------------------------------------------
Private Const FMT_ID_PREFIX As String = "FMT-"
Private Const FMT_ID_WIDTH  As Long = 3
Private Const FMT_ID_SCAN_LIMIT As Long = 999


' ============================================================
' ===== Bridge: dict <-> sheet cells (shared by Save/Load) =====
' ============================================================

Public Function BuildFormatDictFromCells(ByVal target As Object, ByVal uiSections As Collection) As Object
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0128] modEntryFormat.BuildFormatDictFromCells ENTER"  ' [ADR-0100]
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Set result(FIELDS_KEY) = New Collection
    If uiSections Is Nothing Then
        Set BuildFormatDictFromCells = result
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0129] modEntryFormat.BuildFormatDictFromCells EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim mode As String
    mode = LCase(modConfigHolder.GetValueOrDefault("uiSchemaFailMode", "warn"))
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To uiSections.Count
        Set sec = uiSections.Item(i)
        Select Case sec.SectionName
            Case "INPUT"
                Dim cellAddr As String, fieldName As String
                cellAddr = sec.GetValue("Cell")
                fieldName = sec.GetValue("Name")
                ' iter18 ADR-0090 fix: inputDataKey fallback (v2.3 ui_seed M-03.txt
                ' carries inputDataKey instead of "Name"; Build/Write previously
                ' skipped every INPUT and returned an empty fmtId).
                If Len(fieldName) = 0 Then fieldName = MapInputDataKeyToFieldName_M03(sec.GetValue("inputDataKey"))
                If Len(fieldName) > 0 And Len(cellAddr) > 0 Then
                    ' iter18b/iter19c: ui_seed M-03 has two INPUT stanzas for
                    ' "targetFormat" (C4:D4 new-draft + C8:D8 edit-set). The
                    ' iter18b implementation was last-non-empty-wins (comment
                    ' said first-non-empty but code wrote on every non-empty
                    ' hit), which let stale residue in the second cell override
                    ' a fresh value in the first. iter19c: actually preserve
                    ' the first non-empty read (skip subsequent assignments
                    ' once result already holds a non-empty value).
                    Dim cellVal As Variant
                    cellVal = clsCellIO.ReadCellValue(target, FirstCellOf_M03(cellAddr))
                    Dim cellValStr As String
                    If IsNull(cellVal) Then cellValStr = "" Else cellValStr = CStr(cellVal)
                    Dim existingStr As String
                    existingStr = ""
                    If result.Exists(fieldName) Then
                        If Not IsNull(result(fieldName)) Then existingStr = CStr(result(fieldName))
                    End If
                    If Len(existingStr) > 0 Then
                        ' Keep first-non-empty winner; do not overwrite.
                    ElseIf Len(cellValStr) > 0 Then
                        result(fieldName) = cellVal
                    ElseIf Not result.Exists(fieldName) Then
                        result(fieldName) = cellVal
                    End If
                ElseIf mode = "strict" Then
                    Err.Raise vbObjectError + 9201, "modEntryFormat.BuildFormatDictFromCells", _
                              "INPUT stanza missing Name or Cell key"
                End If
            Case "GRID"
                Set result(FIELDS_KEY) = clsGridIO.ReadGridFields(target, sec, mode)
        End Select
    Next i
    Set BuildFormatDictFromCells = result
End Function

Public Sub WriteFormatDictToCells(ByVal target As Object, ByVal uiSections As Collection, ByVal formatDict As Object)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0130] modEntryFormat.WriteFormatDictToCells ENTER"  ' [ADR-0100]
    If uiSections Is Nothing Or formatDict Is Nothing Then Exit Sub
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To uiSections.Count
        Set sec = uiSections.Item(i)
        Select Case sec.SectionName
            Case "INPUT"
                Dim cellAddr As String, fieldName As String
                cellAddr = sec.GetValue("Cell")
                fieldName = sec.GetValue("Name")
                ' iter18 ADR-0090 fix: same inputDataKey fallback as Build side.
                If Len(fieldName) = 0 Then fieldName = MapInputDataKeyToFieldName_M03(sec.GetValue("inputDataKey"))
                If Len(fieldName) > 0 And Len(cellAddr) > 0 Then
                    If formatDict.Exists(fieldName) Then
                        Dim v As Variant
                        v = formatDict(fieldName)
                        Dim writeCell As String
                        writeCell = FirstCellOf_M03(cellAddr)
                        If IsNull(v) Then
                            clsCellIO.WriteCellValue target, writeCell, ""
                        Else
                            clsCellIO.WriteCellValue target, writeCell, CStr(v)
                        End If
                    End If
                End If
            Case "GRID"
                If formatDict.Exists(FIELDS_KEY) Then
                    clsGridIO.WriteGridFields target, sec, formatDict(FIELDS_KEY)
                End If
        End Select
    Next i
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0131] modEntryFormat.WriteFormatDictToCells EXIT-OK"  ' [ADR-0100]
End Sub

Public Function SaveFormat_Workflow(ByVal target As Object, ByVal uiSections As Collection) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0132] modEntryFormat.SaveFormat_Workflow ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim fmtDict As Object
    Set fmtDict = BuildFormatDictFromCells(target, uiSections)
    Dim fmtId As String
    If fmtDict.Exists("FormatID") Then fmtId = CStr(fmtDict("FormatID"))
    If Len(fmtId) = 0 Then
        SaveFormat_Workflow = ""
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0133] modEntryFormat.SaveFormat_Workflow EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim sections As Collection
    Set sections = New Collection
    Dim fmtSec As ClsStanzaSection
    Set fmtSec = New ClsStanzaSection
    fmtSec.Init "FORMAT", 1
    fmtSec.SetValue "FormatID", fmtId
    If fmtDict.Exists("FormatName") Then fmtSec.SetValue "FormatName", CStr(fmtDict("FormatName"))
    sections.Add fmtSec
    If fmtDict.Exists(FIELDS_KEY) Then
        Dim fields As Collection
        Set fields = fmtDict(FIELDS_KEY)
        Dim i As Long
        For i = 1 To fields.Count
            Dim row As Object
            Set row = fields.Item(i)
            Dim fldSec As ClsStanzaSection
            Set fldSec = New ClsStanzaSection
            fldSec.Init "FIELD", 1
                        If row.Exists("Name") Then fldSec.SetValue "FieldName", CStr(row("Name"))
            ' Phase O-5 (ADR-0072 ?7.3): M-03 UI labels (6 ?: 1?s?e?L?X?g/???s?e?L?X?g/???l/???t/?I??/?`?F?b?N)
            ' ? canonical 4 base type (?P??s/???s/????/?I??) ???W?n??????? format.txt ??????B
            If row.Exists("Type") Then fldSec.SetValue "FieldType", NormalizeFieldType(CStr(row("Type")))
            If row.Exists("Required") Then fldSec.SetValue "Required", CStr(row("Required"))
            ' 2026-06-07: per-field Scroll / Rows / DropdownOptions also persisted
            ' so the renderer can honour them at preview / register / edit time.
            If row.Exists("scroll") Then fldSec.SetValue "Scroll", CStr(row("scroll"))
            If row.Exists("rows") Then fldSec.SetValue "Rows", CStr(row("rows"))
            If row.Exists("Options") Then fldSec.SetValue "DropdownOptions", CStr(row("Options"))
            sections.Add fldSec
        Next i
    End If
    Dim ret As Long
    ret = modFormatLoader.SaveFormat(fmtId, sections)
    If ret = 0 Then
        SaveFormat_Workflow = fmtId
        On Error Resume Next
        Dim oLogger006 As clsLogger
        Set oLogger006 = New clsLogger
        oLogger006.Init ThisWorkbook.Worksheets("LOG")
        oLogger006.LogInfo "modEntryFormat", "SaveFormat_Workflow", "Format save complete: " & fmtId, fmtId, "SAVE-EXIT-OK-II-006"
        On Error GoTo 0
    Else
        SaveFormat_Workflow = ""
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0134] modEntryFormat.SaveFormat_Workflow EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0135] modEntryFormat.SaveFormat_Workflow EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    On Error Resume Next
    Dim oLogger032 As clsLogger
    Set oLogger032 = New clsLogger
    oLogger032.Init ThisWorkbook.Worksheets("LOG")
    oLogger032.LogWarn "modEntryFormat", "SaveFormat_Workflow", "Format save failed: " & Err.Description, fmtId, "VALIDATE-WARN-WW-032"
    On Error GoTo 0
    SaveFormat_Workflow = ""
End Function

Public Function LoadFormat_Workflow(ByVal target As Object, ByVal uiSections As Collection, ByVal formatId As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0136] modEntryFormat.LoadFormat_Workflow ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If Len(formatId) = 0 Then
        LoadFormat_Workflow = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0137] modEntryFormat.LoadFormat_Workflow EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim loaded As Collection
    Set loaded = modFormatLoader.LoadFormat(formatId)
    If loaded.Count = 0 Then
        LoadFormat_Workflow = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0138] modEntryFormat.LoadFormat_Workflow EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim formatDict As Object
    Set formatDict = CreateObject("Scripting.Dictionary")
    Set formatDict(FIELDS_KEY) = New Collection
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To loaded.Count
        Set sec = loaded.Item(i)
        Select Case sec.SectionName
            Case "FORMAT"
                ' 2026-06-07: format files written by external tooling use
                ' "FormatId" (lowercase d); accept both spellings.
                If sec.HasKey("FormatID") Then
                    formatDict("FormatID") = sec.GetValue("FormatID")
                ElseIf sec.HasKey("FormatId") Then
                    formatDict("FormatID") = sec.GetValue("FormatId")
                End If
                If sec.HasKey("FormatName") Then formatDict("FormatName") = sec.GetValue("FormatName")
                If sec.HasKey("FormatVersion") Then formatDict("FormatVersion") = sec.GetValue("FormatVersion")
            Case "FIELD"
                Dim row As Object
                Set row = CreateObject("Scripting.Dictionary")
                If sec.HasKey("FieldName") Then row("Name") = sec.GetValue("FieldName")
                If sec.HasKey("FieldType") Then row("Type") = sec.GetValue("FieldType")
                If sec.HasKey("Required") Then row("Required") = sec.GetValue("Required")
                ' 2026-06-07: per-field Scroll / Rows / DropdownOptions round-trip
                If sec.HasKey("Scroll") Then row("scroll") = sec.GetValue("Scroll")
                If sec.HasKey("Rows") Then row("rows") = sec.GetValue("Rows")
                If sec.HasKey("DropdownOptions") Then row("Options") = sec.GetValue("DropdownOptions")
                formatDict(FIELDS_KEY).Add row
        End Select
    Next i
    WriteFormatDictToCells target, uiSections, formatDict
    LoadFormat_Workflow = True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0139] modEntryFormat.LoadFormat_Workflow EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0140] modEntryFormat.LoadFormat_Workflow EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    LoadFormat_Workflow = False
End Function


' ============================================================
' ===== M-03 wired button entry points =======================
' ============================================================

' ============================================================
' Public Sub: Btn_SaveFormat
' Role: M-03 "???" - writes the M-03 cells (basic info + field
'       grid) back to <format_dir>\<fmtId>.txt via the workflow
'       helper. Used by both "?V?K?N?[" and "?t?H?[?}?b?g?C??"
'       sets (same handler in v2.3).
' ============================================================
Public Sub Btn_SaveFormat()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0141] modEntryFormat.Btn_SaveFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_SaveFormat"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config;format_dir", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0142] modEntryFormat.Btn_SaveFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ui As Collection: Set ui = modUILoader.LoadUiDefinition(RoleKanri(), "M-03")
    If ui.Count = 0 Then Exit Sub
    ' [VALIDATION 2026-06-06]
    If Not IsHeadless() Then
        Dim wsSv As Worksheet
        Set wsSv = ThisWorkbook.Worksheets(M03_SHEET_NAME())
        Dim fmtName As String
        fmtName = Trim$(CStr(wsSv.Range(M03_CELL_FMT_NAME).Value))
        If Len(fmtName) = 0 Then
            MsgBox ChrW(&HFF7C) & ChrW(&HFF6A) & ChrW(&HFF70) & ChrW(&HFF8F) & ChrW(&HFF6F) & ChrW(&HFF84) & ChrW(&H540D) & ChrW(&H3092) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002), vbExclamation, ChrW(&H4FDD) & ChrW(&H5B58) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
            modBtnGuard.LogExit BTN, XLSM, False
            Exit Sub
        End If
        Dim gridCount As Long, rChk As Long
        gridCount = 0
        Dim firstErrRow As Long, firstErrMsg As String
        firstErrRow = 0
        For rChk = M03_GRID_FIRST_ROW To M03_GRID_LAST_ROW
            Dim fName As String, fType As String, fOpts As String, fRows As String
            fName = Trim$(CStr(wsSv.Cells(rChk, 2).Value))
            fType = Trim$(CStr(wsSv.Cells(rChk, 3).Value))
            fOpts = Trim$(CStr(wsSv.Cells(rChk, 6).Value))
            fRows = Trim$(CStr(wsSv.Cells(rChk, 5).Value))
            If Len(fName) > 0 Or Len(fType) > 0 Then
                gridCount = gridCount + 1
                If Len(fName) = 0 Or Len(fType) = 0 Then
                    If firstErrRow = 0 Then
                        firstErrRow = rChk
                        firstErrMsg = ChrW(&HFF8C) & ChrW(&HFF72) & ChrW(&HFF70) & ChrW(&HFF99) & ChrW(&HFF84) & ChrW(&H540D) & ChrW(&H3068) & ChrW(&H578B) & ChrW(&H3092) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H0028) & ChrW(&H884C) & CStr(rChk) & ChrW(&H0029)
                    End If
                End If
                If fType = ChrW(&H9078) & ChrW(&H629E) Then
                    If Len(fOpts) = 0 And firstErrRow = 0 Then
                        firstErrRow = rChk
                        firstErrMsg = ChrW(&H9078) & ChrW(&H629E) & ChrW(&H80A2) & ChrW(&H3092) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H0028) & ChrW(&H884C) & CStr(rChk) & ChrW(&H0029)
                    End If
                End If
                If fType = ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C) & ChrW(&HFF83) & ChrW(&HFF77) & ChrW(&HFF7D) & ChrW(&HFF84) Then
                    If (Len(fRows) = 0 Or Val(fRows) < 1) And firstErrRow = 0 Then
                        firstErrRow = rChk
                        firstErrMsg = ChrW(&H884C) & ChrW(&H6570) & ChrW(&H3092) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H0028) & ChrW(&H884C) & CStr(rChk) & ChrW(&H0029)
                    End If
                End If
            End If
        Next rChk
        If gridCount = 0 Then
            MsgBox ChrW(&H5C11) & ChrW(&H306A) & ChrW(&H304F) & ChrW(&H3068) & ChrW(&H3082) & ChrW(&H0031) & ChrW(&HFF8C) & ChrW(&HFF72) & ChrW(&HFF70) & ChrW(&HFF99) & ChrW(&HFF84) & ChrW(&H5FC5) & ChrW(&H8981) & ChrW(&H3067) & ChrW(&H3059) & ChrW(&H3002), vbExclamation, ChrW(&H4FDD) & ChrW(&H5B58) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
            modBtnGuard.LogExit BTN, XLSM, False
            Exit Sub
        End If
        If firstErrRow > 0 Then
            MsgBox firstErrMsg, vbExclamation, ChrW(&H4FDD) & ChrW(&H5B58) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
            modBtnGuard.LogExit BTN, XLSM, False
            Exit Sub
        End If
    End If
    SaveFormat_Workflow ActiveSheet, ui
    On Error Resume Next
    Dim oLogger007 As clsLogger
    Set oLogger007 = New clsLogger
    oLogger007.Init ThisWorkbook.Worksheets("LOG")
    oLogger007.LogInfo "modEntryFormat", "Btn_SaveFormat", "Btn_SaveFormat done", "", "SAVE-EXIT-OK-II-007"
    On Error GoTo 0
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0143] modEntryFormat.Btn_SaveFormat EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0144] modEntryFormat.Btn_SaveFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_SaveFormat: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_LoadFormat
' Role: M-03 "???" (?t?H?[?}?b?g?C???Z?b?g) - reads the fmtId
'       from the M-03 cell, loads <fmtId>.txt and writes it back
'       into the design cells.
' ============================================================
Public Sub Btn_LoadFormat()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0145] modEntryFormat.Btn_LoadFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_LoadFormat"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config;format_dir", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0146] modEntryFormat.Btn_LoadFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ui As Collection: Set ui = modUILoader.LoadUiDefinition(RoleKanri(), "M-03")
    If ui.Count = 0 Then Exit Sub
    ' iter19b: Unprotect M-03 before WriteGridFields. Setup_admin applies
    ' light ProtectSheet (same as Btn_NewFormatDraft already handles), and
    ' WriteCellValue against a protected cell silently no-ops under On Error
    ' Resume Next inside clsCellIO, leaving grid empty (case 18 post=0).
    On Error Resume Next
    ActiveSheet.Unprotect
    On Error GoTo ErrHandler
    Dim tempDict As Object: Set tempDict = BuildFormatDictFromCells(ActiveSheet, ui)
    Dim fmtId As String
    If tempDict.Exists("FormatID") Then fmtId = CStr(tempDict("FormatID"))
    If Len(fmtId) = 0 Then Exit Sub
    LoadFormat_Workflow ActiveSheet, ui, fmtId
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0147] modEntryFormat.Btn_LoadFormat EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0148] modEntryFormat.Btn_LoadFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_LoadFormat: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_NewFormatDraft  (v2.3 NEW)
' Role: M-03 "?V?K?N?[" - initialise the design screen for a new
'       draft. Seeds the FormatID cell with the next auto-
'       numbered fmtId (overridable), clears FormatName / other
'       inputs, and clears the field-design grid. Has no impact
'       on Btn_SaveFormat / Btn_LoadFormat / Btn_AddField /
'       Btn_PreviewInDesign behaviour.
'
'       Auto-numbering scans <format_dir>\*.txt for files matching
'       "FMT-NNN" and picks the lowest free index, padded to
'       FMT_ID_WIDTH digits. If no FMT-NNN file exists, the next
'       fmtId is FMT-001. The user can overwrite the cell value
'       before pressing Save.
' ============================================================
Public Sub Btn_NewFormatDraft()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0149] modEntryFormat.Btn_NewFormatDraft ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_NewFormatDraft"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config;format_dir", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0150] modEntryFormat.Btn_NewFormatDraft EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    ' [VALIDATION 2026-06-06] Check for unsaved design before discarding
    If Not IsHeadless() Then
        Dim wsV As Worksheet
        Set wsV = ThisWorkbook.Worksheets(M03_SHEET_NAME())
        Dim hasUnsaved As Boolean
        hasUnsaved = False
        On Error Resume Next
        If Len(Trim$(CStr(wsV.Range(M03_CELL_FMT_NAME).Value))) > 0 Then hasUnsaved = True
        Dim rChk As Long
        For rChk = M03_GRID_FIRST_ROW To M03_GRID_LAST_ROW
            If Len(Trim$(CStr(wsV.Cells(rChk, 2).Value))) > 0 Then
                hasUnsaved = True
                Exit For
            End If
        Next rChk
        On Error GoTo 0
        If hasUnsaved Then
            Dim ans As VbMsgBoxResult
            ans = MsgBox(ChrW(&H672A) & ChrW(&H4FDD) & ChrW(&H5B58) & ChrW(&H306E) & ChrW(&H8A2D) & ChrW(&H8A08) & ChrW(&H304C) & ChrW(&H3042) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H3002) & vbCrLf & ChrW(&H7834) & ChrW(&H68C4) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H65B0) & ChrW(&H898F) & ChrW(&H8D77) & ChrW(&H7968) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H304B) & ChrW(&H003F), _
                         vbYesNo + vbQuestion, _
                         ChrW(&H78BA) & ChrW(&H8A8D))
            If ans <> vbYes Then
                modBtnGuard.LogExit BTN, XLSM, False
                Exit Sub
            End If
        End If
    End If
    ' iter17 (2026-06-01): Robustified for headless E2E (case 16).
    On Error Resume Next
    Dim diagStep As String
    Dim errNum As Long
    Dim errDesc As String
    diagStep = "enter"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M03_SHEET_NAME())
    If ws Is Nothing Then
        Debug.Print "[ERR] Btn_NewFormatDraft: M-03 sheet not found"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0151] modEntryFormat.Btn_NewFormatDraft EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ' iter17 (2026-06-01): Unprotect the sheet before writing. Setup_admin
    ' applies sheet protection (ProtectSheet "light" via clsSetupOrchestrator),
    ' which blocks Range.Value=... with run-time error 1004. Other admin/search
    ' button handlers (Btn_SearchV23 / modUILoader writers / modSheetButtons)
    ' follow the same Unprotect-then-write pattern; we replicate it here so
    ' the seed lands instead of bouncing off the protection.
    ws.Unprotect

    ' 1. Clear the design inputs (single-cell ClearContents is safe
    '    against merged cells - merge is not broken).
    diagStep = "clear-fmtId"
    ws.Range(ResolveM03FmtIdCell()).ClearContents
    diagStep = "clear-fmtName"
    ws.Range(M03_CELL_FMT_NAME).ClearContents
    diagStep = "clear-idPattern"
    ws.Range(M03_CELL_ID_PATTERN).ClearContents

    ' 2. Clear the field-design grid data rows.
    diagStep = "clear-grid"
    Dim r As Long
    For r = M03_GRID_FIRST_ROW To M03_GRID_LAST_ROW
        ws.Range(M03_GRID_CLEAR_FROM & r & ":" & M03_GRID_CLEAR_TO & r).ClearContents
    Next r

    ' 3. Compute next auto-numbered fmtId. Fall back to FMT-001 if
    '    the helper hands back an empty string for any reason.
    diagStep = "autonumber"
    Dim nextId As String
    nextId = AutoNumberNextFmtId()
    If Len(nextId) = 0 Then nextId = FMT_ID_PREFIX & PadLeft("1", FMT_ID_WIDTH, "0")

    ' 4. Seed the FormatID cell BEFORE activating, so an Activate
    '    failure does not prevent the seed from landing.
    diagStep = "seed-fmtId"
    Err.Clear
    Dim seedCellAddr As String
    seedCellAddr = ResolveM03FmtIdCell()
    ws.Range(seedCellAddr).Value = nextId
    errNum = Err.Number
    errDesc = Err.Description
    If errNum <> 0 Then
        ws.Range(seedCellAddr).Value = "ERR " & diagStep & " " & errNum & " " & errDesc
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0152] modEntryFormat.Btn_NewFormatDraft EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ' 5. Activate the sheet (allowed to fail silently under headless).
    diagStep = "activate"
    Err.Clear
    ws.Activate

    ' 6. Log the result. Logger failures must not back out the seed.
    diagStep = "log"
    Err.Clear
    Dim lg As clsLogger
    Set lg = NewLogger()
    If Not lg Is Nothing Then
        lg.LogInfo "modEntryFormat", "Btn_NewFormatDraft", _
                   "M-03 new-draft seeded with fmtId=" & nextId, _
                   nextId, "M03-NEWDRAFT-OK-II-040"
    End If
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0153] modEntryFormat.Btn_NewFormatDraft EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0154] modEntryFormat.Btn_NewFormatDraft EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.HandleButtonError BTN, Err, XLSM
    modBtnGuard.LogExit BTN, XLSM, False
End Sub

' ============================================================
' Public Sub: Btn_AddField
' Role: M-03 "?t?B?[???h???" - selects the first empty row in
'       the field-design grid so the user can start typing.
' ============================================================
Public Sub Btn_AddField()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0155] modEntryFormat.Btn_AddField ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_AddField"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0156] modEntryFormat.Btn_AddField EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M03_SHEET_NAME())
    Dim r As Long
    For r = M03_GRID_FIRST_ROW To M03_GRID_LAST_ROW
        If Len(Trim(CStr(ws.Range("C" & r).Value))) = 0 Then
            ws.Activate
            On Error Resume Next
            ws.Range("C" & r).Select
            On Error GoTo 0
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0157] modEntryFormat.Btn_AddField EXIT-OK"  ' [ADR-0100]
            Exit Sub
        End If
    Next r
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0158] modEntryFormat.Btn_AddField EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0159] modEntryFormat.Btn_AddField EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_AddField: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_AddField_v21
' Role: legacy v2.1 entry kept for the wired button surface; it
'       hunts the first empty row in column B and selects it.
'       Btn_AddField (column-C scan) is the v2.3 replacement.
' ============================================================
Public Sub Btn_AddField_v21()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0160] modEntryFormat.Btn_AddField_v21 ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_AddField_v21"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0161] modEntryFormat.Btn_AddField_v21 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M03_SHEET_NAME())
    Const GRID_HEADER_ROW As Long = 18
    Const GRID_MAX_ROW As Long = 200
    ' columns: A=seq, B=fieldName, C=fieldType, D=required, E=rows, F=options, G=scroll
    Const COL_SEQ As Long = 1
    Const COL_NAME As Long = 2
    Const COL_TYPE As Long = 3
    Const COL_REQ As Long = 4
    Const COL_SCROLL As Long = 7
    ' v2.3 fix 2026-06-07: unprotect at start, re-protect at end so VBA Borders/
    ' Validation writes are not blocked by a Protect with UserInterfaceOnly=False
    ' (which happens after a save/reopen cycle). Errors are swallowed so any
    ' single failing call does not abort the whole add-row.
    On Error Resume Next
    ws.Unprotect
    On Error GoTo ErrHandler
    Dim r As Long
    For r = GRID_HEADER_ROW + 1 To GRID_MAX_ROW
        If Len(Trim$(CStr(ws.Cells(r, COL_NAME).Value))) = 0 Then
            ' Found next empty fieldName row. Ensure row has border + Locked=False + validation.
            Debug.Print "[Btn_AddField_v21] target row=" & r
            On Error Resume Next
            Dim rng As Range
            Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, 7))
            rng.Borders.LineStyle = xlContinuous
            rng.Borders.Weight = xlThin
            rng.Borders.Color = RGB(191, 191, 191)
            rng.Locked = False
            ' Add validation dropdowns
            Dim vType As Range
            Set vType = ws.Cells(r, COL_TYPE)
            vType.Validation.Delete
            vType.Validation.Add Type:=xlValidateList, Formula1:= _
                ChrW(&H0031) & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8) & "," & _
                ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8) & "," & _
                ChrW(&H6570) & ChrW(&H5024) & "," & _
                ChrW(&H65E5) & ChrW(&H4ED8) & "," & _
                ChrW(&H9078) & ChrW(&H629E) & "," & _
                ChrW(&H30C1) & ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF)
            vType.Validation.InCellDropdown = True
            vType.Validation.IgnoreBlank = True
            Dim vReq As Range
            Set vReq = ws.Cells(r, COL_REQ)
            vReq.Validation.Delete
            vReq.Validation.Add Type:=xlValidateList, Formula1:="TRUE,FALSE"
            vReq.Validation.InCellDropdown = True
            vReq.Validation.IgnoreBlank = True
            Dim vScr As Range
            Set vScr = ws.Cells(r, COL_SCROLL)
            vScr.Validation.Delete
            vScr.Validation.Add Type:=xlValidateList, Formula1:="TRUE,FALSE"
            vScr.Validation.InCellDropdown = True
            vScr.Validation.IgnoreBlank = True
            ' Auto-number sequence (most critical write - if this fails the row
            ' appears blank to the user even though Borders succeeded).
            ws.Cells(r, COL_SEQ).Value = r - GRID_HEADER_ROW
            ' Navigate cursor
            ws.Cells(r, COL_NAME).Select
            On Error GoTo 0
            ' Re-apply light protection so the rest of the sheet stays locked.
            On Error Resume Next
            ws.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True, _
                       AllowFormattingColumns:=True, AllowFormattingRows:=True
            On Error GoTo ErrHandler
            Debug.Print "[Btn_AddField_v21] added row " & r
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0162] modEntryFormat.Btn_AddField_v21 EXIT-OK"  ' [ADR-0100]
            Exit Sub
        End If
    Next r
    Debug.Print "[Btn_AddField_v21] GRID full, no action"
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0163] modEntryFormat.Btn_AddField_v21 EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0164] modEntryFormat.Btn_AddField_v21 EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[Btn_AddField_v21] error: " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_PreviewInDesign
' Role: M-03 "?v???r???[" - activates the M-04 preview sheet
'       and triggers a render. The M-03 design cells provide the
'       source format for the preview (M-04 reads via the
'       [GRID_FROM_FORMAT] stanza which references M-03!C3).
'       v2.3 rename: was Btn_PreviewFormat in v2.2.
' ============================================================
Public Sub Btn_PreviewInDesign()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0165] modEntryFormat.Btn_PreviewInDesign ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_PreviewInDesign"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0166] modEntryFormat.Btn_PreviewInDesign EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    ' [USERFORM 2026-06-06] Launch dynamic UserForm in preview mode
    ' (same pattern as modEntryUserForm.OpenViewWithId for M-08 search).
    Dim fmtId As String
    On Error Resume Next
    Dim fmtIdC4 As String, fmtIdC8 As String
    fmtIdC4 = Trim$(CStr(ThisWorkbook.Worksheets(M03_SHEET_NAME()).Range("C4").Value))
    fmtIdC8 = Trim$(CStr(ThisWorkbook.Worksheets(M03_SHEET_NAME()).Range("C8").Value))
    ' 2026-06-07: prefer the side whose format file actually exists.
    ' C4 is auto-seeded so its file may not yet exist; in that case
    ' fall back to C8 (modify side).
    Dim fmtDir As String
    fmtDir = modConfigHolder.GetValueOrDefault("format_dir", "")
    fmtId = ""
    If Len(fmtIdC4) > 0 Then
        If Len(fmtDir) > 0 Then
            Dim fso_pv As Object
            Set fso_pv = CreateObject("Scripting.FileSystemObject")
            If fso_pv.FileExists(fmtDir & fmtIdC4 & ".txt") Then
                fmtId = fmtIdC4
            End If
        Else
            fmtId = fmtIdC4
        End If
    End If
    If Len(fmtId) = 0 And Len(fmtIdC8) > 0 Then fmtId = fmtIdC8
    If Len(fmtId) = 0 Then fmtId = fmtIdC4
    On Error GoTo ErrHandler
    If IsHeadless() Then
        ThisWorkbook.Worksheets(PreviewSheetName_M04()).Activate
        Exit Sub
    End If
    Dim r As clsUserFormRenderer
    Set r = New clsUserFormRenderer
    Dim ret As String
    ret = r.ShowForm(ChrW(&H7BA1) & ChrW(&H7406), "preview", "", True, fmtId)
    Debug.Print "[Btn_PreviewInDesign] ShowForm ret=" & ret & " fmtId=" & fmtId
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0167] modEntryFormat.Btn_PreviewInDesign EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0168] modEntryFormat.Btn_PreviewInDesign EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_PreviewInDesign: " & Err.Number & " " & Err.Description
End Sub


' ============================================================
' ===== M-02 wired button entry points =======================
' ============================================================

' ============================================================
' Public Sub: Btn_EditFormat
' Role: M-02 "??W" - read the fmtId of the single checked row
'       and open it on M-03 via LoadFormat_Workflow.
' ============================================================
Public Sub Btn_EditFormat()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0169] modEntryFormat.Btn_EditFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_EditFormat"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config;format_dir", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0170] modEntryFormat.Btn_EditFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim fmtId As String
    fmtId = PickM02CheckedRowFmtId()
    If Len(fmtId) = 0 Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M03_SHEET_NAME())
    ws.Activate
    Dim ui As Collection
    Set ui = modUILoader.LoadUiDefinition(RoleKanri(), "M-03")
    If ui.Count = 0 Then Exit Sub
    LoadFormat_Workflow ws, ui, fmtId
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0171] modEntryFormat.Btn_EditFormat EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0172] modEntryFormat.Btn_EditFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_EditFormat: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_PreviewFormat
' Role: M-02 "?v???r???[" - read the fmtId of the single checked
'       row, plant it into M-03!C3 (so M-04's [GRID_FROM_FORMAT]
'       can resolve SourceFormatRef=M-03!C3), activate M-04 and
'       trigger the render.
'       v2.3 rename: was Btn_ShowFormatPreview in v2.2.
' ============================================================
Public Sub Btn_PreviewFormat()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0173] modEntryFormat.Btn_PreviewFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_PreviewFormat"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config;format_dir", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0174] modEntryFormat.Btn_PreviewFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim fmtId As String
    fmtId = PickM02CheckedRowFmtId()
    If Len(fmtId) = 0 Then Exit Sub
    ThisWorkbook.Worksheets(M03_SHEET_NAME()).Range(ResolveM03FmtIdCell()).Value = fmtId
    ' iter19 ADR-0090: CP932 mojibake "?v???r???[" repaired via ChrW helper
    ' PreviewSheetName_M04() (same fix as Btn_PreviewInDesign).
    ThisWorkbook.Worksheets(PreviewSheetName_M04()).Activate
    ' iter19b: skip RenderM04Preview under headless E2E (same blocker as
    ' Btn_PreviewInDesign, see case 20 timeout).
    If IsHeadless() Then Exit Sub
    RenderM04Preview
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0175] modEntryFormat.Btn_PreviewFormat EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0176] modEntryFormat.Btn_PreviewFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_PreviewFormat: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_DisableFormat  (v2.3 NEW)
' Role: M-02 "??????" - logically disables the selected format
'       by writing/overwriting [FORMAT].Status="????" into the
'       on-disk <format_dir>\<fmtId>.txt. The format file itself
'       and any knowledge entries that reference the format are
'       left intact - this is NOT a delete. Disabling simply
'       removes the format from the "new knowledge entry" choice
'       list; existing knowledge keeps working as-is.
'
'       Even formats whose knowledge count is >= 1 can be
'       disabled (this is the documented behaviour).
'
'       Row selection: single checked row on M-02 (zero or >1
'       checks raise a warning and abort).
'
'       After save, the M-02 grid is refreshed via
'       Btn_ReloadFormats so column L (status) reflects the new
'       value immediately.
' ============================================================
Public Sub Btn_DisableFormat()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0177] modEntryFormat.Btn_DisableFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_DisableFormat"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config;format_dir", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0178] modEntryFormat.Btn_DisableFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]

    ' 1. Pick the single checked row.
    Dim fmtId As String
    fmtId = PickM02CheckedRowFmtId()
    If Len(fmtId) = 0 Then Exit Sub

    ' 2. Load the existing format definition.
    Dim sections As Collection
    Set sections = modFormatLoader.LoadFormat(fmtId)
    If sections Is Nothing Then
        WarnUser ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H7121) & ChrW(&H52B9) & ChrW(&H5316), _
                 ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H0020) & fmtId & ChrW(&H0020) & ChrW(&H3092) & ChrW(&H8AAD) & ChrW(&H307F) & ChrW(&H8FBC) & ChrW(&H3081) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3067) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H3002)
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0179] modEntryFormat.Btn_DisableFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    If sections.Count = 0 Then
        WarnUser ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H7121) & ChrW(&H52B9) & ChrW(&H5316), _
                 ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H0020) & fmtId & ChrW(&H0020) & ChrW(&H306E) & ChrW(&H5B9A) & ChrW(&H7FA9) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H304C) & ChrW(&H3042) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002)
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0180] modEntryFormat.Btn_DisableFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ' 3. Find the [FORMAT] section and set/overwrite Status="????".
    '    If the [FORMAT] section is missing for any reason, fall
    '    through to a warning rather than silently inserting one
    '    (the file is malformed in that case).
    Dim sec As ClsStanzaSection
    Dim fmtSec As ClsStanzaSection
    Dim i As Long
    For i = 1 To sections.Count
        Set sec = sections.Item(i)
        If sec.SectionName = "FORMAT" Then
            Set fmtSec = sec
            Exit For
        End If
    Next i
    If fmtSec Is Nothing Then
        WarnUser ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H7121) & ChrW(&H52B9) & ChrW(&H5316), _
                 ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H0020) & fmtId & ChrW(&H0020) & ChrW(&H306B) & ChrW(&H0020) & ChrW(&H005B) & ChrW(&H0046) & ChrW(&H004F) & ChrW(&H0052) & ChrW(&H004D) & ChrW(&H0041) & ChrW(&H0054) & ChrW(&H005D) & ChrW(&H0020) & ChrW(&H30BB) & ChrW(&H30AF) & ChrW(&H30B7) & ChrW(&H30E7) & ChrW(&H30F3) & ChrW(&H304C) & ChrW(&H3042) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002)
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0181] modEntryFormat.Btn_DisableFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    fmtSec.SetValue FMT_KEY_STATUS, FMT_STATUS_DISABLED

    ' 4. Persist back to disk via modFormatLoader.SaveFormat.
    Dim rc As Long
    rc = modFormatLoader.SaveFormat(fmtId, sections)
    If rc <> 0 Then
        WarnUser ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H7121) & ChrW(&H52B9) & ChrW(&H5316), _
                 ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H0020) & fmtId & ChrW(&H0020) & ChrW(&H306E) & ChrW(&H7121) & ChrW(&H52B9) & ChrW(&H5316) & ChrW(&H30D5) & ChrW(&H30E9) & ChrW(&H30B0) & ChrW(&H3092) & ChrW(&H66F8) & ChrW(&H304D) & ChrW(&H8FBC) & ChrW(&H3081) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3067) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H0020) & ChrW(&H0028) & ChrW(&H0072) & ChrW(&H0063) & ChrW(&H003D) & rc & ChrW(&H0029) & ChrW(&H3002)
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0182] modEntryFormat.Btn_DisableFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ' 5. Log the disable action (note: this is NOT delete).
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = NewLogger()
    If Not lg Is Nothing Then
        lg.LogInfo "modEntryFormat", "Btn_DisableFormat", _
                   "Format disabled (logical, not deleted): " & fmtId, _
                   fmtId, "M02-DISABLE-OK-II-041"
    End If
    On Error GoTo ErrHandler

    ' 6. Refresh the M-02 grid so column L shows "????".
    Btn_ReloadFormats
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0183] modEntryFormat.Btn_DisableFormat EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0184] modEntryFormat.Btn_DisableFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    On Error Resume Next
    Dim lgErr As clsLogger
    Set lgErr = NewLogger()
    If Not lgErr Is Nothing Then
        lgErr.LogError "modEntryFormat", "Btn_DisableFormat", _
                       "Disable failed: " & Err.Description, _
                       fmtId, "M02-DISABLE-ERR-EE-041"
    End If
    On Error GoTo 0
    Debug.Print "[ERR] Btn_DisableFormat: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_DeleteFormat
' Role: M-02 "??" - physically removes <fmtId>.txt via
'       clsFormatManager.DeleteFormat. The manager rejects the
'       delete (rc=2) if any knowledge entries still reference
'       the format - that's a hard guard against orphaning data.
'       Row selection: single checked row on M-02 (was ActiveCell
'       in v2.2; v2.3 uses checkbox row select).
' ============================================================
Public Sub Btn_DeleteFormat()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0185] modEntryFormat.Btn_DeleteFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_DeleteFormat"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config;format_dir", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0186] modEntryFormat.Btn_DeleteFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim fmtId As String
    fmtId = PickM02CheckedRowFmtId()
    If Len(fmtId) = 0 Then Exit Sub
    If Not IsHeadless() Then
        If MsgBox(ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & " " & fmtId & " " & ChrW(&H3092) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H304B) & ChrW(&HFF1F), vbYesNo + vbExclamation, _
                  ChrW(&H524A) & ChrW(&H9664) & ChrW(&H78BA) & ChrW(&H8A8D)) <> vbYes Then Exit Sub
    End If
    Dim mgr As clsFormatManager
    Set mgr = New clsFormatManager
    mgr.Init NewLogger()
    Dim rc As Long
    rc = mgr.DeleteFormat(fmtId)
    ShowDeleteFormatResult fmtId, rc
    If rc = 0 Then
        ' Refresh the grid after a successful delete.
        Btn_ReloadFormats
    End If
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0187] modEntryFormat.Btn_DeleteFormat EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0188] modEntryFormat.Btn_DeleteFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_DeleteFormat: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_ReloadFormats  (v2.3 NEW)
' Role: M-02 "?????[?h" - rescan <format_dir>\*.txt and rewrite
'       the M-02 grid data area. Header row is not touched.
'
'       For each format file:
'         - LoadFormat() to get its [FORMAT] + [FIELD] sections.
'         - Extract meta:
'             FormatID      = file basename (canonical)
'             FormatName    = [FORMAT].FormatName
'             FormatVersion = [FORMAT].FormatVersion
'             FieldCount    = count of [FIELD] sections
'             UpdatedAt     = filesystem mtime of <fmtId>.txt
'             Status        = [FORMAT].Status (blank means active)
'             KnowledgeCount= ListKnowledgesByFormat(fmtId).Count
'
'       Row selection state (column B checkbox) is cleared as
'       part of the reload - the user re-checks after refresh.
' ============================================================
Public Sub Btn_ReloadFormats()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0189] modEntryFormat.Btn_ReloadFormats ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_ReloadFormats"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config;format_dir", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0190] modEntryFormat.Btn_ReloadFormats EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M02_SHEET_NAME())

    ' 2026-06-06: explicit Unprotect before writes (sheet protection blocks VBA writes after reload)
    Dim wasProtected As Boolean: wasProtected = ws.ProtectContents
    If wasProtected Then ws.Unprotect

    ' 1. Clear the existing data rows (B..L on every data row).
    ClearM02GridDataRows ws

    ' 2026-06-06: unlock data cells so users can write checkbox marks in column B
    ws.Range(M02_COL_ROWCHECK & M02_DATA_FIRST_ROW & ":" & M02_COL_ROWCHECK & M02_DATA_LAST_ROW).Locked = False

    ' 2. Enumerate all format files.
    Dim allIds As Collection
    Set allIds = modFormatLoader.ListAllFormats()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fmtDir As String
    fmtDir = modConfigHolder.GetFormatDir()

    Dim writeRow As Long
    writeRow = M02_DATA_FIRST_ROW

    Dim idVar As Variant
    Dim writtenRows As Long
    writtenRows = 0
    For Each idVar In allIds
        If writeRow > M02_DATA_LAST_ROW Then Exit For
        Dim fmtId As String
        fmtId = CStr(idVar)
        WriteM02RowFor ws, writeRow, fmtId, fmtDir, fso
        writeRow = writeRow + 1
        writtenRows = writtenRows + 1
    Next idVar

    ' 2026-06-06: re-apply Protect after writes
    If wasProtected Then ws.Protect Password:="", UserInterfaceOnly:=True

    ' 3. Log the reload count.
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = NewLogger()
    If Not lg Is Nothing Then
        lg.LogInfo "modEntryFormat", "Btn_ReloadFormats", _
                   "M-02 grid reloaded; rows=" & writtenRows, _
                   "", "M02-RELOAD-OK-II-042"
    End If
    On Error GoTo 0
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0191] modEntryFormat.Btn_ReloadFormats EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0192] modEntryFormat.Btn_ReloadFormats EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    On Error Resume Next
    Dim lgErr As clsLogger
    Set lgErr = NewLogger()
    If Not lgErr Is Nothing Then
        lgErr.LogError "modEntryFormat", "Btn_ReloadFormats", _
                       "Reload failed: " & Err.Description, _
                       "", "M02-RELOAD-ERR-EE-042"
    End If
    On Error GoTo 0
    Debug.Print "[ERR] Btn_ReloadFormats: " & Err.Number & " " & Err.Description
End Sub


' ============================================================
' ===== M-02 grid private helpers ============================
' ============================================================

' ============================================================
' Private Sub: ClearM02GridDataRows
' Role: zap columns B..L on every data row of the M-02 grid.
'       Single Range.ClearContents covering the whole block is
'       cheaper than the per-row loop but keeps the same shape
'       used elsewhere in the module for readability.
' ============================================================
Private Sub ClearM02GridDataRows(ByVal ws As Worksheet)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0193] modEntryFormat.ClearM02GridDataRows ENTER"  ' [ADR-0100]
    Dim addr As String
    addr = M02_COL_ROWCHECK & M02_DATA_FIRST_ROW & ":" & _
           M02_COL_UPDATED & M02_DATA_LAST_ROW
    On Error Resume Next
    ws.Range(addr).ClearContents
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0194] modEntryFormat.ClearM02GridDataRows EXIT-OK"  ' [ADR-0100]
End Sub

' ============================================================
' Private Sub: WriteM02RowFor
' Role: populate a single M-02 grid row for the given fmtId.
'       Reads the format definition, derives the meta values,
'       and writes them across columns B..L.
'       Column B (row-check) is intentionally left blank - the
'       user picks rows after reload.
' ============================================================
Private Sub WriteM02RowFor(ByVal ws As Worksheet, _
                           ByVal rowNum As Long, _
                           ByVal fmtId As String, _
                           ByVal fmtDir As String, _
                           ByVal fso As Object)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0195] modEntryFormat.WriteM02RowFor ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    ' Load the sections (may be empty if file missing/unparsable).
    Dim sections As Collection
    Set sections = modFormatLoader.LoadFormat(fmtId)

    Dim fmtName As String
    Dim fmtVer  As String
    Dim fmtStat As String
    Dim fmtIdPat As String
    Dim fmtNxtNum As String
    Dim fieldCount As Long
    fieldCount = 0

    If Not sections Is Nothing Then
        Dim sec As ClsStanzaSection
        Dim i As Long
        For i = 1 To sections.Count
            Set sec = sections.Item(i)
            Select Case sec.SectionName
                Case "FORMAT"
                    If sec.HasKey("FormatName") Then fmtName = sec.GetValue("FormatName")
                    If sec.HasKey("FormatVersion") Then fmtVer = sec.GetValue("FormatVersion")
                    If sec.HasKey(FMT_KEY_STATUS) Then fmtStat = sec.GetValue(FMT_KEY_STATUS)
                    If sec.HasKey("IDPattern") Then fmtIdPat = sec.GetValue("IDPattern")
                    If sec.HasKey("NextNumber") Then fmtNxtNum = sec.GetValue("NextNumber")
                Case "FIELD"
                    fieldCount = fieldCount + 1
            End Select
        Next i
    End If

    ' Knowledge count - via modKnowledgeFileIO (defensive: any
    ' exception falls through to zero).
    Dim knwCount As Long
    knwCount = 0
    On Error Resume Next
    Dim knwList As Collection
    Set knwList = modKnowledgeFileIO.ListKnowledgesByFormat(fmtId)
    If Not knwList Is Nothing Then knwCount = knwList.Count
    On Error GoTo ErrHandler

    ' Updated-at: filesystem mtime of <fmtId>.txt.
    Dim updatedAt As String
    updatedAt = ""
    On Error Resume Next
    Dim fpath As String
    fpath = fmtDir & fmtId & ".txt"
    If fso.FileExists(fpath) Then
        updatedAt = Format(fso.GetFile(fpath).DateLastModified, "yyyy-mm-dd hh:nn")
    End If
    On Error GoTo ErrHandler

    ' Sequence number - 1-based index within the visible block.
    ws.Range(M02_COL_NO & rowNum).Value = rowNum - M02_DATA_FIRST_ROW + 1
    ws.Range(M02_COL_FMT_ID & rowNum).Value = fmtId
    ws.Range(M02_COL_FMT_NAME & rowNum).Value = fmtName
    ws.Range(M02_COL_FIELD_COUNT & rowNum).Value = fieldCount
    ws.Range(M02_COL_KNW_COUNT & rowNum).Value = knwCount
    ws.Range(M02_COL_UPDATED & rowNum).Value = updatedAt
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0196] modEntryFormat.WriteM02RowFor EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0197] modEntryFormat.WriteM02RowFor EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' Don't let a single broken format kill the whole reload.
    Debug.Print "[WriteM02RowFor] " & fmtId & " row=" & rowNum & _
                " err=" & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Private Function: PickM02CheckedRowFmtId
' Role: scan column B (M02_COL_ROWCHECK) across the data rows
'       and return the fmtId of the single checked row.
'       Returns "" if zero or more than one rows are checked
'       (the caller is expected to exit silently in that case;
'       the user-facing warning is issued from here).
' ============================================================
Private Function PickM02CheckedRowFmtId() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0198] modEntryFormat.PickM02CheckedRowFmtId ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M02_SHEET_NAME())

    Dim checkedCount As Long
    Dim hitRow As Long
    checkedCount = 0
    hitRow = 0

    Dim r As Long
    For r = M02_DATA_FIRST_ROW To M02_DATA_LAST_ROW
        If IsRowChecked(ws.Range(M02_COL_ROWCHECK & r).Value) Then
            checkedCount = checkedCount + 1
            hitRow = r
            If checkedCount > 1 Then Exit For
        End If
    Next r

    If checkedCount = 0 Then
        WarnUser ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H4E00) & ChrW(&H89A7), _
                 ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H884C) & ChrW(&H3092) & ChrW(&H0031) & ChrW(&H3064) & ChrW(&H3060) & ChrW(&H3051) & ChrW(&H9078) & ChrW(&H629E) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
        PickM02CheckedRowFmtId = ""
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0199] modEntryFormat.PickM02CheckedRowFmtId EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If checkedCount > 1 Then
        WarnUser ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H4E00) & ChrW(&H89A7), _
                 ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C) & ChrW(&H304C) & ChrW(&H9078) & ChrW(&H629E) & ChrW(&H3055) & ChrW(&H308C) & ChrW(&H3066) & ChrW(&H3044) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H3002) & ChrW(&H0031) & ChrW(&H3064) & ChrW(&H3060) & ChrW(&H3051) & ChrW(&H9078) & ChrW(&H629E) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
        PickM02CheckedRowFmtId = ""
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0200] modEntryFormat.PickM02CheckedRowFmtId EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim fmtId As String
    fmtId = Trim(CStr(ws.Range(M02_COL_FMT_ID & hitRow).Value))
    If Len(fmtId) = 0 Then
        WarnUser ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H4E00) & ChrW(&H89A7), _
                 ChrW(&H9078) & ChrW(&H629E) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H884C) & ChrW(&H306B) & ChrW(&H0020) & ChrW(&H0046) & ChrW(&H006F) & ChrW(&H0072) & ChrW(&H006D) & ChrW(&H0061) & ChrW(&H0074) & ChrW(&H0049) & ChrW(&H0044) & ChrW(&H0020) & ChrW(&H304C) & ChrW(&H3042) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002)
        PickM02CheckedRowFmtId = ""
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0201] modEntryFormat.PickM02CheckedRowFmtId EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    PickM02CheckedRowFmtId = fmtId
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0202] modEntryFormat.PickM02CheckedRowFmtId EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0203] modEntryFormat.PickM02CheckedRowFmtId EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    PickM02CheckedRowFmtId = ""
End Function

' ============================================================
' Private Function: IsRowChecked
' Role: normalise the column-B value into a boolean. Accepts
'       TRUE / True / "TRUE" / "1" / "yes" / "Y" / "x" / "X" /
'       checkmark glyphs. Anything else (including empty) is
'       treated as unchecked.
' ============================================================
Private Function IsRowChecked(ByVal v As Variant) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0204] modEntryFormat.IsRowChecked ENTER"  ' [ADR-0100]
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        IsRowChecked = CBool(v)
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0205] modEntryFormat.IsRowChecked EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim s As String
    s = LCase(Trim(CStr(v)))
    Select Case s
        Case "", "false", "0", "no", "n"
            IsRowChecked = False
        Case "true", "1", "yes", "y", "x", ChrW(&H2713), ChrW(&H2714), ChrW(&H25A0)
            IsRowChecked = True
        Case Else
            ' Be lenient: any non-empty, non-"false" value counts
            ' as checked. This lets users put a single-character
            ' marker (e.g. "o", "*") in column B without surprise.
            IsRowChecked = (Len(s) > 0)
    End Select
End Function


' ============================================================
' ===== Auto-numbering for Btn_NewFormatDraft ================
' ============================================================

' ============================================================
' Private Function: AutoNumberNextFmtId
' Role: find the lowest "FMT-NNN" not already present in
'       <format_dir>. NNN is zero-padded to FMT_ID_WIDTH. If
'       FMT_ID_SCAN_LIMIT is exhausted, fall back to FMT-001
'       (the user will overwrite anyway).
' ============================================================
Private Function AutoNumberNextFmtId() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0206] modEntryFormat.AutoNumberNextFmtId ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim used As Object
    Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = 1   ' TextCompare (case-insensitive)

    Dim allIds As Collection
    Set allIds = modFormatLoader.ListAllFormats()
    Dim idVar As Variant
    For Each idVar In allIds
        used(CStr(idVar)) = True
    Next idVar

    Dim i As Long
    For i = 1 To FMT_ID_SCAN_LIMIT
        Dim candidate As String
        candidate = FMT_ID_PREFIX & PadLeft(CStr(i), FMT_ID_WIDTH, "0")
        If Not used.Exists(candidate) Then
            AutoNumberNextFmtId = candidate
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0207] modEntryFormat.AutoNumberNextFmtId EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next i
    AutoNumberNextFmtId = FMT_ID_PREFIX & PadLeft("1", FMT_ID_WIDTH, "0")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0208] modEntryFormat.AutoNumberNextFmtId EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0209] modEntryFormat.AutoNumberNextFmtId EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    AutoNumberNextFmtId = FMT_ID_PREFIX & PadLeft("1", FMT_ID_WIDTH, "0")
End Function

' ============================================================
' Public Sub: SeedM03FormatIdIfEmpty (2026-06-07)
' Role: clsFormatDesignScreen.Setup hook. If the FormatID cell on M-03
'       (typically C4 for the new-draft side, resolved via
'       ResolveM03FmtIdCell -> ui_seed) is empty, seed it with the next
'       available auto-numbered FmtId (FMT-NNN). This lets a freshly
'       installed workbook show "FMT-001" (or next free) on first open
'       without requiring the user to press the New Draft button.
' Behavior: does NOT overwrite a non-empty existing value, so the seed
'       only fires the first time the screen has no FormatID.
' ============================================================
Public Sub SeedM03FormatIdIfEmpty()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0260] modEntryFormat.SeedM03FormatIdIfEmpty ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M03_SHEET_NAME())
    Dim cell As String
    cell = ResolveM03FmtIdCell()
    If Len(cell) = 0 Then cell = M03_CELL_FMT_ID_FALLBACK
    Dim cur As String
    cur = Trim$(CStr(ws.Range(cell).Value))
    If Len(cur) > 0 Then
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0261] modEntryFormat.SeedM03FormatIdIfEmpty EXIT-OK (already seeded)"  ' [ADR-0100]
        Exit Sub
    End If
    Dim nextId As String
    nextId = AutoNumberNextFmtId()
    If Len(nextId) = 0 Then nextId = FMT_ID_PREFIX & PadLeft("1", FMT_ID_WIDTH, "0")
    ' Unprotect-write-reprotect so the seed lands even when M-03 is light-protected.
    On Error Resume Next
    ws.Unprotect
    On Error GoTo ErrHandler
    ws.Range(cell).Value = nextId
    On Error Resume Next
    ws.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True
    On Error GoTo ErrHandler
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0262] modEntryFormat.SeedM03FormatIdIfEmpty EXIT-OK seeded=" & nextId  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0263] modEntryFormat.SeedM03FormatIdIfEmpty EXIT-ERR err=" & Err.Number & " " & Err.Description  ' [ADR-0100]
End Sub

' ============================================================
' Private Function: PadLeft
' Role: zero-pad a string on the left to the given width. VBA
'       lacks a built-in, so we keep a tiny local helper rather
'       than reaching for Format(... , "000") which would couple
'       us to a per-width string template.
' ============================================================
Private Function PadLeft(ByVal s As String, ByVal width As Long, ByVal pad As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0210] modEntryFormat.PadLeft ENTER"  ' [ADR-0100]
    Dim out As String
    out = s
    Do While Len(out) < width
        out = pad & out
    Loop
    PadLeft = out
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0211] modEntryFormat.PadLeft EXIT-OK"  ' [ADR-0100]
End Function


' ============================================================
' ===== Shared helpers (preserved from v2.2) =================
' ============================================================

' ============================================================
' Public Sub: Btn_BackToMain_v21
' Role: M-03 navigation back to the main sheet. Prompts when
'       there is unsaved input on the M-03 design cells.
'       Preserved verbatim from v2.2 - the M-03 design grid is
'       still in the same place.
' ============================================================
Public Sub Btn_BackToMain_v21()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0212] modEntryFormat.Btn_BackToMain_v21 ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_BackToMain_v21"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0213] modEntryFormat.Btn_BackToMain_v21 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim hasInput As Boolean
    hasInput = HasUnsavedInput_M03(ActiveSheet)
    If hasInput Then
        On Error Resume Next
        Dim oLogger031v As clsLogger
        Set oLogger031v = New clsLogger
        oLogger031v.Init ThisWorkbook.Worksheets("LOG")
        oLogger031v.LogWarn "modEntryFormat", "Btn_BackToMain_v21", "M-03 unsaved input detected, confirm dialog shown", "", "VALIDATE-WARN-WW-031"
        On Error GoTo ErrHandler
        ' Headless guard: under COM automation skip the modal dialog and
        ' fall through to the default (discard and navigate to main).
        Dim r As VbMsgBoxResult
        If IsHeadless() Then
            r = vbNo
        Else
            r = MsgBox(ChrW(&H672A) & ChrW(&H4FDD) & ChrW(&H5B58) & ChrW(&H306E) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H304C) & ChrW(&H3042) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H3002) & ChrW(&H623B) & ChrW(&H308B) & ChrW(&H524D) & ChrW(&H306B) & ChrW(&H4FDD) & ChrW(&H5B58) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H304B) & ChrW(&HFF1F) & vbCrLf & _
                       "Yes=Save and back / No=Discard and back / Cancel=Stay editing", _
                       vbYesNoCancel + vbExclamation, "Confirm")
        End If
        Select Case r
            Case vbYes
                Btn_SaveFormat
                NavigateToMain_v21
            Case vbNo
                NavigateToMain_v21
            Case vbCancel
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0214] modEntryFormat.Btn_BackToMain_v21 EXIT-OK"  ' [ADR-0100]
                Exit Sub
        End Select
    Else
        NavigateToMain_v21
    End If
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0215] modEntryFormat.Btn_BackToMain_v21 EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0216] modEntryFormat.Btn_BackToMain_v21 EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    On Error Resume Next
    Dim oLogger032b As clsLogger
    Set oLogger032b = New clsLogger
    oLogger032b.Init ThisWorkbook.Worksheets("LOG")
    oLogger032b.LogError "modEntryFormat", "Btn_BackToMain_v21", "Back to main transition error: " & Err.Description, "", "BACKTOMAIN-ERR-EE-032"
    On Error GoTo 0
    If Not IsHeadless() Then
        MsgBox ChrW(&H30E1) & ChrW(&H30A4) & ChrW(&H30F3) & ChrW(&H753B) & ChrW(&H9762) & ChrW(&H623B) & ChrW(&H308A) & ChrW(&H51E6) & ChrW(&H7406) & ChrW(&H3067) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC) & ChrW(&H003A) & ChrW(&H0020) & Err.Description, vbCritical, ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
    Else
        Debug.Print "[ERR] Btn_BackToMain_v21: " & Err.Number & " " & Err.Description
    End If
End Sub

Private Function HasUnsavedInput_M03(ByVal ws As Object) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0217] modEntryFormat.HasUnsavedInput_M03 ENTER"  ' [ADR-0100]
    ' ADR-0090: Primary path enumerates [INPUT] cells from ui_seed M-03 SSOT.
    ' Legacy hard-code cells (C3 / B7) are unioned as fallback so the
    ' detection coverage does not shrink even if a stanza key is added/removed.
    On Error GoTo ErrHandler
    Dim ui As Collection
    Set ui = modUILoader.LoadUiDefinition(RoleKanri(), "M-03")
    Dim i As Long
    If Not ui Is Nothing Then
        For i = 1 To ui.Count
            Dim sec As ClsStanzaSection
            Set sec = ui.Item(i)
            If sec.SectionName = "INPUT" Then
                Dim cellExpr As String
                cellExpr = sec.GetValue("Cell")
                If Len(cellExpr) > 0 Then
                    Dim firstCell As String
                    firstCell = FirstCellOf_M03(cellExpr)
                    If Len(CStr(ws.Range(firstCell).Value)) > 0 Then
                        HasUnsavedInput_M03 = True
                        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0218] modEntryFormat.HasUnsavedInput_M03 EXIT-OK"  ' [ADR-0100]
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
    ' Legacy fallback cells (SSOT-driven coverage may not include these)
    ' iter18 (2026-06-01, ADR-0090 Phase 2): the FormatID-cell address
    ' is now routed through ResolveM03FmtIdCell() rather than the bare
    ' "C3" literal, so any future ui_seed amendment that adds an
    ' explicit FormatID INPUT lands here too without further edits.
    If Len(CStr(ws.Range(ResolveM03FmtIdCell()).Value)) > 0 Then
        HasUnsavedInput_M03 = True
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0219] modEntryFormat.HasUnsavedInput_M03 EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If Len(CStr(ws.Range("B7").Value)) > 0 Then
        HasUnsavedInput_M03 = True
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0220] modEntryFormat.HasUnsavedInput_M03 EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    HasUnsavedInput_M03 = False
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0221] modEntryFormat.HasUnsavedInput_M03 EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0222] modEntryFormat.HasUnsavedInput_M03 EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    HasUnsavedInput_M03 = False
End Function

Private Function FirstCellOf_M03(ByVal cellExpr As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0223] modEntryFormat.FirstCellOf_M03 ENTER"  ' [ADR-0100]
    ' Returns first cell of "Cell=A1:B1" form, e.g. "A1"
    Dim p As Long
    p = InStr(1, cellExpr, ":")
    If p > 0 Then
        FirstCellOf_M03 = Left(cellExpr, p - 1)
    Else
        FirstCellOf_M03 = cellExpr
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0224] modEntryFormat.FirstCellOf_M03 EXIT-OK"  ' [ADR-0100]
End Function

' ============================================================
' Public Function: ResolveM03FmtIdCell  (ADR-0090 Phase 2, 2026-06-01 iter18)
' Role: SSOT-driven resolution of the M-03 FormatID-input cell.
'       Reads ui_seed M-03 [INPUT] stanzas and returns the Cell
'       address of the entry whose inputDataKey is "FormatID" or
'       "targetFormat". When no such stanza is present (or the
'       lookup fails) falls back to the legacy hard-code "C3" so
'       that existing installs keep working untouched.
'
'       Both production write paths (Btn_NewFormatDraft,
'       Btn_PreviewFormat) and the E2E harness (modE2E_AllButtons)
'       go through this single helper so the cell-address constant
'       is no longer duplicated as a hard-code literal across
'       multiple modules (ADR-0090 compliance, hard-code free).
'
'       See ADR-0093 (2026-06-01) for the separate ui_seed-vs-
'       production cell-address drift (ui_seed currently exposes
'       C4:D4 / C8:D8 for "targetFormat", while production code has
'       historically written/read C3). The fallback path documents
'       that drift; resolving it requires syncing ui_seed and the
'       admin workbook seed step which is out of scope here.
' ============================================================
Public Function ResolveM03FmtIdCell() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0225] modEntryFormat.ResolveM03FmtIdCell ENTER"  ' [ADR-0100]
    On Error GoTo Fallback
    Dim ui As Collection
    Set ui = modUILoader.LoadUiDefinition(RoleKanri(), "M-03")
    If ui Is Nothing Then GoTo Fallback
    If ui.Count = 0 Then GoTo Fallback
    Dim i As Long
    ' Only accept the explicit "FormatID" / "formatId" / "fmtId" SSOT
    ' keys here. The current ui_seed M-03.txt exposes "targetFormat"
    ' (C4:D4 and C8:D8) which production code never read - bringing
    ' production over to that cell would silently change the M-03
    ' write target on existing installs. ADR-0093 (2026-06-01) tracks
    ' that drift as a deliberate, separately-scoped follow-up; until
    ' the ui_seed is amended to expose an explicit FormatID INPUT
    ' the fallback ("C3") preserves existing behaviour.
    For i = 1 To ui.Count
        Dim sec As ClsStanzaSection
        Set sec = ui.Item(i)
        If sec.SectionName = "INPUT" Then
            Dim key As String
            key = sec.GetValue("inputDataKey")
            ' iter18b ADR-0090: accept the v2.3 ui_seed key set ("targetFormat")
            ' as well as the explicit FormatID variants. The production write paths
            ' (Btn_NewFormatDraft / Btn_PreviewFormat) and the E2E harness all go
            ' through this helper so both write to / read from the same cell.
            If key = "FormatID" Or key = "formatId" Or key = "fmtId" Or key = "targetFormat" Then
                Dim cellExpr As String
                cellExpr = sec.GetValue("Cell")
                If Len(cellExpr) > 0 Then
                    ResolveM03FmtIdCell = FirstCellOf_M03(cellExpr)
                    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0226] modEntryFormat.ResolveM03FmtIdCell EXIT-OK"  ' [ADR-0100]
                    Exit Function
                End If
            End If
        End If
    Next i
Fallback:
    ResolveM03FmtIdCell = M03_CELL_FMT_ID_FALLBACK
End Function

Private Sub NavigateToMain_v21()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0227] modEntryFormat.NavigateToMain_v21 ENTER"  ' [ADR-0100]
    ' v2.3 (2026-05-26): M-01 / Main ??p?~ (ADR-0053 ??2.1)?B
    ' v2.1 legacy path ????? On Error Resume Next ??S???????????v?B
    ' R-2-Fix2b (2026-05-28): M-01 / Main menu removed (ADR-0053 2.1); the old
    ' hard Worksheets("Main"/"M-01").Activate was dead code. Existence-checked
    ' silent skip; keeps a recovery path if a menu sheet is reintroduced.
    On Error Resume Next
    Dim wsMenu As Worksheet
    Set wsMenu = Nothing
    Set wsMenu = ThisWorkbook.Worksheets("Main")
    If wsMenu Is Nothing Then Set wsMenu = ThisWorkbook.Worksheets("M-01")
    If Not wsMenu Is Nothing Then wsMenu.Activate
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0228] modEntryFormat.NavigateToMain_v21 EXIT-OK"  ' [ADR-0100]
End Sub

' ============================================================
' Public Sub: Btn_ConfirmDiff
' Role: M-12 migration confirmation (legacy v2.1/v2.2 entry).
'       Kept here for compile-time compatibility; M-12 in v2.3
'       is reduced to a read-only check that lives in
'       modEntrySettings.Btn_CheckFormat. This handler stays
'       wired in case the legacy migrate UI is still seeded.
' ============================================================
Public Sub Btn_ConfirmDiff()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0229] modEntryFormat.Btn_ConfirmDiff ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_ConfirmDiff"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0230] modEntryFormat.Btn_ConfirmDiff EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ws As Object: Set ws = ActiveSheet
    Dim ui As Collection: Set ui = modUILoader.LoadUiDefinition(RoleKanri(), "M-03")
    If ui.Count = 0 Then Exit Sub

    Dim tempDict As Object: Set tempDict = BuildFormatDictFromCells(ws, ui)
    Dim fmtId As String
    If tempDict.Exists("FormatID") Then fmtId = CStr(tempDict("FormatID"))
    If Len(fmtId) = 0 Then Exit Sub

    Dim newVer As Long
    newVer = 2
    If tempDict.Exists("FormatVersion") Then
        Dim verStr As String
        verStr = CStr(tempDict("FormatVersion"))
        If Len(verStr) > 0 And IsNumeric(verStr) Then newVer = CLng(verStr)
    End If

    Dim affected As Collection
    Set affected = modKnowledgeFileIO.ListKnowledgesByFormat(fmtId)
    Dim affCount As Long
    If Not affected Is Nothing Then affCount = affected.Count

    Dim msg As String
    msg = "FormatID=" & fmtId & vbCrLf
    msg = msg & "NewVersion=" & newVer & vbCrLf
    msg = msg & "Affected knowledge count: " & affCount & vbCrLf
    msg = msg & "Migrate now?" & vbCrLf
    msg = msg & "(Yes=migrate / No=cancel)"

    Dim r As VbMsgBoxResult
    If IsHeadless() Then
        r = vbYes
    Else
        r = MsgBox(msg, vbYesNo + vbQuestion, "Confirm format diff")
    End If

    If r = vbYes Then
        ' clsFieldMigrator is archived out of v2.3 - this branch
        ' is left as a documented no-op so legacy wired buttons
        ' that still call Btn_ConfirmDiff do not crash the host.
        If Not IsHeadless() Then
            MsgBox ChrW(&H0076) & ChrW(&H0032) & ChrW(&H002E) & ChrW(&H0033) & ChrW(&H0020) & ChrW(&H3067) & ChrW(&H306F) & ChrW(&H79FB) & ChrW(&H884C) & ChrW(&H6A5F) & ChrW(&H80FD) & ChrW(&H306F) & ChrW(&H5229) & ChrW(&H7528) & ChrW(&H3067) & ChrW(&H304D) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H0020) & ChrW(&H0028) & ChrW(&H004D) & ChrW(&H002D) & ChrW(&H0031) & ChrW(&H0032) & ChrW(&H0020) & ChrW(&H81EA) & ChrW(&H52D5) & ChrW(&H53CD) & ChrW(&H6620) & ChrW(&H306F) & ChrW(&H5EC3) & ChrW(&H6B62) & ChrW(&H3055) & ChrW(&H308C) & ChrW(&H307E) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H0029) & ChrW(&H3002), _
                   vbInformation, ChrW(&H7D50) & ChrW(&H679C)
        Else
            Debug.Print "[Btn_ConfirmDiff] migrate path is a no-op under v2.3"
        End If
    End If
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0231] modEntryFormat.Btn_ConfirmDiff EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0232] modEntryFormat.Btn_ConfirmDiff EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    If Not IsHeadless() Then
        MsgBox ChrW(&H5DEE) & ChrW(&H5206) & ChrW(&H78BA) & ChrW(&H5B9A) & ChrW(&H51E6) & ChrW(&H7406) & ChrW(&H3067) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC) & ChrW(&H003A) & ChrW(&H0020) & Err.Description, vbCritical, ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
    Else
        Debug.Print "[ERR] Btn_ConfirmDiff: " & Err.Number & " " & Err.Description
    End If
End Sub

' ------------------------------------------------------------
' IsHeadless: TRUE when running under COM automation with no
' interactive UI. Used to suppress modal dialogs (MsgBox /
' InputBox) that would otherwise block a headless harness.
' Checks two independent signals because Application.Interactive
' alone defaults to TRUE under COM automation:
'   - Application.Interactive = False -> harness explicitly off
'   - Application.Visible     = False -> Excel window is hidden
' If EITHER indicates headless, treat the run as headless.
' ------------------------------------------------------------
Private Function IsHeadless() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0233] modEntryFormat.IsHeadless ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim notInteractive As Boolean
    Dim notVisible As Boolean
    notInteractive = (Application.Interactive = False)
    notVisible = (Application.Visible = False)
    IsHeadless = (notInteractive Or notVisible)
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0234] modEntryFormat.IsHeadless EXIT-OK"  ' [ADR-0100]
End Function

' ============================================================
' Private Sub: RenderM04Preview
' Role: rerender the M-04 preview grid. Creates a fresh
'       clsFormatPreviewScreen, inits with a minimal spec
'       (ScreenId/SheetName=M-04, no renderer), and calls
'       RenderPreview.
' ============================================================
Private Sub RenderM04Preview()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0235] modEntryFormat.RenderM04Preview ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim spec As clsScreenSpec
    Set spec = New clsScreenSpec
    spec.ScreenId = "M-04"
    spec.SheetName = "M-04"
    Dim prev As clsFormatPreviewScreen
    Set prev = New clsFormatPreviewScreen
    prev.Init Nothing, spec
    prev.RenderPreview
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0236] modEntryFormat.RenderM04Preview EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0237] modEntryFormat.RenderM04Preview EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] RenderM04Preview: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Private Sub: ShowDeleteFormatResult
' Role: pretty-print clsFormatManager.DeleteFormat rc codes.
'       0=success / 1=outside kanri.xlsm reject /
'       2=knowledge-still-using reject / else=failure.
' ============================================================
Private Sub ShowDeleteFormatResult(ByVal fmtId As String, ByVal rc As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0238] modEntryFormat.ShowDeleteFormatResult ENTER"  ' [ADR-0100]
    If IsHeadless() Then Exit Sub
    Dim msg As String
    Select Case rc
        Case 0
            ' format deleted JP
            msg = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & " " & fmtId & " " & ChrW(&H3092) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H3002)
        Case 1
            ' delete from admin only JP
            msg = ChrW(&H7BA1) & ChrW(&H7406) & ".xlsm " & ChrW(&H304B) & ChrW(&H3089) & ChrW(&H3057) & ChrW(&H304B) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H3067) & ChrW(&H304D) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002)
        Case 2
            ' knowledge still uses format JP
            msg = ChrW(&H3053) & ChrW(&H306E) & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & " " & fmtId & " " & ChrW(&H3092) & ChrW(&H4F7F) & ChrW(&H3046) & ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H304C) & ChrW(&H6B8B) & ChrW(&H3063) & ChrW(&H3066) & ChrW(&H3044) & ChrW(&H308B) & ChrW(&H305F) & ChrW(&H3081) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H3067) & ChrW(&H304D) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002)
        Case Else
            ' delete failed JP
            msg = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & " " & fmtId & " " & ChrW(&H306E) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H306B) & ChrW(&H5931) & ChrW(&H6557) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H3002)
    End Select
    MsgBox msg, vbInformation, ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H524A) & ChrW(&H9664)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0239] modEntryFormat.ShowDeleteFormatResult EXIT-OK"  ' [ADR-0100]
End Sub

' ============================================================
' Private Function: NewLogger
' Role: create a clsLogger bound to the LOG sheet. Nothing on
'       failure so callers can fall back to Debug.Print.
' ============================================================
Private Function NewLogger() As clsLogger
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0240] modEntryFormat.NewLogger ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ThisWorkbook.Worksheets("LOG")
    Set NewLogger = lg
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0241] modEntryFormat.NewLogger EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0242] modEntryFormat.NewLogger EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set NewLogger = Nothing
End Function

' ============================================================
' Private Sub: WarnUser
' Role: MsgBox-with-headless-guard, so the warning is silent
'       under COM automation but visible to a real Excel user.
' ============================================================
Private Sub WarnUser(ByVal title As String, ByVal body As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0243] modEntryFormat.WarnUser ENTER"  ' [ADR-0100]
    If IsHeadless() Then
        Debug.Print "[WARN " & title & "] " & body
        Exit Sub
    End If
    MsgBox body, vbExclamation, title
End Sub

' =============================

' ====================================================================
' iter17 (2026-06-01): Trailing function-block restored after SOP-TRUNC.
' RoleKanri / M02_SHEET_NAME / M03_SHEET_NAME / NormalizeFieldType were
' silently cut off in the bas even though callers above reference them.
' Restored as ChrW() source so VBE compile succeeds for case 16.
' ====================================================================
Private Function RoleKanri() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0244] modEntryFormat.RoleKanri ENTER"  ' [ADR-0100]
    RoleKanri = ChrW(&H7BA1) & ChrW(&H7406)
End Function

Private Function M02_SHEET_NAME() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0245] modEntryFormat.M02_SHEET_NAME ENTER"  ' [ADR-0100]
    M02_SHEET_NAME = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & _
                     ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & _
                     ChrW(&H4E00) & ChrW(&H89A7)
End Function

Private Function M03_SHEET_NAME() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0246] modEntryFormat.M03_SHEET_NAME ENTER"  ' [ADR-0100]
    M03_SHEET_NAME = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & _
                     ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & _
                     ChrW(&H8A2D) & ChrW(&H8A08)
End Function

' iter18 ADR-0090: preview sheet name (M-04) via ChrW so CP932 round-trip
' through Edit/Write does not mojibake the literal. Previous "?v???r???["
' inline literal silently broke Btn_PreviewInDesign / Btn_PreviewFormat.
Public Function PreviewSheetName_M04() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0247] modEntryFormat.PreviewSheetName_M04 ENTER"  ' [ADR-0100]
    PreviewSheetName_M04 = ChrW(&H30D7) & ChrW(&H30EC) & ChrW(&H30D3) & _
                           ChrW(&H30E5) & ChrW(&H30FC)
End Function

' iter18 ADR-0090: map ui_seed M-03 inputDataKey -> in-memory formatDict
' slot name expected by SaveFormat_Workflow / LoadFormat_Workflow.
Public Function MapInputDataKeyToFieldName_M03(ByVal inputKey As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0248] modEntryFormat.MapInputDataKeyToFieldName_M03 ENTER"  ' [ADR-0100]
    Select Case inputKey
        Case "targetFormat", "FormatID", "formatId", "fmtId"
            MapInputDataKeyToFieldName_M03 = "FormatID"
        Case "formatName", "FormatName", "fmtName"
            MapInputDataKeyToFieldName_M03 = "FormatName"
        Case "formatVersion", "FormatVersion", "fmtVersion"
            MapInputDataKeyToFieldName_M03 = "FormatVersion"
        Case Else
            MapInputDataKeyToFieldName_M03 = ""
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0249] modEntryFormat.MapInputDataKeyToFieldName_M03 EXIT-OK"  ' [ADR-0100]
End Function

' NormalizeFieldType: 6 UI labels -> 4 canonical base types
' (per ADR-0072 7.3). Already-canonical values pass through.
Public Function NormalizeFieldType(ByVal raw As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0250] modEntryFormat.NormalizeFieldType ENTER"  ' [ADR-0100]
    Dim t As String
    t = Trim(raw)
    Dim single_line As String, multi_line As String
    Dim date_type As String, select_type As String
    single_line = ChrW(&H5358) & ChrW(&H4E00) & ChrW(&H884C)
    multi_line = ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C)
    date_type = ChrW(&H65E5) & ChrW(&H4ED8)
    select_type = ChrW(&H9078) & ChrW(&H629E)
    Dim lbl_1line As String, lbl_multiline As String, lbl_number As String
    Dim lbl_date As String, lbl_select As String, lbl_check As String
    lbl_1line = "1" & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8)
    lbl_multiline = ChrW(&H8907) & ChrW(&H6570) & ChrW(&H884C) & ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8)
    lbl_number = ChrW(&H6570) & ChrW(&H5024)
    lbl_date = ChrW(&H65E5) & ChrW(&H4ED8)
    lbl_select = ChrW(&H9078) & ChrW(&H629E)
    lbl_check = ChrW(&H30C1) & ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF)
    If t = lbl_1line Or t = lbl_number Then
        NormalizeFieldType = single_line
    ElseIf t = lbl_multiline Then
        NormalizeFieldType = multi_line
    ElseIf t = lbl_date Then
        NormalizeFieldType = date_type
    ElseIf t = lbl_select Or t = lbl_check Then
        NormalizeFieldType = select_type
    Else
        NormalizeFieldType = t
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0251] modEntryFormat.NormalizeFieldType EXIT-OK"  ' [ADR-0100]
End Function
```
