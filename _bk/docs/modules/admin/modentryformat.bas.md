---
title: modEntryFormat.bas
description: modEntryFormat.bas のソースコード（コピペ用）
---

# modEntryFormat.bas

**配置先**: `管理.xlsm` 用の VBA モジュール  
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\admin\`
- ファイル名: `modEntryFormat.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

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
' ADR-0089: ui_seed role dir name SSOT.
'   ui_seed dir is CJK-named (U+7BA1 U+7406 admin,
'   U+691C U+7D22 search, U+767B U+9332 U+4FEE U+6B63
'   register). The loader resolves files via
'   GetUiDir() & roleName & "\" & screenId & ".txt".
'   Hard-coded "Kanri" / "Kensaku" ASCII romaji never
'   matched the actual dir and silently returned an empty
'   Collection (no-op). Source must stay ASCII (ADR-0070 /
'   ADR-0063 CP932 round-trip safety), so literals are built
'   at runtime via ChrW (VBA Const cannot hold a ChrW
'   expression). RoleKanri() returns the 2-char CJK string
'   U+7BA1 U+7406.
' ------------------------------------------------------------
Private Function RoleKanri() As String
    RoleKanri = ChrW(&H7BA1) & ChrW(&H7406)
End Function

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
Private Const M02_SHEET_NAME      As String = "?t?H?[?}?b?g??"   ' v2.3 (2026-05-26): ?\????????? (?? "M-02")
Private Const M02_DATA_FIRST_ROW  As Long = 11
Private Const M02_DATA_LAST_ROW   As Long = 60
Private Const M02_COL_ROWCHECK    As String = "B"
Private Const M02_COL_NO          As String = "C"
Private Const M02_COL_FMT_ID      As String = "D"
Private Const M02_COL_FMT_NAME    As String = "E"
Private Const M02_COL_ID_PATTERN  As String = "F"
Private Const M02_COL_NEXT_NUM    As String = "G"
Private Const M02_COL_VERSION     As String = "H"
Private Const M02_COL_FIELD_COUNT As String = "I"
Private Const M02_COL_KNW_COUNT   As String = "J"
Private Const M02_COL_UPDATED     As String = "K"
Private Const M02_COL_STATUS      As String = "L"

' Status text written into [FORMAT].Status when the user disables
' a format. The matching "enabled" text is intentionally not used
' inside this module - any non-disabled format displays whatever
' value the on-disk [FORMAT].Status carries (typically empty).
Private Const FMT_STATUS_DISABLED As String = "????"

' [FORMAT].Status key name (the stanza key inside <fmtId>.txt).
Private Const FMT_KEY_STATUS As String = "Status"

' ------------------------------------------------------------
' M-03 cell layout (v2.3 - reusing the v2.2 UI seed positions
' that Btn_SaveFormat / Btn_LoadFormat already drive through
' modUILoader. The IDPattern slot (I3) is cleared on new-draft
' as a defensive step, even though the v2.3 design doc removes
' the IDPattern input; the seed cleanup keeps a stale value from
' surviving a "draft" reset.)
' ------------------------------------------------------------
Private Const M03_SHEET_NAME       As String = "?t?H?[?}?b?g??v"   ' v2.3: ?? "M-03"
Private Const M03_CELL_FMT_ID      As String = "C3"
Private Const M03_CELL_FMT_NAME    As String = "E3"
Private Const M03_CELL_ID_PATTERN  As String = "I3"
Private Const M03_GRID_FIRST_ROW   As Long = 7
Private Const M03_GRID_LAST_ROW    As Long = 56
Private Const M03_GRID_CLEAR_FROM  As String = "B"
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
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Set result(FIELDS_KEY) = New Collection
    If uiSections Is Nothing Then
        Set BuildFormatDictFromCells = result
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
                If Len(fieldName) > 0 And Len(cellAddr) > 0 Then
                    result(fieldName) = clsCellIO.ReadCellValue(target, cellAddr)
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
                If Len(fieldName) > 0 And Len(cellAddr) > 0 Then
                    If formatDict.Exists(fieldName) Then
                        Dim v As Variant
                        v = formatDict(fieldName)
                        If IsNull(v) Then
                            clsCellIO.WriteCellValue target, cellAddr, ""
                        Else
                            clsCellIO.WriteCellValue target, cellAddr, CStr(v)
                        End If
                    End If
                End If
            Case "GRID"
                If formatDict.Exists(FIELDS_KEY) Then
                    clsGridIO.WriteGridFields target, sec, formatDict(FIELDS_KEY)
                End If
        End Select
    Next i
End Sub

Public Function SaveFormat_Workflow(ByVal target As Object, ByVal uiSections As Collection) As String
    On Error GoTo ErrHandler
    Dim fmtDict As Object
    Set fmtDict = BuildFormatDictFromCells(target, uiSections)
    Dim fmtId As String
    If fmtDict.Exists("FormatID") Then fmtId = CStr(fmtDict("FormatID"))
    If Len(fmtId) = 0 Then
        SaveFormat_Workflow = ""
        Exit Function
    End If
    Dim sections As Collection
    Set sections = New Collection
    Dim fmtSec As ClsStanzaSection
    Set fmtSec = New ClsStanzaSection
    fmtSec.Init "FORMAT", 1
    fmtSec.SetValue "FormatID", fmtId
    If fmtDict.Exists("FormatName") Then fmtSec.SetValue "FormatName", CStr(fmtDict("FormatName"))
    If fmtDict.Exists("FormatVersion") Then fmtSec.SetValue "FormatVersion", CStr(fmtDict("FormatVersion"))
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
            ' Phase O-5 (ADR-0072 §7.3): M-03 UI labels (6 種: 1行テキスト/複数行テキスト/数値/日付/選択/チェック)
            ' を canonical 4 base type (単一行/複数行/日時/選択) に集約してから format.txt に書く。
            If row.Exists("Type") Then fldSec.SetValue "FieldType", NormalizeFieldType(CStr(row("Type")))
            If row.Exists("Required") Then fldSec.SetValue "Required", CStr(row("Required"))
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
    Exit Function
ErrHandler:
    On Error Resume Next
    Dim oLogger032 As clsLogger
    Set oLogger032 = New clsLogger
    oLogger032.Init ThisWorkbook.Worksheets("LOG")
    oLogger032.LogWarn "modEntryFormat", "SaveFormat_Workflow", "Format save failed: " & Err.Description, fmtId, "VALIDATE-WARN-WW-032"
    On Error GoTo 0
    SaveFormat_Workflow = ""
End Function

Public Function LoadFormat_Workflow(ByVal target As Object, ByVal uiSections As Collection, ByVal formatId As String) As Boolean
    On Error GoTo ErrHandler
    If Len(formatId) = 0 Then
        LoadFormat_Workflow = False
        Exit Function
    End If
    Dim loaded As Collection
    Set loaded = modFormatLoader.LoadFormat(formatId)
    If loaded.Count = 0 Then
        LoadFormat_Workflow = False
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
                If sec.HasKey("FormatID") Then formatDict("FormatID") = sec.GetValue("FormatID")
                If sec.HasKey("FormatName") Then formatDict("FormatName") = sec.GetValue("FormatName")
                If sec.HasKey("FormatVersion") Then formatDict("FormatVersion") = sec.GetValue("FormatVersion")
            Case "FIELD"
                Dim row As Object
                Set row = CreateObject("Scripting.Dictionary")
                If sec.HasKey("FieldName") Then row("Name") = sec.GetValue("FieldName")
                If sec.HasKey("FieldType") Then row("Type") = sec.GetValue("FieldType")
                If sec.HasKey("Required") Then row("Required") = sec.GetValue("Required")
                formatDict(FIELDS_KEY).Add row
        End Select
    Next i
    WriteFormatDictToCells target, uiSections, formatDict
    LoadFormat_Workflow = True
    Exit Function
ErrHandler:
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
    On Error GoTo ErrHandler
    Dim ui As Collection: Set ui = modUILoader.LoadUiDefinition(RoleKanri(), "M-03")
    If ui.Count = 0 Then Exit Sub
    SaveFormat_Workflow ActiveSheet, ui
    On Error Resume Next
    Dim oLogger007 As clsLogger
    Set oLogger007 = New clsLogger
    oLogger007.Init ThisWorkbook.Worksheets("LOG")
    oLogger007.LogInfo "modEntryFormat", "Btn_SaveFormat", "Btn_SaveFormat done", "", "SAVE-EXIT-OK-II-007"
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Btn_SaveFormat: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_LoadFormat
' Role: M-03 "???" (?t?H?[?}?b?g?C???Z?b?g) - reads the fmtId
'       from the M-03 cell, loads <fmtId>.txt and writes it back
'       into the design cells.
' ============================================================
Public Sub Btn_LoadFormat()
    On Error GoTo ErrHandler
    Dim ui As Collection: Set ui = modUILoader.LoadUiDefinition(RoleKanri(), "M-03")
    If ui.Count = 0 Then Exit Sub
    Dim tempDict As Object: Set tempDict = BuildFormatDictFromCells(ActiveSheet, ui)
    Dim fmtId As String
    If tempDict.Exists("FormatID") Then fmtId = CStr(tempDict("FormatID"))
    If Len(fmtId) = 0 Then Exit Sub
    LoadFormat_Workflow ActiveSheet, ui, fmtId
    Exit Sub
ErrHandler:
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
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M03_SHEET_NAME)

    ' 1. Clear the design inputs.
    ws.Range(M03_CELL_FMT_ID).ClearContents
    ws.Range(M03_CELL_FMT_NAME).ClearContents
    ws.Range(M03_CELL_ID_PATTERN).ClearContents

    ' 2. Clear the field-design grid (data rows only - header row
    '    sits at M03_GRID_FIRST_ROW - 1 and is left untouched).
    Dim r As Long
    For r = M03_GRID_FIRST_ROW To M03_GRID_LAST_ROW
        On Error Resume Next
        ws.Range(M03_GRID_CLEAR_FROM & r & ":" & M03_GRID_CLEAR_TO & r).ClearContents
        On Error GoTo ErrHandler
    Next r

    ' 3. Seed the FormatID cell with the next auto-numbered id.
    Dim nextId As String
    nextId = AutoNumberNextFmtId()
    ws.Range(M03_CELL_FMT_ID).Value = nextId

    ' 4. Activate the sheet so the user lands on the draft form.
    ws.Activate

    On Error Resume Next
    Dim lg As clsLogger
    Set lg = NewLogger()
    If Not lg Is Nothing Then
        lg.LogInfo "modEntryFormat", "Btn_NewFormatDraft", _
                   "M-03 new-draft seeded with fmtId=" & nextId, _
                   nextId, "M03-NEWDRAFT-OK-II-040"
    End If
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Btn_NewFormatDraft: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_AddField
' Role: M-03 "?t?B?[???h???" - selects the first empty row in
'       the field-design grid so the user can start typing.
' ============================================================
Public Sub Btn_AddField()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M03_SHEET_NAME)
    Dim r As Long
    For r = M03_GRID_FIRST_ROW To M03_GRID_LAST_ROW
        If Len(Trim(CStr(ws.Range("C" & r).Value))) = 0 Then
            ws.Activate
            On Error Resume Next
            ws.Range("C" & r).Select
            On Error GoTo 0
            Exit Sub
        End If
    Next r
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Btn_AddField: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_AddField_v21
' Role: legacy v2.1 entry kept for the wired button surface; it
'       hunts the first empty row in column B and selects it.
'       Btn_AddField (column-C scan) is the v2.3 replacement.
' ============================================================
Public Sub Btn_AddField_v21()
    On Error GoTo ErrHandler
    Dim ws As Object: Set ws = ActiveSheet
    Dim r As Long
    For r = 7 To 50
        If Len(CStr(ws.Range("B" & r).Value)) = 0 Then
            On Error Resume Next
            ws.Range("B" & r).Select
            On Error GoTo 0
            Debug.Print "[Btn_AddField_v21] selected B" & r
            Exit Sub
        End If
    Next r
    Debug.Print "[Btn_AddField_v21] GRID full, no action"
    Exit Sub
ErrHandler:
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
    On Error GoTo ErrHandler
    ThisWorkbook.Worksheets("?v???r???[").Activate
    RenderM04Preview
    Exit Sub
ErrHandler:
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
    On Error GoTo ErrHandler
    Dim fmtId As String
    fmtId = PickM02CheckedRowFmtId()
    If Len(fmtId) = 0 Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M03_SHEET_NAME)
    ws.Activate
    Dim ui As Collection
    Set ui = modUILoader.LoadUiDefinition(RoleKanri(), "M-03")
    If ui.Count = 0 Then Exit Sub
    LoadFormat_Workflow ws, ui, fmtId
    Exit Sub
ErrHandler:
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
    On Error GoTo ErrHandler
    Dim fmtId As String
    fmtId = PickM02CheckedRowFmtId()
    If Len(fmtId) = 0 Then Exit Sub
    ThisWorkbook.Worksheets(M03_SHEET_NAME).Range(M03_CELL_FMT_ID).Value = fmtId
    ThisWorkbook.Worksheets("?v???r???[").Activate
    RenderM04Preview
    Exit Sub
ErrHandler:
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
    On Error GoTo ErrHandler

    ' 1. Pick the single checked row.
    Dim fmtId As String
    fmtId = PickM02CheckedRowFmtId()
    If Len(fmtId) = 0 Then Exit Sub

    ' 2. Load the existing format definition.
    Dim sections As Collection
    Set sections = modFormatLoader.LoadFormat(fmtId)
    If sections Is Nothing Then
        WarnUser "Disable format", _
                 "Format " & fmtId & " could not be loaded."
        Exit Sub
    End If
    If sections.Count = 0 Then
        WarnUser "Disable format", _
                 "Format " & fmtId & " has no definition file."
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
        WarnUser "Disable format", _
                 "Format " & fmtId & " is missing the [FORMAT] section."
        Exit Sub
    End If
    fmtSec.SetValue FMT_KEY_STATUS, FMT_STATUS_DISABLED

    ' 4. Persist back to disk via modFormatLoader.SaveFormat.
    Dim rc As Long
    rc = modFormatLoader.SaveFormat(fmtId, sections)
    If rc <> 0 Then
        WarnUser "Disable format", _
                 "Failed to write disable flag for " & fmtId & _
                 " (rc=" & rc & ")."
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
    Exit Sub
ErrHandler:
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
    On Error GoTo ErrHandler
    Dim fmtId As String
    fmtId = PickM02CheckedRowFmtId()
    If Len(fmtId) = 0 Then Exit Sub
    If Not IsHeadless() Then
        If MsgBox("Delete format " & fmtId & " ?", vbYesNo + vbExclamation, _
                  "Delete format") <> vbYes Then Exit Sub
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
    Exit Sub
ErrHandler:
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
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M02_SHEET_NAME)

    ' 1. Clear the existing data rows (B..L on every data row).
    ClearM02GridDataRows ws

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
    Exit Sub
ErrHandler:
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
    Dim addr As String
    addr = M02_COL_ROWCHECK & M02_DATA_FIRST_ROW & ":" & _
           M02_COL_STATUS & M02_DATA_LAST_ROW
    On Error Resume Next
    ws.Range(addr).ClearContents
    On Error GoTo 0
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
    ws.Range(M02_COL_ID_PATTERN & rowNum).Value = fmtIdPat
    ws.Range(M02_COL_NEXT_NUM & rowNum).Value = fmtNxtNum
    ws.Range(M02_COL_VERSION & rowNum).Value = fmtVer
    ws.Range(M02_COL_FIELD_COUNT & rowNum).Value = fieldCount
    ws.Range(M02_COL_KNW_COUNT & rowNum).Value = knwCount
    ws.Range(M02_COL_UPDATED & rowNum).Value = updatedAt
    ws.Range(M02_COL_STATUS & rowNum).Value = fmtStat
    Exit Sub
ErrHandler:
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
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(M02_SHEET_NAME)

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
        WarnUser "Format list", _
                 "Check exactly one format row first."
        PickM02CheckedRowFmtId = ""
        Exit Function
    End If
    If checkedCount > 1 Then
        WarnUser "Format list", _
                 "Multiple rows checked. Check exactly one."
        PickM02CheckedRowFmtId = ""
        Exit Function
    End If

    Dim fmtId As String
    fmtId = Trim(CStr(ws.Range(M02_COL_FMT_ID & hitRow).Value))
    If Len(fmtId) = 0 Then
        WarnUser "Format list", _
                 "The checked row has no FormatID."
        PickM02CheckedRowFmtId = ""
        Exit Function
    End If
    PickM02CheckedRowFmtId = fmtId
    Exit Function
ErrHandler:
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
    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        IsRowChecked = CBool(v)
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
            Exit Function
        End If
    Next i
    AutoNumberNextFmtId = FMT_ID_PREFIX & PadLeft("1", FMT_ID_WIDTH, "0")
    Exit Function
ErrHandler:
    AutoNumberNextFmtId = FMT_ID_PREFIX & PadLeft("1", FMT_ID_WIDTH, "0")
End Function

' ============================================================
' Private Function: PadLeft
' Role: zero-pad a string on the left to the given width. VBA
'       lacks a built-in, so we keep a tiny local helper rather
'       than reaching for Format(... , "000") which would couple
'       us to a per-width string template.
' ============================================================
Private Function PadLeft(ByVal s As String, ByVal width As Long, ByVal pad As String) As String
    Dim out As String
    out = s
    Do While Len(out) < width
        out = pad & out
    Loop
    PadLeft = out
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
    On Error GoTo ErrHandler
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
            r = MsgBox("Unsaved input. Save before back to main?" & vbCrLf & _
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
                Exit Sub
        End Select
    Else
        NavigateToMain_v21
    End If
    Exit Sub
ErrHandler:
    On Error Resume Next
    Dim oLogger032b As clsLogger
    Set oLogger032b = New clsLogger
    oLogger032b.Init ThisWorkbook.Worksheets("LOG")
    oLogger032b.LogError "modEntryFormat", "Btn_BackToMain_v21", "Back to main transition error: " & Err.Description, "", "BACKTOMAIN-ERR-EE-032"
    On Error GoTo 0
    If Not IsHeadless() Then
        MsgBox "Btn_BackToMain_v21 error: " & Err.Description, vbCritical, "Error"
    Else
        Debug.Print "[ERR] Btn_BackToMain_v21: " & Err.Number & " " & Err.Description
    End If
End Sub

Private Function HasUnsavedInput_M03(ByVal ws As Object) As Boolean
    On Error GoTo ErrHandler
    Dim v1 As String, v2 As String, v3 As String
    v1 = CStr(ws.Range("C3").Value)
    v2 = CStr(ws.Range("C4").Value)
    v3 = CStr(ws.Range("C5").Value)
    If Len(v1) > 0 Or Len(v2) > 0 Or Len(v3) > 0 Then
        HasUnsavedInput_M03 = True
        Exit Function
    End If
    If Len(CStr(ws.Range("B7").Value)) > 0 Then
        HasUnsavedInput_M03 = True
        Exit Function
    End If
    HasUnsavedInput_M03 = False
    Exit Function
ErrHandler:
    HasUnsavedInput_M03 = False
End Function

Private Sub NavigateToMain_v21()
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
    On Error GoTo ErrHandler
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
            MsgBox "Migration is not available in v2.3 (M-12 auto-reflect retired).", _
                   vbInformation, "Result"
        Else
            Debug.Print "[Btn_ConfirmDiff] migrate path is a no-op under v2.3"
        End If
    End If
    Exit Sub
ErrHandler:
    If Not IsHeadless() Then
        MsgBox "Btn_ConfirmDiff error: " & Err.Description, vbCritical, "Error"
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
Public Function IsHeadless() As Boolean
    On Error Resume Next
    Dim notInteractive As Boolean
    Dim notVisible As Boolean
    notInteractive = (Application.Interactive = False)
    notVisible = (Application.Visible = False)
    IsHeadless = (notInteractive Or notVisible)
    On Error GoTo 0
End Function

' ============================================================
' Private Sub: RenderM04Preview
' Role: rerender the M-04 preview grid. Creates a fresh
'       clsFormatPreviewScreen, inits with a minimal spec
'       (ScreenId/SheetName=M-04, no renderer), and calls
'       RenderPreview.
' ============================================================
Private Sub RenderM04Preview()
    On Error GoTo ErrHandler
    Dim spec As clsScreenSpec
    Set spec = New clsScreenSpec
    spec.ScreenId = "M-04"
    spec.SheetName = "M-04"
    Dim prev As clsFormatPreviewScreen
    Set prev = New clsFormatPreviewScreen
    prev.Init Nothing, spec
    prev.RenderPreview
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] RenderM04Preview: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Private Sub: ShowDeleteFormatResult
' Role: pretty-print clsFormatManager.DeleteFormat rc codes.
'       0=success / 1=outside kanri.xlsm reject /
'       2=knowledge-still-using reject / else=failure.
' ============================================================
Private Sub ShowDeleteFormatResult(ByVal fmtId As String, ByVal rc As Long)
    If IsHeadless() Then Exit Sub
    Dim msg As String
    Select Case rc
        Case 0
            msg = "Format " & fmtId & " was deleted."
        Case 1
            msg = "Rejected: format delete is allowed from the admin workbook only."
        Case 2
            msg = "Rejected: knowledge entries still use format " & fmtId & "."
        Case Else
            msg = "Failed to delete format " & fmtId & "."
    End Select
    MsgBox msg, vbInformation, "Delete format"
End Sub

' ============================================================
' Private Function: NewLogger
' Role: create a clsLogger bound to the LOG sheet. Nothing on
'       failure so callers can fall back to Debug.Print.
' ============================================================
Private Function NewLogger() As clsLogger
    On Error GoTo ErrHandler
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ThisWorkbook.Worksheets("LOG")
    Set NewLogger = lg
    Exit Function
ErrHandler:
    Set NewLogger = Nothing
End Function

' ============================================================
' Private Sub: WarnUser
' Role: MsgBox-with-headless-guard, so the warning is silent
'       under COM automation but visible to a real Excel user.
' ============================================================
Private Sub WarnUser(ByVal title As String, ByVal body As String)
    If IsHeadless() Then
        Debug.Print "[WARN " & title & "] " & body
        Exit Sub
    End If
    MsgBox body, vbExclamation, title
End Sub

' ================================================================
' Phase O-5 (2026-05-27): NormalizeFieldType
' Per ADR-0072 §7.3 + Phase N-1 A2, the canonical FieldType enum is
' 4 base types (単一行 / 複数行 / 日時 / 選択). The M-03 UI label
' set is 6 ("1行テキスト" / "複数行テキスト" / "数値" / "日付" / "選択" / "チェック").
' This function maps the 6 UI labels into the 4 base types so the
' format .txt stays canonical and downstream loaders see exactly
' 4 enum values.
' Already-canonical values pass through unchanged.
' ================================================================
Public Function NormalizeFieldType(ByVal raw As String) As String
    Dim t As String
    t = Trim$(raw)
    Select Case t
        Case "1行テキスト", "数値"
            NormalizeFieldType = "単一行"
        Case "複数行テキスト"
            NormalizeFieldType = "複数行"
        Case "日付"
            NormalizeFieldType = "日時"
        Case "選択", "チェック"
            NormalizeFieldType = "選択"
        Case "単一行", "複数行", "日時"
            NormalizeFieldType = t
        Case Else
            NormalizeFieldType = t
    End Select
End Function
```
