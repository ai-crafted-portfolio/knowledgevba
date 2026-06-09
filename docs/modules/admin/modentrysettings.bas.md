---
title: modEntrySettings.bas
description: modEntrySettings.bas のソースコード（コピペ用）
---

# modEntrySettings.bas

**配置先**: `管理.xlsm` 用の VBA モジュール
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\admin\`
- ファイル名: `modEntrySettings.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modEntrySettings"
' ============================================================
' modEntrySettings (Entry layer, kanri.xlsm only / v2.3)
' Role:
'   kanri.xlsm settings entry points. v2.3 scope (M-10/M-11/M-12/M-14):
'     M-10 storage   : path 4 dirs (data_dir / format_dir / ui_dir / backup_dir)
'     M-11 settings  : debugLevel
'     M-12 check     : read-only format vs. knowledge consistency check
'     M-14 log admin : delete all log rows
'   v2.3 changes vs. v2.2:
'     - Removed all M-12 auto-reflect (migrate / cancel / restore / row-mapping)
'       symbols (clsFieldMigrator is archived out of the project).
'     - Removed all M-13 file-format symbols (M-13 screen retired).
'     - M-10 / M-11 save buttons now run: validate -> SaveConfigKeys
'       (writes config.txt via modConfigLoader) -> SetConfigKeys (sync
'       in-memory holder via modConfigHolder.SetConfigKeys).
'     - Removed M-10 / M-11 cancel buttons (Btn_Cancel*_v21).
'     - Added Btn_DeleteAllLog (M-14 destructive clear).
'     - Added Btn_CheckFormat (M-12 read-only format check).
' Naming (after v2.3):
'   M-11 settings (debugLevel):
'     Btn_OpenSettings_v21   -- open (holder -> sheet)
'     Btn_SaveSettings_v21   -- validate -> SaveConfigKeys -> SetConfigKeys
'   M-10 storage (4 dirs):
'     Btn_OpenStorage_v21    -- open (holder -> sheet)
'     Btn_SaveStorage_v21    -- validate -> SaveConfigKeys -> SetConfigKeys
'   M-12 read-only check:
'     Btn_CheckFormat        -- read-only consistency check, results to grid
'   M-14 log admin:
'     Btn_DeleteAllLog       -- destructive clear of LOG data rows
' Logging:
'   SAVE-EXIT-OK-II-010 / 011 (M-11 Open / Apply OK)
'   SAVE-EXIT-OK-II-013 / 014 (M-10 Open / Apply OK)
'   VALIDATE-WARN-WW-033      (M-11 Apply validation fail)
'   VALIDATE-WARN-WW-035      (M-10 Apply validation fail)
'   BACKTOMAIN-ERR-EE-034     (M-11 Apply exception)
'   BACKTOMAIN-ERR-EE-036     (M-10 Apply exception)
'   LOG-M14-CLEAR-OK / GUARD  (Btn_DeleteAllLog)
'   LOG-M12-CHECK-OK / EMPTY  (Btn_CheckFormat)
' ASCII-only inside VBA string literals (CP932 mojibake avoidance).
' ============================================================
Option Explicit

Private Const MOD_NAME As String = "modEntrySettings"

' SPEC_DRIFT-4b fix (2026-05-30): The previous Const definitions for
' XLSM_KEY / SHEET_SETTINGS / SHEET_STORAGE / SHEET_MIGRATE held
' Japanese string literals that were silently mojibaked (encoded as
' U+FFFD) during a CP932->UTF-8 Edit/Write round-trip, which caused
' SaveConfigKeys to target the wrong config filename (mojibake _config.txt
' instead of "kanri" / Japanese-kanri _config.txt). The keys are now
' produced via ChrW() functions defined near the bottom of the file
' (Functions cannot appear inside the module-level declaration block,
' so they live after the public Subs and helpers). SHEET_LOG remains
' ASCII so kept as a Const here.
Private Const SHEET_LOG As String = "LOG"

' --- M-11 cell layout (debugLevel only) ---
' SPEC_DRIFT-4c fix (2026-05-30): ui_seed M-11.txt v2.3 places the
' debugLevel INPUT at D13:E13 (top-left anchor D13). C6 is inside
' the NOTE merged region A6:G12 so reads/writes silently no-op.
Private Const ADDR_DEBUG_LEVEL As String = "D13"

' --- M-10 cell layout (4 dirs) ---
' SPEC_DRIFT-4b fix (2026-05-30): ui_seed M-10.txt v2.3 layout has the
' 4 INPUT cells at D11:E11 (data_dir) / D12:E12 (format_dir) /
' D13:E13 (ui_dir) / D14:E14 (backup_dir). The top-left anchor of each
' merged range is the cell we read/write.
' Old layout (pre-v2.3): C2/C3/C4/C5 (kept here for traceability).
Private Const ADDR_DATA_DIR    As String = "D11"
Private Const ADDR_FORMAT_DIR  As String = "D12"
Private Const ADDR_UI_DIR      As String = "D13"
Private Const ADDR_BACKUP_DIR  As String = "D14"
Private Const ADDR_CONFIG_DIR  As String = "D15"

' --- M-12 cell layout (format check) ---
' C8 = target format id (matches ui_seed M-12.txt [INPUT] Cell=C8:E8
' merged area, the top-left anchor of the merged C8:E8 input slot).
' The result grid starts at row 15 and uses 4 columns: KnowledgeNo /
' FieldName / CurrentValue / Issue.
Private Const ADDR_M12_FORMAT_ID  As String = "C8"
Private Const M12_RESULT_START_ROW As Long = 15
Private Const M12_RESULT_COL_KNO   As Long = 1   ' column A
Private Const M12_RESULT_COL_FLD   As Long = 2   ' column B
Private Const M12_RESULT_COL_VAL   As Long = 3   ' column C
Private Const M12_RESULT_COL_ISS   As Long = 4   ' column D
' Defensive upper bound when clearing the result grid (avoids touching
' arbitrary cells on a freshly-seeded sheet).
Private Const M12_RESULT_CLEAR_ROWS As Long = 500

' --- LOG sheet layout (mirrors clsLogger / modCommon) ---
Private Const LOG_DATA_START_ROW As Long = 9

' --- config.txt default fallbacks (used only when holder is empty) ---
Private Const DEF_DATA_DIR    As String = "C:\KnowledgeMgr\data\"
Private Const DEF_FORMAT_DIR  As String = "C:\KnowledgeMgr\formats\"
Private Const DEF_UI_DIR      As String = "C:\KnowledgeMgr\ui\"
Private Const DEF_BACKUP_DIR  As String = "C:\KnowledgeMgr\data\backup\"
Private Const DEF_CONFIG_DIR  As String = "C:\KnowledgeMgr\config\"
Private Const DEF_DEBUG_LEVEL As String = "INFO"
' Phase FC (2026-06-03, ADR-0096): M-11 [VALIDATION] / [CHECKBOX] layout.
'   B5     debugLevel selection (data validation list)
'   B7..B9 CheckBoxes (caption only). LinkedCells in column H.
Private Const ADDR_DEBUG_LEVEL_FC As String = "D10"   ' 2026-06-07: new M-11 layout INPUT D10:E10
' 2026-06-07: Checkbox UI for autoReload / migBackup / sysSheetVisibility
' was retired. Constants kept as comments for backward archaeology.
'   ADDR_LC_AUTORELOAD = H7 / ADDR_LC_MIGBACKUP = H8 / ADDR_LC_SYSSHEET = H9
'   DEF_AUTORELOAD/MIGBACKUP/SYSSHEET = "TRUE"



' ============================================================
' ===== M-11 settings (debugLevel) entry =====================
' ============================================================

' ============================================================
' Public Sub: Btn_OpenSettings_v21
' Role: open M-11 settings sheet; mirror debugLevel from holder.
'       Assumes config holder is already initialized at workbook
'       startup.
' ============================================================
Public Sub Btn_OpenSettings_v21()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0252] modEntrySettings.Btn_OpenSettings_v21 ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_OpenSettings_v21"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0253] modEntrySettings.Btn_OpenSettings_v21 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    ws.Activate
    LoadSettingsToSheet ws
    LogInfoSafe "Btn_OpenSettings_v21", _
        "Settings opened", "SAVE-EXIT-OK-II-010"
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0254] modEntrySettings.Btn_OpenSettings_v21 EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0255] modEntrySettings.Btn_OpenSettings_v21 EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_OpenSettings_v21: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_SaveSettings_v21
' Role: M-11 sheet -> validate -> persist via SaveConfigKeys ->
'       sync in-memory via SetConfigKeys. Only debugLevel is
'       owned by M-11; the path keys remain unchanged on disk
'       (SaveConfigKeys keeps non-listed keys as-is).
' ============================================================
Public Sub Btn_SaveSettings_v21()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0256] modEntrySettings.Btn_SaveSettings_v21 ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_SaveSettings_v21"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0257] modEntrySettings.Btn_SaveSettings_v21 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)

    Dim ngLabel As String
    Dim ngAddr As String
    If Not ValidateSettingsSheet(ws, ngLabel, ngAddr) Then
        LogWarnSafe "Btn_SaveSettings_v21", _
            "Validation failed at " & ngAddr & " (" & ngLabel & ")", _
            "VALIDATE-WARN-WW-033"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0258] modEntrySettings.Btn_SaveSettings_v21 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ApplySettingsToHolder ws

    LogInfoSafe "Btn_SaveSettings_v21", _
        "Settings applied (config.txt + holder)", "SAVE-EXIT-OK-II-011"
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0259] modEntrySettings.Btn_SaveSettings_v21 EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0260] modEntrySettings.Btn_SaveSettings_v21 EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    LogErrorSafe "Btn_SaveSettings_v21", _
        "Apply failed: " & Err.Number & " " & Err.Description, _
        "BACKTOMAIN-ERR-EE-034"
    Debug.Print "[ERR] Btn_SaveSettings_v21: " & Err.Number & " " & Err.Description
End Sub


' ============================================================
' ===== M-10 storage (4 dirs) entry ==========================
' ============================================================

' ============================================================
' Public Sub: Btn_OpenStorage_v21
' Role: SPEC_DRIFT-4 fix (2026-05-30): open FolderPicker dialog and
'   write the selected folder path into the D-column input cell of
'   the row whose A-column checkbox glyph is U+25A0 (filled square).
'   When invoked with no row checked, show an instruction message.
'   M-10 sheet layout (per ui_seed M-10.txt):
'     A11/A12/A13/A14 : checkbox glyph (U+25A0 = checked, U+25A1 = unchecked)
'     D11:E11 .. D14:E14 : input cell for data_dir / format_dir /
'                          ui_dir / backup_dir
'   Only one row is expected to be checked at a time (DefaultChecked
'   in ui_seed is TRUE only on A11).
' ============================================================
Public Sub Btn_OpenStorage_v21()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0261] modEntrySettings.Btn_OpenStorage_v21 ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_OpenStorage_v21"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0262] modEntrySettings.Btn_OpenStorage_v21 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STORAGE)
    ws.Activate

    ' Locate the row whose checkbox cell holds U+25A0 (filled square).
    Dim targetRow As Long
    targetRow = 0
    Dim i As Long
    For i = 11 To 14
        If CStr(ws.Cells(i, 1).Value) = ChrW(&H25A0) Then
            targetRow = i
            Exit For
        End If
    Next i

    If targetRow = 0 Then
        ' "saki ni hidari no check wo irete kudasai" (no checkbox is checked)
        If Not modCommon.IsHeadless() Then
            MsgBox ChrW(&H5148) & ChrW(&H306B) & ChrW(&H5DE6) & _
                   ChrW(&H306E) & ChrW(&H30C1) & ChrW(&H30A7) & _
                   ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H3092) & _
                   ChrW(&H5165) & ChrW(&H308C) & ChrW(&H3066) & _
                   ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & _
                   ChrW(&H3044), vbExclamation, _
                   ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & _
                   ChrW(&H8A2D) & ChrW(&H5B9A)
        End If
        LogWarnSafe "Btn_OpenStorage_v21", _
            "no checkbox row marked", "SAVE-EXIT-OK-II-013"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0263] modEntrySettings.Btn_OpenStorage_v21 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ' Headless guard (2026-06-01): COM-automated runs never show the
    ' FolderPicker - they would block waiting for a non-existent user.
    ' In that case log and exit; the M-10 row is left untouched.
    If modCommon.IsHeadless() Then
        LogInfoSafe "Btn_OpenStorage_v21", _
            "headless-skip FolderPicker row=" & targetRow, _
            "SAVE-EXIT-OK-II-013"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0264] modEntrySettings.Btn_OpenStorage_v21 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ' FolderPicker dialog.
    Dim fd As Object
    Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker = 4
    fd.AllowMultiSelect = False
    ' Title: "kakuno-saki folder wo erabu"
    fd.Title = ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & _
               ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30EB) & _
               ChrW(&H30C0) & ChrW(&H3092) & ChrW(&H9078) & _
               ChrW(&H3076)

    If fd.Show <> -1 Then
        LogInfoSafe "Btn_OpenStorage_v21", _
            "FolderPicker cancelled row=" & targetRow, _
            "SAVE-EXIT-OK-II-013"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0265] modEntrySettings.Btn_OpenStorage_v21 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    Dim selectedPath As String
    selectedPath = CStr(fd.SelectedItems(1))

    ' Write to D-column input cell (D11..D14, anchor of merged D:E).
    ws.Cells(targetRow, 4).Value = selectedPath

    LogInfoSafe "Btn_OpenStorage_v21", _
        "Storage folder picked row=" & targetRow & " path=" & selectedPath, _
        "SAVE-EXIT-OK-II-013"
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0266] modEntrySettings.Btn_OpenStorage_v21 EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0267] modEntrySettings.Btn_OpenStorage_v21 EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    LogErrorSafe "Btn_OpenStorage_v21", _
        "FolderPicker failed: " & Err.Number & " " & Err.Description, _
        "BACKTOMAIN-ERR-EE-036"
    Debug.Print "[ERR] Btn_OpenStorage_v21: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Public Sub: Btn_SaveStorage_v21
' Role: M-10 sheet -> validate -> persist via SaveConfigKeys ->
'       sync in-memory via SetConfigKeys. Only the 4 path keys
'       are owned by M-10; debugLevel stays untouched on disk.
' ============================================================
Public Sub Btn_SaveStorage_v21()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0268] modEntrySettings.Btn_SaveStorage_v21 ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_SaveStorage_v21"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0269] modEntrySettings.Btn_SaveStorage_v21 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STORAGE)

    Dim ngLabel As String
    Dim ngAddr As String
    If Not ValidateStorageSheet(ws, ngLabel, ngAddr) Then
        LogWarnSafe "Btn_SaveStorage_v21", _
            "Validation failed at " & ngAddr & " (" & ngLabel & ")", _
            "VALIDATE-WARN-WW-035"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0270] modEntrySettings.Btn_SaveStorage_v21 EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ApplyStorageToHolder ws

    LogInfoSafe "Btn_SaveStorage_v21", _
        "Storage settings applied (config.txt + holder)", _
        "SAVE-EXIT-OK-II-014"
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0271] modEntrySettings.Btn_SaveStorage_v21 EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0272] modEntrySettings.Btn_SaveStorage_v21 EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    LogErrorSafe "Btn_SaveStorage_v21", _
        "Apply failed: " & Err.Number & " " & Err.Description, _
        "BACKTOMAIN-ERR-EE-036"
    Debug.Print "[ERR] Btn_SaveStorage_v21: " & Err.Number & " " & Err.Description
End Sub


' ============================================================
' ===== M-11 private helpers (debugLevel) ====================
' ============================================================

' ============================================================
' Private Sub: LoadSettingsToSheet
' Role: holder -> M-11 sheet (debugLevel).
' ============================================================
Private Sub LoadSettingsToSheet(ByVal ws As Worksheet)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0273] modEntrySettings.LoadSettingsToSheet ENTER"  ' [ADR-0100]
    ' Legacy [INPUT] at D13 (kept for backwards compatibility).
    ws.Range(ADDR_DEBUG_LEVEL).Value = SafeGetCfg("debugLevel", DEF_DEBUG_LEVEL)
    ' Phase FC: also mirror to the new [VALIDATION] cell B5.
    On Error Resume Next
    ws.Range(ADDR_DEBUG_LEVEL_FC).Value = SafeGetCfg("debugLevel", DEF_DEBUG_LEVEL)
    ' Phase FC: pre-fill the CheckBox LinkedCells in column H.
    ' 2026-06-07: checkbox cells retired; no read-from-config needed.
    ' Sync CheckBox shape values to LinkedCell state (paranoia: some
    ' Excel COM hosts skip the shape->cell mirror on initial open).
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0274] modEntrySettings.LoadSettingsToSheet EXIT-OK"  ' [ADR-0100]
End Sub

' ============================================================
' Public Sub: EnsureSettingsCellsEditable
' Role: [BUGFIX 2026-06-06] Make M-11 settings input cells (B5
'       debugLevel dropdown, H7/H8/H9 CheckBox LinkedCells) user-
'       editable by setting Locked=False. clsSheetRenderer protects
'       the sheet with all cells Locked=True by default; without
'       these unlocks the user gets "the cell is on a protected
'       sheet" dialog when clicking a CheckBox (the LinkedCell
'       update happens as user write, not as a macro write, so
'       UserInterfaceOnly=True does not exempt it). Called from
'       ThisWorkbook OnTime startup hook after Btn_RefreshAllSheets.
' ============================================================
Public Sub EnsureSettingsCellsEditable()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-EU01] modEntrySettings.EnsureSettingsCellsEditable ENTER"
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    ws.Range(ADDR_DEBUG_LEVEL_FC).Locked = False
    ' 2026-06-07: checkbox cells retired; no Locked-toggle needed.
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-EU02] modEntrySettings.EnsureSettingsCellsEditable EXIT-OK"
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-EU03] modEntrySettings.EnsureSettingsCellsEditable EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description
End Sub

' ============================================================
' Private Sub: ApplySettingsToHolder
' Role: M-11 sheet -> config.txt + holder, via the v2.3
'       SaveConfigKeys + SetConfigKeys pair (single-key dict).
'       SaveConfigKeys leaves other keys in config.txt untouched.
' ============================================================
Private Sub ApplySettingsToHolder(ByVal ws As Worksheet)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0275] modEntrySettings.ApplySettingsToHolder ENTER"  ' [ADR-0100]
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim dl As String
    ' B5 (new) wins over D13 (legacy) when both are non-empty.
    dl = CStr(ws.Range(ADDR_DEBUG_LEVEL_FC).Value)
    If Len(Trim$(dl)) = 0 Then
        dl = CStr(ws.Range(ADDR_DEBUG_LEVEL).Value)
    End If
    If Len(Trim$(dl)) = 0 Then dl = DEF_DEBUG_LEVEL
    d("debugLevel") = dl
    ' CheckBox states via LinkedCell (TRUE / FALSE).
    ' 2026-06-07: checkbox UI retired; not writing 3 booleans to dict.
    PersistConfigKeys d
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0276] modEntrySettings.ApplySettingsToHolder EXIT-OK"  ' [ADR-0100]
End Sub

' Convert a VBA Boolean into the canonical TRUE/FALSE string token used
' in config.txt (matches the v2.3 admin / settei convention).
Private Function BoolToToken(ByVal b As Boolean) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0277] modEntrySettings.BoolToToken ENTER"  ' [ADR-0100]
    If b Then
        BoolToToken = "TRUE"
    Else
        BoolToToken = "FALSE"
    End If
End Function

' ============================================================
' Private Function: ValidateSettingsSheet
' Role: debugLevel must be non-empty.
' Return: True = OK, False = NG (outLabel / outAddr carry the NG)
' ============================================================
Private Function ValidateSettingsSheet(ByVal ws As Worksheet, _
                                       ByRef outLabel As String, _
                                       ByRef outAddr As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0278] modEntrySettings.ValidateSettingsSheet ENTER"  ' [ADR-0100]
    ' [BUGFIX 2026-06-06] FAIL T-M11-CLICK-01: original code only checked D13
    ' (ADDR_DEBUG_LEVEL legacy), but the UI dropdown is at B5
    ' (ADDR_DEBUG_LEVEL_FC). D13 stays empty after Setup_admin (only B5 gets
    ' the data validation default INFO via stanza), so validation always
    ' failed and save silently aborted. Check B5 first, fall back to D13.
    Dim dl As String
    dl = CStr(ws.Range(ADDR_DEBUG_LEVEL_FC).Value)
    If Len(Trim$(dl)) = 0 Then
        dl = CStr(ws.Range(ADDR_DEBUG_LEVEL).Value)
    End If
    If Len(Trim$(dl)) = 0 Then
        outAddr = ADDR_DEBUG_LEVEL_FC
        outLabel = "debugLevel"
        ValidateSettingsSheet = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0279] modEntrySettings.ValidateSettingsSheet EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    ValidateSettingsSheet = True
End Function


' ============================================================
' ===== M-10 private helpers (4 dirs) ========================
' ============================================================

' ============================================================
' Private Sub: LoadStorageToSheet
' Role: holder -> M-10 sheet (4 dirs).
' ============================================================
Public Sub LoadStorageToSheet(ByVal ws As Worksheet)   ' 2026-06-07: Public to allow ThisWorkbook hook
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0280] modEntrySettings.LoadStorageToSheet ENTER"  ' [ADR-0100]
    ws.Range(ADDR_DATA_DIR).Value = SafeGetCfg("data_dir", DEF_DATA_DIR)
    ws.Range(ADDR_FORMAT_DIR).Value = SafeGetCfg("format_dir", DEF_FORMAT_DIR)
    ws.Range(ADDR_UI_DIR).Value = SafeGetCfg("ui_dir", DEF_UI_DIR)
    ws.Range(ADDR_BACKUP_DIR).Value = SafeGetCfg("backup_dir", DEF_BACKUP_DIR)
    ws.Range(ADDR_CONFIG_DIR).Value = SafeGetCfg("config_dir", DEF_CONFIG_DIR)
End Sub

' ============================================================
' Private Sub: ApplyStorageToHolder
' Role: M-10 sheet -> config.txt + holder, via SaveConfigKeys +
'       SetConfigKeys. Only the 4 path keys are written; the
'       debugLevel key in config.txt stays untouched.
' ============================================================
Private Sub ApplyStorageToHolder(ByVal ws As Worksheet)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0281] modEntrySettings.ApplyStorageToHolder ENTER"  ' [ADR-0100]
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("data_dir") = CStr(ws.Range(ADDR_DATA_DIR).Value)
    d("format_dir") = CStr(ws.Range(ADDR_FORMAT_DIR).Value)
    d("ui_dir") = CStr(ws.Range(ADDR_UI_DIR).Value)
    d("backup_dir") = CStr(ws.Range(ADDR_BACKUP_DIR).Value)
    d("config_dir") = CStr(ws.Range(ADDR_CONFIG_DIR).Value)
    PersistConfigKeys d
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0282] modEntrySettings.ApplyStorageToHolder EXIT-OK"  ' [ADR-0100]
End Sub

' ============================================================
' Private Function: ValidateStorageSheet
' Role: each of the 4 path keys must be non-empty.
' Return: True = OK, False = NG (outLabel / outAddr carry the NG)
' ============================================================
Private Function ValidateStorageSheet(ByVal ws As Worksheet, _
                                      ByRef outLabel As String, _
                                      ByRef outAddr As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0283] modEntrySettings.ValidateStorageSheet ENTER"  ' [ADR-0100]
    If Len(CStr(ws.Range(ADDR_DATA_DIR).Value)) = 0 Then
        outAddr = ADDR_DATA_DIR
        outLabel = "data_dir"
        ValidateStorageSheet = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0284] modEntrySettings.ValidateStorageSheet EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If Len(CStr(ws.Range(ADDR_FORMAT_DIR).Value)) = 0 Then
        outAddr = ADDR_FORMAT_DIR
        outLabel = "format_dir"
        ValidateStorageSheet = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0285] modEntrySettings.ValidateStorageSheet EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If Len(CStr(ws.Range(ADDR_UI_DIR).Value)) = 0 Then
        outAddr = ADDR_UI_DIR
        outLabel = "ui_dir"
        ValidateStorageSheet = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0286] modEntrySettings.ValidateStorageSheet EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If Len(CStr(ws.Range(ADDR_BACKUP_DIR).Value)) = 0 Then
        outAddr = ADDR_BACKUP_DIR
        outLabel = "backup_dir"
        ValidateStorageSheet = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0287] modEntrySettings.ValidateStorageSheet EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If Len(CStr(ws.Range(ADDR_CONFIG_DIR).Value)) = 0 Then
        outAddr = ADDR_CONFIG_DIR
        outLabel = "config_dir"
        ValidateStorageSheet = False
        Exit Function
    End If
    ValidateStorageSheet = True
End Function


' ============================================================
' ===== M-14 log admin (Btn_DeleteAllLog) ====================
' ============================================================

' ============================================================
' Public Sub: Btn_DeleteAllLog
' Role: M-14 destructive clear of LOG sheet data rows.
'   - Interactive : MsgBox confirm -> on OK clears rows 9..lastRow,
'                   then logs INFO with the cleared count.
'   - Headless    : skip the clear, emit WARN log only (no dialog
'                   to confirm a destructive action in batch).
' Note: clsLogger.ClearLog cannot be reused here because v2.3 must
'   ensure the row count is captured and logged after the clear.
' ============================================================
Public Sub Btn_DeleteAllLog()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0288] modEntrySettings.Btn_DeleteAllLog ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_DeleteAllLog"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0289] modEntrySettings.Btn_DeleteAllLog EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim clearedCount As Long
    If lastRow >= LOG_DATA_START_ROW Then
        clearedCount = lastRow - LOG_DATA_START_ROW + 1
    Else
        clearedCount = 0
    End If

    If modCommon.IsHeadless() Then
        LogWarnSafe "Btn_DeleteAllLog", _
            "headless: clear skipped (confirm required), candidate=" & _
            clearedCount, "LOG-M14-CLEAR-GUARD"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0290] modEntrySettings.Btn_DeleteAllLog EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    If MsgBox(ChrW(&H30ED) & ChrW(&H30B0) & ChrW(&H5168) & ChrW(&H884C) & ChrW(&H3092) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H304B) & ChrW(&HFF1F) & ChrW(&H3053) & ChrW(&H306E) & ChrW(&H64CD) & ChrW(&H4F5C) & ChrW(&H306F) & ChrW(&H5143) & ChrW(&H306B) & ChrW(&H623B) & ChrW(&H305B) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002), _
              vbOKCancel + vbExclamation, ChrW(&H30ED) & ChrW(&H30B0) & ChrW(&H5168) & ChrW(&H524A) & ChrW(&H9664)) <> vbOK Then
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0291] modEntrySettings.Btn_DeleteAllLog EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    If clearedCount > 0 Then
        ws.Range(ws.Cells(LOG_DATA_START_ROW, 1), _
                 ws.Cells(lastRow, ws.Columns.Count)).ClearContents
    End If

    LogInfoSafe "Btn_DeleteAllLog", _
        "log rows cleared count=" & clearedCount, "LOG-M14-CLEAR-OK"
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0292] modEntrySettings.Btn_DeleteAllLog EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0293] modEntrySettings.Btn_DeleteAllLog EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_DeleteAllLog: " & Err.Number & " " & Err.Description
End Sub


' ============================================================
' ===== M-12 read-only format check (Btn_CheckFormat) ========
' ============================================================

' ============================================================
' Public Sub: Btn_CheckFormat
' Role: M-12 "Check format". Read-only consistency check between a
'   format definition (<format_dir>\<fmtId>.txt) and all knowledge
'   entries that use that format (<data_dir>\*.txt). Writes any
'   anomalies into the M-12 result grid (4 columns: KnowledgeNo,
'   FieldName, CurrentValue, Issue). On a clean run shows an
'   "no anomalies" marker. Does NOT touch clsFieldMigrator (archived
'   in v2.3) and does NOT mutate any files.
' ============================================================
Public Sub Btn_CheckFormat()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0294] modEntrySettings.Btn_CheckFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' [BTN-GUARD-PRELUDE-BEGIN] auto-injected by inject_btn_template.py
    Const BTN As String = "Btn_CheckFormat"
    Dim XLSM As String
    XLSM = ChrW(&H7BA1) & ChrW(&H7406)
    modBtnGuard.LogEnter BTN, XLSM
    If Not modBtnGuard.CheckPrereq(BTN, "config;format_dir", XLSM) Then
        modBtnGuard.LogExit BTN, XLSM, False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0295] modEntrySettings.Btn_CheckFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If
    ' [BTN-GUARD-PRELUDE-END]
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MIGRATE)

    Dim fmtId As String
    fmtId = Trim$(CStr(ws.Range(ADDR_M12_FORMAT_ID).Value))
    If Len(fmtId) = 0 Then
        If Not modCommon.IsHeadless() Then
            MsgBox ADDR_M12_FORMAT_ID & ChrW(12395) & ChrW(23550) & ChrW(35937) & _
                   ChrW(12501) & ChrW(12457) & ChrW(12540) & ChrW(12510) & _
                   ChrW(12483) & ChrW(12488) & ChrW(12434) & ChrW(20837) & _
                   ChrW(12428) & ChrW(12390) & ChrW(12367) & ChrW(12384) & _
                   ChrW(12373) & ChrW(12356), vbExclamation, ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H78BA) & ChrW(&H8A8D)
        End If
        LogWarnSafe "Btn_CheckFormat", _
            "no target format id at " & ADDR_M12_FORMAT_ID, _
            "LOG-M12-CHECK-NOFMT"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0296] modEntrySettings.Btn_CheckFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    Dim fields As Object
    Set fields = LoadFormatFieldsAsDict(fmtId)
    If fields Is Nothing Then
        If Not modCommon.IsHeadless() Then
            MsgBox ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H0020) & ChrW(&H0027) & fmtId & ChrW(&H0027) & ChrW(&H0020) & ChrW(&H304C) & ChrW(&H898B) & ChrW(&H3064) & ChrW(&H304B) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002), _
                   vbExclamation, ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H78BA) & ChrW(&H8A8D)
        End If
        LogWarnSafe "Btn_CheckFormat", _
            "format not found id=" & fmtId, "LOG-M12-CHECK-NOFMTDEF"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0297] modEntrySettings.Btn_CheckFormat EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    Dim knowledgeList As Collection
    Set knowledgeList = modKnowledgeFileIO.ListKnowledgesByFormat(fmtId)

    Dim issues As Collection
    Set issues = CollectFormatIssues(knowledgeList, fields)

    ClearCheckResultGrid ws
    WriteCheckResultGrid ws, issues

    LogInfoSafe "Btn_CheckFormat", _
        "check done fmtId=" & fmtId & " knowledge=" & knowledgeList.Count & _
        " issues=" & issues.Count, "LOG-M12-CHECK-OK"
    ' [BTN-GUARD-EXIT-OK] auto-injected
    modBtnGuard.LogExit BTN, XLSM, True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0298] modEntrySettings.Btn_CheckFormat EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0299] modEntrySettings.Btn_CheckFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' [BTN-GUARD-ERR-LOG] auto-injected
    modBtnGuard.LogExit BTN, XLSM, False
    Debug.Print "[ERR] Btn_CheckFormat: " & Err.Number & " " & Err.Description
End Sub

' ============================================================
' Private Function: LoadFormatFieldsAsDict
' Role: parse <format_dir>\<fmtId>.txt via modFormatLoader and
'       return a Dictionary keyed by FieldName whose value is a
'       child Dictionary with the rule keys (FieldType / Required /
'       MaxLength / DropdownOptions). Returns Nothing when the
'       format file does not parse.
' ============================================================
Private Function LoadFormatFieldsAsDict(ByVal fmtId As String) As Object
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0300] modEntrySettings.LoadFormatFieldsAsDict ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim secs As Collection
    Set secs = modFormatLoader.LoadFormat(fmtId)
    If secs Is Nothing Then
        Set LoadFormatFieldsAsDict = Nothing
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0301] modEntrySettings.LoadFormatFieldsAsDict EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To secs.Count
        Set sec = secs.Item(i)
        If sec.SectionName = "FIELD" Then
            Dim fieldName As String
            fieldName = sec.GetValue("FieldName")
            If Len(fieldName) > 0 Then
                Dim rules As Object
                Set rules = CreateObject("Scripting.Dictionary")
                rules("FieldType") = sec.GetValue("FieldType")
                rules("Required") = sec.GetValue("Required")
                rules("MaxLength") = sec.GetValue("MaxLength")
                rules("DropdownOptions") = sec.GetValue("DropdownOptions")
                Set result(fieldName) = rules
            End If
        End If
    Next i

    Set LoadFormatFieldsAsDict = result
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0302] modEntrySettings.LoadFormatFieldsAsDict EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0303] modEntrySettings.LoadFormatFieldsAsDict EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set LoadFormatFieldsAsDict = Nothing
End Function

' ============================================================
' Private Function: CollectFormatIssues
' Role: walk every knowledge entry and accumulate a Collection of
'       Variant(0..3) tuples (KnowledgeNo, FieldName, CurrentValue,
'       Issue) reporting fields whose stored value violates the
'       format rules. Read-only; the source files are not touched.
' ============================================================
Private Function CollectFormatIssues(ByVal knowledgeList As Collection, _
                                     ByVal fields As Object) As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0304] modEntrySettings.CollectFormatIssues ENTER"  ' [ADR-0100]
    Dim issues As Collection
    Set issues = New Collection
    On Error GoTo ErrHandler

    Dim kno As Variant
    For Each kno In knowledgeList
        AccumulateIssuesForOneKnowledge CStr(kno), fields, issues
    Next kno

    Set CollectFormatIssues = issues
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0305] modEntrySettings.CollectFormatIssues EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0306] modEntrySettings.CollectFormatIssues EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set CollectFormatIssues = issues
End Function

' ============================================================
' Private Sub: AccumulateIssuesForOneKnowledge
' Role: load one knowledge file and, for each format field, run the
'       value through the field rule checker; append any anomaly to
'       the issues collection.
' ============================================================
Private Sub AccumulateIssuesForOneKnowledge(ByVal knowledgeNo As String, _
                                            ByVal fields As Object, _
                                            ByVal issues As Collection)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0307] modEntrySettings.AccumulateIssuesForOneKnowledge ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim data As Object
    Set data = modKnowledgeFileIO.LoadKnowledge(knowledgeNo)
    If data Is Nothing Then Exit Sub

    Dim fieldName As Variant
    For Each fieldName In fields.Keys
        Dim rules As Object
        Set rules = fields(fieldName)
        Dim curValue As String
        If data.Exists(CStr(fieldName)) Then
            curValue = CStr(data(CStr(fieldName)))
        Else
            curValue = ""
        End If
        Dim issueText As String
        issueText = DetectFieldIssue(rules, curValue)
        If Len(issueText) > 0 Then
            issues.Add Array(knowledgeNo, CStr(fieldName), curValue, issueText)
        End If
    Next fieldName
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0308] modEntrySettings.AccumulateIssuesForOneKnowledge EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0309] modEntrySettings.AccumulateIssuesForOneKnowledge EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    ' One bad knowledge file must not abort the whole sweep.
    issues.Add Array(knowledgeNo, "", "", _
        "load failed: " & Err.Number & " " & Err.Description)
End Sub

' ============================================================
' Private Function: DetectFieldIssue
' Role: pure value-vs-rule check. Returns "" when the value is
'       acceptable; otherwise returns a short ASCII issue message
'       (one issue per call - first violation wins).
' Rules:
'   - Required=True and value empty            -> "required field is empty"
'   - FieldType=number and value not numeric   -> "value is not numeric"
'   - MaxLength set and Len(value) > MaxLength -> "value exceeds max length N"
'   - DropdownOptions set and value not listed -> "value not in dropdown options"
' ============================================================
Private Function DetectFieldIssue(ByVal rules As Object, _
                                  ByVal value As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0310] modEntrySettings.DetectFieldIssue ENTER"  ' [ADR-0100]
    Dim required As String
    Dim fieldType As String
    Dim maxLen As String
    Dim options As String
    required = CStr(rules("Required"))
    fieldType = LCase$(CStr(rules("FieldType")))
    maxLen = CStr(rules("MaxLength"))
    options = CStr(rules("DropdownOptions"))

    If LCase$(required) = "true" And Len(value) = 0 Then
        DetectFieldIssue = "required field is empty"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0311] modEntrySettings.DetectFieldIssue EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    If Len(value) > 0 And fieldType = "number" Then
        If Not IsNumeric(value) Then
            DetectFieldIssue = "value is not numeric"
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0312] modEntrySettings.DetectFieldIssue EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    End If

    If Len(maxLen) > 0 And IsNumeric(maxLen) Then
        If Len(value) > CLng(maxLen) Then
            DetectFieldIssue = "value exceeds max length " & maxLen
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0313] modEntrySettings.DetectFieldIssue EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    End If

    If Len(options) > 0 And Len(value) > 0 Then
        If Not IsValueInDropdownOptions(value, options) Then
            DetectFieldIssue = "value not in dropdown options"
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0314] modEntrySettings.DetectFieldIssue EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    End If

    DetectFieldIssue = ""
End Function

' ============================================================
' Private Function: IsValueInDropdownOptions
' Role: DropdownOptions is a "|"-separated list (the project's
'       canonical format). Returns True when value is a member.
' ============================================================
Private Function IsValueInDropdownOptions(ByVal value As String, _
                                          ByVal options As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0315] modEntrySettings.IsValueInDropdownOptions ENTER"  ' [ADR-0100]
    Dim parts() As String
    parts = Split(options, "|")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) = value Then
            IsValueInDropdownOptions = True
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0316] modEntrySettings.IsValueInDropdownOptions EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next i
    IsValueInDropdownOptions = False
End Function

' ============================================================
' Private Sub: ClearCheckResultGrid
' Role: wipe the 4 result columns over a defensive row window
'       (M12_RESULT_CLEAR_ROWS rows starting at M12_RESULT_START_ROW)
'       so a previous run never leaks into the new one.
' ============================================================
Private Sub ClearCheckResultGrid(ByVal ws As Worksheet)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0317] modEntrySettings.ClearCheckResultGrid ENTER"  ' [ADR-0100]
    Dim firstRow As Long
    Dim lastRow As Long
    firstRow = M12_RESULT_START_ROW
    lastRow = firstRow + M12_RESULT_CLEAR_ROWS - 1
    ws.Range(ws.Cells(firstRow, M12_RESULT_COL_KNO), _
             ws.Cells(lastRow, M12_RESULT_COL_ISS)).ClearContents
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0318] modEntrySettings.ClearCheckResultGrid EXIT-OK"  ' [ADR-0100]
End Sub

' ============================================================
' Private Sub: WriteCheckResultGrid
' Role: write the issues collection into the 4-column result grid.
'       Empty issues -> single-row "no anomalies" marker in the
'       Issue column (so the screen visibly reflects the run).
' ============================================================
Private Sub WriteCheckResultGrid(ByVal ws As Worksheet, _
                                 ByVal issues As Collection)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0319] modEntrySettings.WriteCheckResultGrid ENTER"  ' [ADR-0100]
    If issues.Count = 0 Then
        ws.Cells(M12_RESULT_START_ROW, M12_RESULT_COL_ISS).Value = _
            "no anomalies"
        LogInfoSafe "Btn_CheckFormat", _
            "no anomalies (empty result grid marker emitted)", _
            "LOG-M12-CHECK-EMPTY"
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0320] modEntrySettings.WriteCheckResultGrid EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    Dim r As Long
    Dim i As Long
    r = M12_RESULT_START_ROW
    For i = 1 To issues.Count
        Dim row As Variant
        row = issues.Item(i)
        ws.Cells(r, M12_RESULT_COL_KNO).Value = row(0)
        ws.Cells(r, M12_RESULT_COL_FLD).Value = row(1)
        ws.Cells(r, M12_RESULT_COL_VAL).Value = row(2)
        ws.Cells(r, M12_RESULT_COL_ISS).Value = row(3)
        r = r + 1
    Next i
End Sub


' ============================================================
' ===== common helpers =======================================
' ============================================================

' ============================================================
' Private Sub: PersistConfigKeys
' Role: shared backend for the M-10 / M-11 Save buttons.
'       writes the supplied key dictionary into config.txt via
'       modConfigLoader.SaveConfigKeys (file-side, leaves other
'       keys untouched), then syncs the in-memory holder via
'       modConfigHolder.SetConfigKeys.
'       Both calls are exception-guarded so a transient I/O glitch
'       on one path does not break the other.
' ============================================================
Private Sub PersistConfigKeys(ByVal keyValues As Object)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0321] modEntrySettings.PersistConfigKeys ENTER"  ' [ADR-0100]
    On Error Resume Next
    modConfigLoader.SaveConfigKeys XLSM_KEY, keyValues
    On Error GoTo 0

    On Error Resume Next
    modConfigHolder.SetConfigKeys keyValues
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0322] modEntrySettings.PersistConfigKeys EXIT-OK"  ' [ADR-0100]
End Sub

' ============================================================
' Private Function: SafeGetCfg
' Role: exception-safe wrapper around modConfigHolder.GetValueOrDefault.
' ============================================================
Private Function SafeGetCfg(ByVal key As String, ByVal defaultValue As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0323] modEntrySettings.SafeGetCfg ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim v As String
    v = modConfigHolder.GetValueOrDefault(key, defaultValue)
    If Len(v) = 0 Then v = defaultValue
    SafeGetCfg = v
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0324] modEntrySettings.SafeGetCfg EXIT-OK"  ' [ADR-0100]
End Function

' ============================================================
' Private Function: NewLogger
' Role: clsLogger pre-initialized on the LOG sheet. Returns
'       Nothing on failure so callers can fall through silently.
' ============================================================
Private Function NewLogger() As clsLogger
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0325] modEntrySettings.NewLogger ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ThisWorkbook.Worksheets(SHEET_LOG)
    Set NewLogger = lg
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0326] modEntrySettings.NewLogger EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0327] modEntrySettings.NewLogger EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set NewLogger = Nothing
End Function

' ============================================================
' Private Sub: LogInfoSafe / LogWarnSafe / LogErrorSafe
' Role: thin exception-safe wrappers around clsLogger, same emit
'       style as the M-10 / M-11 subs.
' ============================================================
Private Sub LogInfoSafe(ByVal funcName As String, ByVal msg As String, _
                        ByVal logId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0328] modEntrySettings.LogInfoSafe ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = NewLogger()
    If Not lg Is Nothing Then lg.LogInfo MOD_NAME, funcName, msg, _
        "", logId
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0329] modEntrySettings.LogInfoSafe EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub LogWarnSafe(ByVal funcName As String, ByVal msg As String, _
                        ByVal logId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0330] modEntrySettings.LogWarnSafe ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = NewLogger()
    If Not lg Is Nothing Then lg.LogWarn MOD_NAME, funcName, msg, _
        "", logId
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0331] modEntrySettings.LogWarnSafe EXIT-OK"  ' [ADR-0100]
End Sub

Private Sub LogErrorSafe(ByVal funcName As String, ByVal msg As String, _
                         ByVal logId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0332] modEntrySettings.LogErrorSafe ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = NewLogger()
    If Not lg Is Nothing Then lg.LogError MOD_NAME, funcName, msg, _
        "", logId
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0333] modEntrySettings.LogErrorSafe EXIT-OK"  ' [ADR-0100]
End Sub


' ============================================================
' ===== SPEC_DRIFT-4b fix (2026-05-30) =======================
' ----- ChrW()-built JP identifiers (avoid CP932/UTF-8 race) -
' ============================================================
' Background:
'   The previous Const definitions held Japanese string literals
'   that were silently corrupted to U+FFFD during a CP932->UTF-8
'   Edit/Write round-trip, which broke SaveConfigKeys (mojibake
'   _config.txt filename) and ThisWorkbook.Worksheets(...) lookups
'   for the M-10 / M-11 / M-12 sheets.
' Implementation:
'   Identifiers are now built from ChrW(&Hxxxx) so the source
'   file stays portable across encodings. Callers refer to these
'   as zero-arg parameterless functions, which VBA treats
'   syntactically the same as a Const reference.

' XLSM_KEY = "kanri" (kanri.xlsm config-file logical name)
'   U+7BA1 U+7406
Private Function XLSM_KEY() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0334] modEntrySettings.XLSM_KEY ENTER"  ' [ADR-0100]
    XLSM_KEY = ChrW(&H7BA1) & ChrW(&H7406)
End Function

' SHEET_SETTINGS = "settei" (M-11 sheet name)
'   U+8A2D U+5B9A
Private Function SHEET_SETTINGS() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0335] modEntrySettings.SHEET_SETTINGS ENTER"  ' [ADR-0100]
    SHEET_SETTINGS = ChrW(&H8A2D) & ChrW(&H5B9A)
End Function

' SHEET_STORAGE = "kakuno-saki settei" (M-10 sheet name)
'   U+683C U+7D0D U+5148 U+8A2D U+5B9A
Private Function SHEET_STORAGE() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0336] modEntrySettings.SHEET_STORAGE ENTER"  ' [ADR-0100]
    SHEET_STORAGE = ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & _
                    ChrW(&H8A2D) & ChrW(&H5B9A)
End Function

' SHEET_MIGRATE = "format henkou check" (M-12 sheet name)
'   U+30D5 U+30A9 U+30FC U+30DE U+30C3 U+30C8 U+5909 U+66F4
'   U+30C1 U+30A7 U+30C3 U+30AF
Private Function SHEET_MIGRATE() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0337] modEntrySettings.SHEET_MIGRATE ENTER"  ' [ADR-0100]
    SHEET_MIGRATE = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & _
                    ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & _
                    ChrW(&H5909) & ChrW(&H66F4) & ChrW(&H30C1) & _
                    ChrW(&H30A7) & ChrW(&H30C3) & ChrW(&H30AF)
End Function


' [USER-REQ 2026-06-09] Btn_DeleteOrphans
' Walk all 5 storage folders (data/format/ui/backup/config) and delete
' files matching "unnecessary" naming patterns:
'   __*  (double-underscore tmp files)
'   *.bak.*, *.bak_*  (backup-of-edit artifacts)
'   *.disabled, *.broken_*  (disabled files)
'   *.tmp, *.tmp.*  (temp files)
Public Sub Btn_DeleteOrphans()
    On Error GoTo ErrHandler
    Dim dirs As Collection
    Set dirs = New Collection
    dirs.Add modConfigHolder.GetDataDir()
    dirs.Add modConfigHolder.GetFormatDir()
    dirs.Add modConfigHolder.GetUiDir()
    dirs.Add modConfigHolder.GetBackupDir()
    dirs.Add modConfigHolder.GetConfigDir()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' First pass: count matching files
    Dim cnt As Long
    cnt = 0
    Dim d As Variant
    For Each d In dirs
        cnt = cnt + CountOrphansRecursive(fso, CStr(d))
    Next d

    If cnt = 0 Then
        MsgBox ChrW(&H4E0D) & ChrW(&H8981) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H306F) & ChrW(&H898B) & ChrW(&H3064) & ChrW(&H304B) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3067) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H3002), vbInformation, ChrW(&H4E0D) & ChrW(&H8981) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H524A) & ChrW(&H9664)
        Exit Sub
    End If

    Dim resp As VbMsgBoxResult
    resp = MsgBox(CStr(cnt) & ChrW(&H4EF6) & ChrW(&H306E) & ChrW(&H4E0D) & ChrW(&H8981) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H3092) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H3002) & ChrW(&H3088) & ChrW(&H308D) & ChrW(&H3057) & ChrW(&H3044) & ChrW(&H3067) & ChrW(&H3059) & ChrW(&H304B) & ChrW(&HFF1F), vbYesNo + vbExclamation, ChrW(&H4E0D) & ChrW(&H8981) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H78BA) & ChrW(&H8A8D))
    If resp <> vbYes Then Exit Sub

    Dim deleted As Long
    deleted = 0
    For Each d In dirs
        deleted = deleted + DeleteOrphansRecursive(fso, CStr(d))
    Next d

    MsgBox CStr(deleted) & ChrW(&H4EF6) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H3002), vbInformation, ChrW(&H4E0D) & ChrW(&H8981) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H5B8C) & ChrW(&H4E86)
    Exit Sub
ErrHandler:
    MsgBox "Btn_DeleteOrphans error: " & Err.Description, vbCritical, "Error"
End Sub

Private Function IsOrphanFile(ByVal fname As String) As Boolean
    Dim n As String
    n = LCase(fname)
    If Left(n, 2) = "__" Then IsOrphanFile = True: Exit Function
    If InStr(n, ".bak.") > 0 Then IsOrphanFile = True: Exit Function
    If InStr(n, ".bak_") > 0 Then IsOrphanFile = True: Exit Function
    If Right(n, 9) = ".disabled" Then IsOrphanFile = True: Exit Function
    If InStr(n, ".broken_") > 0 Then IsOrphanFile = True: Exit Function
    If Right(n, 4) = ".tmp" Then IsOrphanFile = True: Exit Function
    If InStr(n, ".tmp.") > 0 Then IsOrphanFile = True: Exit Function
End Function

Private Function CountOrphansRecursive(ByVal fso As Object, ByVal dirPath As String) As Long
    On Error Resume Next
    If Not fso.FolderExists(dirPath) Then Exit Function
    Dim folder As Object
    Set folder = fso.GetFolder(dirPath)
    Dim n As Long
    n = 0
    Dim f As Object
    For Each f In folder.Files
        If IsOrphanFile(f.Name) Then n = n + 1
    Next f
    Dim sub_ As Object
    For Each sub_ In folder.SubFolders
        n = n + CountOrphansRecursive(fso, sub_.Path)
    Next sub_
    CountOrphansRecursive = n
End Function

Private Function DeleteOrphansRecursive(ByVal fso As Object, ByVal dirPath As String) As Long
    On Error Resume Next
    If Not fso.FolderExists(dirPath) Then Exit Function
    Dim folder As Object
    Set folder = fso.GetFolder(dirPath)
    Dim n As Long
    n = 0
    Dim f As Object
    Dim toDelete As Collection
    Set toDelete = New Collection
    For Each f In folder.Files
        If IsOrphanFile(f.Name) Then toDelete.Add f.Path
    Next f
    Dim p As Variant
    For Each p In toDelete
        On Error Resume Next
        fso.DeleteFile CStr(p), True
        If Err.Number = 0 Then n = n + 1
        Err.Clear
    Next p
    Dim sub_ As Object
    For Each sub_ In folder.SubFolders
        n = n + DeleteOrphansRecursive(fso, sub_.Path)
    Next sub_
    DeleteOrphansRecursive = n
End Function

' [USER-REQ 2026-06-09] Btn_DeleteBackups
' Delete ALL files under backup_dir (knowledge backup history).
Public Sub Btn_DeleteBackups()
    On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim bdir As String
    bdir = modConfigHolder.GetBackupDir()
    If Not fso.FolderExists(bdir) Then
        MsgBox ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30EB) & ChrW(&H30C0) & ChrW(&H304C) & ChrW(&H3042) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002), vbInformation
        Exit Sub
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(bdir)
    Dim cnt As Long
    cnt = folder.Files.Count
    If cnt = 0 Then
        MsgBox ChrW(&H524A) & ChrW(&H9664) & ChrW(&H5BFE) & ChrW(&H8C61) & ChrW(&H304C) & ChrW(&H3042) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002), vbInformation, ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H524A) & ChrW(&H9664)
        Exit Sub
    End If

    Dim resp As VbMsgBoxResult
    resp = MsgBox(ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & CStr(cnt) & ChrW(&H4EF6) & ChrW(&H3092) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H3002) & ChrW(&H3088) & ChrW(&H308D) & ChrW(&H3057) & ChrW(&H3044) & ChrW(&H3067) & ChrW(&H3059) & ChrW(&H304B) & ChrW(&HFF1F), vbYesNo + vbExclamation, ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H78BA) & ChrW(&H8A8D))
    If resp <> vbYes Then Exit Sub

    Dim deleted As Long
    deleted = 0
    Dim f As Object
    Dim toDelete As Collection
    Set toDelete = New Collection
    For Each f In folder.Files
        toDelete.Add f.Path
    Next f
    Dim p As Variant
    For Each p In toDelete
        On Error Resume Next
        fso.DeleteFile CStr(p), True
        If Err.Number = 0 Then deleted = deleted + 1
        Err.Clear
    Next p

    MsgBox CStr(deleted) & ChrW(&H4EF6) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3057) & ChrW(&H305F) & ChrW(&H3002), vbInformation, ChrW(&H30D0) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30A2) & ChrW(&H30C3) & ChrW(&H30D7) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H5B8C) & ChrW(&H4E86)
    Exit Sub
ErrHandler:
    MsgBox "Btn_DeleteBackups error: " & Err.Description, vbCritical, "Error"
End Sub
```
