---
title: modConfigLoader.bas
description: modConfigLoader.bas のソースコード（コピペ用）
---

# modConfigLoader.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-30 14:44 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modConfigLoader.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modConfigLoader"
' ================================================================
' Module:  modConfigLoader (Fix-6 / ADR-0133)
' Summary: Reads the 9-cell worksheet SSOT (the "settings sheet")
'          at startup and feeds modConfigHolder via SetConfig.
'          Replaces the retired text-file loader. ADR-0133
'          supersedes ADR-0132 (install-injected path constants).
'          Layout: column B = label, column D = value (ui_seed
'          input convention), rows 11..17 (4 paths + 3 behaviour).
'          First run (all 4 path cells empty): InputBox for BASE_DIR
'          interactive / default base headless, derive 6 dirs, seed
'          3 behaviour defaults, write back, save the workbook.
'          Partial-missing path: explicit error + abort (interactive)
'          or default-fill + warn (headless).
' Version: v2.4 Fix-6 (2026-06-18)
' Depends: modConfigHolder (holder), modCommon (IsHeadless / gDebugLevel)
' Related: ADR-0133 / ADR-0131 / ADR-0094 / ADR-0053 section 2.9
' ================================================================
Option Explicit

Private Const DEFAULT_BASE As String = "C:\KnowledgeMgr"
Private Const DEFAULT_DEBUG_LEVEL As String = "INFO"
Private Const DEFAULT_LOG_ROTATION_ROWS As String = "10000"
Private Const DEFAULT_UI_SCHEMA_FAIL_MODE As String = "safeDefault"

Private Const FIRST_DATA_ROW As Long = 11    ' worksheet row of first key (data_dir)
Private Const KEY_COL As Long = 2            ' column B = key label
Private Const VALUE_COL As Long = 4          ' column D = value (ui_seed input cell)
Private Const PATH_KEY_COUNT As Long = 4     ' first 4 keys are paths

' ----------------------------------------------------------------
' Public API
' ----------------------------------------------------------------

' Display name of the settings sheet ("kakuno-saki settei"). JP via ChrW.
Public Function SettingsSheetName() As String
    SettingsSheetName = ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & ChrW(&H8A2D) & ChrW(&H5B9A)
End Function

' Compat shim (ADR-0131 used this to name the unreadable config file in a
' MsgBox). With Fix-6 the source is a worksheet, so return its display name.
Public Function LastConfigPath() As String
    LastConfigPath = SettingsSheetName()
End Function

' ================================================================
' Function: LoadConfig
' Summary:  Orchestrator step-1 entry. Ensures + seeds the settings
'           sheet, reads the 9 cells, handles empty/partial, feeds
'           the holder. Returns False only on interactive
'           partial-missing (caller aborts the remaining setup).
' ================================================================
Public Function LoadConfig(ByVal xlsmName As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0960] modConfigLoader.LoadConfig ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = EnsureSettingsSheet()
    If ws Is Nothing Then
        LoadConfig = False
        Exit Function
    End If

    ' [M10-MIGRATE] in-place upgrade: drop old log_dir/config_dir rows and
    ' shift the survivors to the current 7-key layout before reading.
    MigrateOldStorageLayout ws
    ' [M10-MIGRATE] in-place upgrade: point ui_dir at the role subfolder when
    ' the configured path is the bare base (pre-C1) so the renderer resolves.
    EnsureUiDirRoleResolvable ws

    Dim d As Object
    Set d = ReadSheetToDict(ws)

    Dim emptyPaths As Long
    emptyPaths = CountEmptyPaths(d)

    If emptyPaths = PATH_KEY_COUNT Then
        ' First run: no path is set yet.
        DeriveDirs AskBaseDir(), d
        ApplyBehaviorDefaults d
        WriteDictToSheet ws, d
        SaveBookQuiet
        modConfigHolder.SetConfig d
        LoadConfig = True
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0961] modConfigLoader.LoadConfig EXIT-OK first-run"  ' [ADR-0100]
        Exit Function
    End If

    If emptyPaths > 0 Then
        ' Partial-missing path cells.
        ' [Fix-6 / ADR-0133-followup 2026-06-19] Auto-heal instead of the old
        ' interactive hard-abort. The common real case is a config migrated
        ' from a pre-Fix-6 config that had no base-dir line: the sub-dir
        ' cells are filled but the base cell is empty, so emptyPaths is partial
        ' and the first-run InputBox never fires. The user then sees an opaque
        ' "base dir is empty or invalid" abort with no idea what to enter.
        ' Instead, infer BASE_DIR from an existing path and derive only the
        ' missing cells, then continue. Only when no base can be inferred do we
        ' fall back to the previous default-fill (headless) / notify (interactive).
        Dim base As String
        base = InferBaseDir(d)
        If Len(base) > 0 Then
            FillMissingPathsFromBase d, base
            ApplyBehaviorDefaults d
            WriteDictToSheet ws, d
            SaveBookQuiet
            modConfigHolder.SetConfig d
            LoadConfig = True
            If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0962b] modConfigLoader.LoadConfig EXIT-OK partial-autoheal base=" & base  ' [ADR-0100]
            Exit Function
        End If
        If modCommon.IsHeadless() Then
            FillMissingPathsFromDefault d
            ApplyBehaviorDefaults d
            WriteDictToSheet ws, d
            SaveBookQuiet
            modConfigHolder.SetConfig d
            LoadConfig = True
            Exit Function
        Else
            NotifyMissingKey FirstEmptyPathKey(d)
            LoadConfig = False
            If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0962] modConfigLoader.LoadConfig EXIT abort partial-missing"  ' [ADR-0100]
            Exit Function
        End If
    End If

    ' All path cells present.
    ApplyBehaviorDefaults d
    modConfigHolder.SetConfig d
    LoadConfig = True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0963] modConfigLoader.LoadConfig EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0964] modConfigLoader.LoadConfig EXIT-ERR errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    LoadConfig = False
End Function

' [M10-MIGRATE] In-place upgrade cleanup. The pre-2026-06 settings layout
' had 8 data rows incl. log_dir (D14) and config_dir (D16); the current
' 7-key layout dropped both, so an upgraded book shows the still-valid
' values two rows too low (backup in the debugLevel row, etc.). Detect by
' the new debugLevel cell (D15) holding a path (the old backup_dir) and
' shift the survivors up; clear the old logRotationRows straggler in D18.
' Idempotent: after migration D15 is a debug level (no path) so it never
' re-triggers - no separate flag needed.
Private Sub MigrateOldStorageLayout(ByVal ws As Object)
    On Error Resume Next
    Dim d15 As String
    d15 = CStr(ws.Cells(15, VALUE_COL).Value)
    If InStr(d15, "\") = 0 And InStr(d15, ":") = 0 Then Exit Sub  ' already new layout
    ' capture still-valid old values (old D15=backup, D17=debugLevel, D18=logRotationRows)
    Dim oldBackup As String, oldDebug As String, oldLogRot As String
    oldBackup = CStr(ws.Cells(15, VALUE_COL).Value)
    oldDebug = CStr(ws.Cells(17, VALUE_COL).Value)
    oldLogRot = CStr(ws.Cells(18, VALUE_COL).Value)
    Dim wasP As Boolean
    wasP = ws.ProtectContents
    If wasP Then ws.Unprotect Password:=""
    ' new layout: D14=backup_dir, D15=debugLevel, D16=logRotationRows, D17=uiSchemaFailMode
    ws.Cells(14, VALUE_COL).Value = oldBackup
    ws.Cells(15, VALUE_COL).Value = oldDebug
    ws.Cells(16, VALUE_COL).Value = oldLogRot
    ws.Cells(17, VALUE_COL).Value = "safeDefault"
    ws.Cells(18, VALUE_COL).ClearContents
    If wasP Then ws.Protect Password:="", UserInterfaceOnly:=True
    modCommon.AppendProgressLog "[M10-MIGRATE] old log_dir/config_dir layout migrated to 7-key layout"
End Sub

' [M10-MIGRATE] In-place upgrade: a pre-C1 book has ui_dir = the bare base
' (...\ui\) but screen files now live under ...\ui\<role>\. C1 made consumers
' read GetUiDir() verbatim, so the renderer would miss every file. When D13 is
' not already role-qualified and a <path>\<role>\ subfolder exists, point D13
' there. Idempotent; a genuine custom path (no such subfolder) is left as-is.
Private Sub EnsureUiDirRoleResolvable(ByVal ws As Object)
    On Error Resume Next
    Dim role As String
    role = Replace$(ThisWorkbook.Name, ".xlsm", "")
    If Len(role) = 0 Then Exit Sub
    Dim base As String
    base = CStr(ws.Cells(13, VALUE_COL).Value)
    Do While Len(base) > 0 And Right$(base, 1) = "\"
        base = Left$(base, Len(base) - 1)
    Loop
    If Len(base) = 0 Then Exit Sub
    If Len(base) > Len(role) Then
        If Right$(base, Len(role) + 1) = "\" & role Then Exit Sub  ' already role-qualified
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(base & "\" & role) Then Exit Sub
    Dim wasP As Boolean
    wasP = ws.ProtectContents
    If wasP Then ws.Unprotect Password:=""
    ws.Cells(13, VALUE_COL).Value = base & "\" & role & "\"
    If wasP Then ws.Protect Password:="", UserInterfaceOnly:=True
    modCommon.AppendProgressLog "[M10-MIGRATE] ui_dir upgraded to role-qualified: " & base & "\" & role & "\"
End Sub

' ================================================================
' Function: Reload
' Summary:  Re-read the 9 cells into the holder (clsLogger polling /
'           dynamic debugLevel). No seeding, no abort.
' ================================================================
Public Function Reload() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0965] modConfigLoader.Reload ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = GetSettingsSheet()
    If ws Is Nothing Then
        Reload = False
        Exit Function
    End If
    Dim d As Object
    Set d = ReadSheetToDict(ws)
    ApplyBehaviorDefaults d
    modConfigHolder.SetConfig d
    Reload = True
    Exit Function
ErrHandler:
    Reload = False
End Function

' ================================================================
' Function: EnsureSettingsSheet
' Summary:  Return the settings sheet, creating + seeding the frame
'           (labels + validation) when it does not exist yet. Public
'           so M-10 / the orchestrator can guarantee it exists.
' ================================================================
Public Function EnsureSettingsSheet() As Worksheet
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0966] modConfigLoader.EnsureSettingsSheet ENTER"  ' [ADR-0100]
    Dim ws As Worksheet
    Set ws = GetSettingsSheet()
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Not ws Is Nothing Then ws.Name = SettingsSheetName()
        On Error GoTo 0
        If Not ws Is Nothing Then SeedSheetFrame ws
    End If
    Set EnsureSettingsSheet = ws
End Function

' ================================================================
' Sub: RestoreSheetFromHolder
' Summary:  Post-render hook. ApplyUiStanzas does Cells.Clear on the
'           admin M-10 sheet (== settings sheet); after the re-render
'           we write the holder values back into B2..B10 so the SSOT
'           round-trips across every Workbook_Open. No-op when the
'           holder is empty or the sheet is absent.
' ================================================================
Public Sub RestoreSheetFromHolder()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0967] modConfigLoader.RestoreSheetFromHolder ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If Not modConfigHolder.IsInitialized() Then Exit Sub
    Dim ws As Worksheet
    Set ws = GetSettingsSheet()
    If ws Is Nothing Then Exit Sub
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim keys As Variant
    keys = AllKeys()
    Dim i As Long
    For i = 0 To UBound(keys)
        d(keys(i)) = modConfigHolder.GetValue(CStr(keys(i)))
    Next i
    WriteDictToSheet ws, d
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0968] modConfigLoader.RestoreSheetFromHolder EXIT-ERR errNum=" & Err.Number  ' [ADR-0100]
End Sub

' ================================================================
' Sub: WriteSettingsFromDict
' Summary:  Persist edited keys (M-10 save). Merge keyValues over the
'           current sheet values, write B2..B10, save the workbook,
'           and refresh the holder. Replaces the retired text-file
'           writeback pair.
' ================================================================
Public Sub WriteSettingsFromDict(ByVal keyValues As Object)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0969] modConfigLoader.WriteSettingsFromDict ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If keyValues Is Nothing Then Exit Sub
    Dim ws As Worksheet
    Set ws = EnsureSettingsSheet()
    If ws Is Nothing Then Exit Sub
    Dim d As Object
    Set d = ReadSheetToDict(ws)
    Dim k As Variant
    For Each k In keyValues.Keys
        d(CStr(k)) = CStr(keyValues(k))
    Next k
    WriteDictToSheet ws, d
    SaveBookQuiet
    modConfigHolder.SetConfig d
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0970] modConfigLoader.WriteSettingsFromDict EXIT-ERR errNum=" & Err.Number  ' [ADR-0100]
End Sub

' ================================================================
' Function: SnapshotSettings
' Summary:  Fix-6 followup. Capture the 9 settings cells from the
'           settings sheet into a fresh Dictionary so a re-render
'           (Cells.Clear) cannot lose user-entered values. Never
'           fails: a missing sheet yields empty-string values.
' ================================================================
Public Function SnapshotSettings() As Object
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = GetSettingsSheet()
    If ws Is Nothing Then
        Set SnapshotSettings = NewEmptyKeyDict()
        Exit Function
    End If
    Set SnapshotSettings = ReadSheetToDict(ws)
    Exit Function
ErrHandler:
    Set SnapshotSettings = NewEmptyKeyDict()
End Function

' ================================================================
' Sub: RestoreSettings
' Summary:  Fix-6 followup. Write a snapshot Dictionary back into the
'           9 settings cells after a re-render. Creates the settings
'           sheet when absent. Never raises.
' ================================================================
Public Sub RestoreSettings(ByVal d As Object)
    On Error GoTo ErrHandler
    If d Is Nothing Then Exit Sub
    Dim ws As Worksheet
    Set ws = EnsureSettingsSheet()
    If ws Is Nothing Then Exit Sub
    WriteDictToSheet ws, d
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0971] modConfigLoader.RestoreSettings EXIT-ERR errNum=" & Err.Number  ' [ADR-0100]
End Sub

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------

Private Function AllKeys() As Variant
    AllKeys = Array("data_dir", "format_dir", "ui_dir", "backup_dir", "debugLevel", "logRotationRows", "uiSchemaFailMode")
End Function

' Fix-6 followup: a Dictionary pre-populated with all 9 keys set to "".
Private Function NewEmptyKeyDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim keys As Variant
    keys = AllKeys()
    Dim i As Long
    For i = 0 To UBound(keys)
        d(keys(i)) = ""
    Next i
    Set NewEmptyKeyDict = d
End Function

Private Function GetSettingsSheet() As Worksheet
    On Error Resume Next
    Set GetSettingsSheet = ThisWorkbook.Worksheets(SettingsSheetName())
    On Error GoTo 0
End Function

Private Function ReadSheetToDict(ByVal ws As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim keys As Variant
    keys = AllKeys()
    Dim i As Long
    For i = 0 To UBound(keys)
        d(keys(i)) = Trim(CStr(ws.Cells(FIRST_DATA_ROW + i, VALUE_COL).Value))
    Next i
    Set ReadSheetToDict = d
End Function

Private Sub WriteDictToSheet(ByVal ws As Object, ByVal d As Object)
    Dim wasP As Boolean
    On Error Resume Next
    wasP = ws.ProtectContents
    If wasP Then ws.Unprotect Password:=""
    On Error GoTo 0
    Dim keys As Variant
    keys = AllKeys()
    Dim i As Long
    For i = 0 To UBound(keys)
        ws.Cells(FIRST_DATA_ROW + i, VALUE_COL).Value = CStr(d(keys(i)))
    Next i
    On Error Resume Next
    If wasP Then ws.Protect Password:="", UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Sub SeedSheetFrame(ByVal ws As Object)
    Dim wasP As Boolean
    On Error Resume Next
    wasP = ws.ProtectContents
    If wasP Then ws.Unprotect Password:=""
    On Error GoTo 0
    ws.Cells(1, KEY_COL).Value = SettingsSheetName()
    Dim keys As Variant
    keys = AllKeys()
    Dim i As Long
    For i = 0 To UBound(keys)
        ws.Cells(FIRST_DATA_ROW + i, KEY_COL).Value = keys(i)
    Next i
    AddListValidation ws.Cells(FIRST_DATA_ROW + 4, VALUE_COL), "DEBUG,INFO,WARN,ERROR,OFF,TRACE"
    AddListValidation ws.Cells(FIRST_DATA_ROW + 6, VALUE_COL), "safeDefault,warn,abort"
    On Error Resume Next
    ws.Columns(KEY_COL).ColumnWidth = 18
    ws.Columns(VALUE_COL).ColumnWidth = 34
    If wasP Then ws.Protect Password:="", UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Sub AddListValidation(ByVal rng As Object, ByVal listCsv As String)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=listCsv
    On Error GoTo 0
End Sub

Private Function CountEmptyPaths(ByVal d As Object) As Long
    Dim keys As Variant
    keys = AllKeys()
    Dim i As Long
    Dim c As Long
    For i = 0 To PATH_KEY_COUNT - 1
        If Len(Trim(CStr(d(keys(i))))) = 0 Then c = c + 1
    Next i
    CountEmptyPaths = c
End Function

Private Function FirstEmptyPathKey(ByVal d As Object) As String
    Dim keys As Variant
    keys = AllKeys()
    Dim i As Long
    For i = 0 To PATH_KEY_COUNT - 1
        If Len(Trim(CStr(d(keys(i))))) = 0 Then
            FirstEmptyPathKey = CStr(keys(i))
            Exit Function
        End If
    Next i
    FirstEmptyPathKey = ""
End Function

Private Sub DeriveDirs(ByVal base As String, ByVal d As Object)
    Dim b As String
    b = base
    Do While Right(b, 1) = "\"
        b = Left(b, Len(b) - 1)
    Loop
    d("data_dir") = b & "\data\"
    d("format_dir") = b & "\formats\"
    ' [C1] ui_dir installer default is role-qualified (<base>\ui\<role>\),
    ' role from ThisWorkbook (ADR-0090). Seeded on first run only; the cell is
    ' then a free absolute path the user may change to anything.
    Dim uiRole As String
    uiRole = Replace$(ThisWorkbook.Name, ".xlsm", "")
    If Len(uiRole) > 0 Then
        d("ui_dir") = b & "\ui\" & uiRole & "\"
    Else
        d("ui_dir") = b & "\ui\"
    End If
    d("backup_dir") = b & "\backup\"
End Sub

Private Sub FillMissingPathsFromDefault(ByVal d As Object)
    Dim def As Object
    Set def = CreateObject("Scripting.Dictionary")
    DeriveDirs DEFAULT_BASE, def
    Dim keys As Variant
    keys = AllKeys()
    Dim i As Long
    For i = 0 To PATH_KEY_COUNT - 1
        If Len(Trim(CStr(d(keys(i))))) = 0 Then d(keys(i)) = def(keys(i))
    Next i
End Sub

' [Fix-6 / ADR-0133-followup] Infer BASE_DIR from whatever path cells are
' already filled, so a partially-migrated config can self-heal without an
' InputBox. The parent of the first filled sub-path is used as the base.
' Returns "" when nothing can be inferred.
Private Function InferBaseDir(ByVal d As Object) As String
    Dim subKeys As Variant
    subKeys = Array("data_dir", "format_dir", "ui_dir", "backup_dir")
    Dim i As Long
    Dim v As String
    For i = 0 To UBound(subKeys)
        v = Trim(CStr(d(subKeys(i))))
        If Len(v) > 0 Then
            InferBaseDir = ParentDir(v)
            If Len(InferBaseDir) > 0 Then Exit Function
        End If
    Next i
    InferBaseDir = ""
End Function

' Derive every still-empty path cell from an inferred base (same layout as
' DeriveDirs). Filled cells are left untouched so user customisations survive.
Private Sub FillMissingPathsFromBase(ByVal d As Object, ByVal base As String)
    Dim derived As Object
    Set derived = CreateObject("Scripting.Dictionary")
    DeriveDirs base, derived
    Dim keys As Variant
    keys = AllKeys()
    Dim i As Long
    For i = 0 To PATH_KEY_COUNT - 1
        If Len(Trim(CStr(d(keys(i))))) = 0 Then d(keys(i)) = derived(keys(i))
    Next i
End Sub

' Strip every trailing path separator ("C:\KnowledgeMgr\" -> "C:\KnowledgeMgr").
Private Function StripTrailingSep(ByVal p As String) As String
    Dim s As String
    s = p
    Do While Len(s) > 0 And Right(s, 1) = "\"
        s = Left(s, Len(s) - 1)
    Loop
    StripTrailingSep = s
End Function

' Parent directory of a path ("C:\KnowledgeMgr\data\" -> "C:\KnowledgeMgr").
Private Function ParentDir(ByVal p As String) As String
    Dim s As String
    s = StripTrailingSep(p)
    Dim pos As Long
    pos = InStrRev(s, "\")
    If pos > 1 Then
        ParentDir = Left(s, pos - 1)
    Else
        ParentDir = ""
    End If
End Function

Private Sub ApplyBehaviorDefaults(ByVal d As Object)
    If Len(Trim(CStr(d("debugLevel")))) = 0 Then d("debugLevel") = DEFAULT_DEBUG_LEVEL
    If Len(Trim(CStr(d("logRotationRows")))) = 0 Then d("logRotationRows") = DEFAULT_LOG_ROTATION_ROWS
    If Len(Trim(CStr(d("uiSchemaFailMode")))) = 0 Then d("uiSchemaFailMode") = DEFAULT_UI_SCHEMA_FAIL_MODE
End Sub

Private Function AskBaseDir() As String
    If modCommon.IsHeadless() Then
        AskBaseDir = DEFAULT_BASE
        Exit Function
    End If
    Dim r As String
    r = InputBox(InputPrompt(), SettingsSheetName(), DEFAULT_BASE)
    If Len(Trim(r)) = 0 Then
        AskBaseDir = DEFAULT_BASE
    Else
        AskBaseDir = r
    End If
End Function

Private Sub NotifyMissingKey(ByVal key As String)
    If modCommon.IsHeadless() Then Exit Sub
    MsgBox MsgMissingA() & key & MsgMissingB(), vbExclamation, MsgErrTitle()
End Sub

Private Sub SaveBookQuiet()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

' ---- JP user-facing strings (ChrW so the source is installer-safe) ----

Private Function InputPrompt() As String
    InputPrompt = ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & ChrW(&H306E) & ChrW(&H30D9) & ChrW(&H30FC) & ChrW(&H30B9) & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30EB) & ChrW(&H30C0) & ChrW(&H3092) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & vbCrLf & DEFAULT_BASE
End Function

Private Function MsgMissingA() As String
    MsgMissingA = ChrW(&H8A2D) & ChrW(&H5B9A) & ChrW(&H30B7) & ChrW(&H30FC) & ChrW(&H30C8) & ChrW(&H306E) & ChrW(&H20)
End Function

Private Function MsgMissingB() As String
    MsgMissingB = ChrW(&H20) & ChrW(&H304C) & ChrW(&H7A7A) & ChrW(&H307E) & ChrW(&H305F) & ChrW(&H306F) & ChrW(&H4E0D) & ChrW(&H6B63) & ChrW(&H3067) & ChrW(&H3059) & ChrW(&H3002) & ChrW(&H8A2D) & ChrW(&H5B9A) & ChrW(&H30B7) & ChrW(&H30FC) & ChrW(&H30C8) & ChrW(&H3067) & ChrW(&H6307) & ChrW(&H5B9A) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
End Function

Private Function MsgErrTitle() As String
    MsgErrTitle = ChrW(&H8A2D) & ChrW(&H5B9A) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
End Function
```
