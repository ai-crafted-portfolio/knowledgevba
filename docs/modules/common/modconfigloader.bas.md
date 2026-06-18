---
title: modConfigLoader.bas
description: modConfigLoader.bas のソースコード（コピペ用）
---

# modConfigLoader.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-18 18:14 JST

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
'          input convention), rows 11..19 (6 paths + 3 behaviour).
'          First run (all 6 path cells empty): InputBox for BASE_DIR
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
Private Const PATH_KEY_COUNT As Long = 6     ' first 6 keys are paths

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

' ----------------------------------------------------------------
' Private helpers
' ----------------------------------------------------------------

Private Function AllKeys() As Variant
    AllKeys = Array("data_dir", "format_dir", "ui_dir", "log_dir", "backup_dir", "config_dir", "debugLevel", "logRotationRows", "uiSchemaFailMode")
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
    AddListValidation ws.Cells(FIRST_DATA_ROW + 6, VALUE_COL), "DEBUG,INFO,WARN,ERROR,OFF,TRACE"
    AddListValidation ws.Cells(FIRST_DATA_ROW + 8, VALUE_COL), "safeDefault,warn,abort"
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
    d("ui_dir") = b & "\ui\"
    d("log_dir") = b & "\log\"
    d("backup_dir") = b & "\backup\"
    d("config_dir") = b & "\"
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
