---
title: modConfigHolder.bas
description: modConfigHolder.bas のソースコード（コピペ用）
---

# modConfigHolder.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-30 14:44 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modConfigHolder.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modConfigHolder"
' ================================================================
' Module:  modConfigHolder (Phase 2 task 2.3 / utility layer)
' Summary: In-memory holder for config.txt values.
'          ADR-0053 section 7.3 N6 NM-2 pattern: modConfigLoader
'          reads config and pushes it here via SetConfig; the rest
'          of the codebase reads via GetValue.
' Version: v2.1 (2026-05-16 EOD, Q1-Q57 resolved)
'          v2.3 Phase A: added SetConfigKeys (partial merge API)
'          to mirror modConfigLoader.SaveConfigKeys behavior in
'          memory.
' Depends: modCommon / modFileIO. Fix-6 ADR-0133: path defaults removed; the 7-cell
' Related: ADR-0053 section 5.2 / 7.3 N6
'          Q7  (debugLevel default ERROR)
'          Q17 (GetDebugLevel As Long, enum return)
'          Fix-6 ADR-0133: getters return the worksheet SSOT value
' Note:    Japanese block comments were lost in an earlier accident
'          and restored here in ASCII to avoid encoding fragility.
'          Logic unchanged; original public API surface preserved.
' ================================================================
Option Explicit

' v2.1: debugLevel enum (6 levels). Mirrors the constants in modCommon
'       but redeclared here to avoid a holder->util uplink.
Private Const DEBUG_LEVEL_OFF   As Long = 0
Private Const DEBUG_LEVEL_ERROR As Long = 1
Private Const DEBUG_LEVEL_WARN  As Long = 2
Private Const DEBUG_LEVEL_INFO  As Long = 3
Private Const DEBUG_LEVEL_DEBUG As Long = 4
Private Const DEBUG_LEVEL_TRACE As Long = 5

Private m_config As Object              ' Scripting.Dictionary
Private m_isInitialized As Boolean
Private m_lastMissingUiFile As String   ' C3: last missing required UI def file

' ----------------------------------------------------------------
' Public I/F
' ----------------------------------------------------------------

' ================================================================
' Function: SetConfig
' Summary: Full replace: accept the whole config dict from
'          modConfigLoader and store it. Existing keys (if any)
'          are dropped.
' Param:   ByVal configDict As Object - Scripting.Dictionary (key=>value)
' ================================================================
Public Sub SetConfig(ByVal configDict As Object)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0940] modConfigHolder.SetConfig ENTER"  ' [ADR-0100]
    Set m_config = configDict
    m_isInitialized = True
End Sub

' ================================================================
' Function: GetValue
' Summary: Get value for key (in-memory lookup).
' Param:   ByVal key As String - config key
' Return:  String - value, or vbNullString if missing/uninitialized
' ================================================================
Public Function GetValue(ByVal key As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0941] modConfigHolder.GetValue ENTER"  ' [ADR-0100]
    If Not m_isInitialized Then
        GetValue = vbNullString
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0942] modConfigHolder.GetValue EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If m_config Is Nothing Then
        GetValue = vbNullString
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0943] modConfigHolder.GetValue EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If m_config.Exists(key) Then
        GetValue = CStr(m_config(key))
    Else
        GetValue = vbNullString
    End If
End Function

' ================================================================
' Function: GetValueOrDefault
' Summary: GetValue with default-value fallback when missing/empty.
' ================================================================
Public Function GetValueOrDefault(ByVal key As String, ByVal defaultValue As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0944] modConfigHolder.GetValueOrDefault ENTER"  ' [ADR-0100]
    Dim v As String
    v = GetValue(key)
    If Len(v) = 0 Then
        GetValueOrDefault = defaultValue
    Else
        GetValueOrDefault = v
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0945] modConfigHolder.GetValueOrDefault EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' Function: HasKey
' Summary: Whether key is defined in m_config.
' ================================================================
Public Function HasKey(ByVal key As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0946] modConfigHolder.HasKey ENTER"  ' [ADR-0100]
    If Not m_isInitialized Or m_config Is Nothing Then
        HasKey = False
    Else
        HasKey = m_config.Exists(key)
    End If
End Function

' ================================================================
' Function: IsInitialized
' Summary: True iff SetConfig (or SetConfigKeys) has been called.
' ================================================================
Public Function IsInitialized() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0947] modConfigHolder.IsInitialized ENTER"  ' [ADR-0100]
    IsInitialized = m_isInitialized
End Function

' ================================================================
' Function: GetDebugLevel (Q17 fixed return Long, enum value)
' Summary: Returns debugLevel as Long enum value.
'          Default is ERROR (Q7).
' Return:  Long - DEBUG_LEVEL_OFF/ERROR/WARN/INFO/DEBUG/TRACE
' ================================================================
Public Function GetDebugLevel() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0948] modConfigHolder.GetDebugLevel ENTER"  ' [ADR-0100]
    Dim s As String
    s = UCase(GetValueOrDefault("debugLevel", "ERROR"))
    Select Case s
        Case "OFF":   GetDebugLevel = DEBUG_LEVEL_OFF
        Case "ERROR": GetDebugLevel = DEBUG_LEVEL_ERROR
        Case "WARN":  GetDebugLevel = DEBUG_LEVEL_WARN
        Case "INFO":  GetDebugLevel = DEBUG_LEVEL_INFO
        Case "DEBUG": GetDebugLevel = DEBUG_LEVEL_DEBUG
        Case "TRACE": GetDebugLevel = DEBUG_LEVEL_TRACE
        Case Else:    GetDebugLevel = DEBUG_LEVEL_ERROR
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0949] modConfigHolder.GetDebugLevel EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' Wrapper accessors for common dir keys
' ================================================================
Public Function GetDataDir() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0950] modConfigHolder.GetDataDir ENTER"  ' [ADR-0100]
    GetDataDir = GetValueOrDefault("data_dir", vbNullString)
End Function

Public Function GetBackupDir() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0951] modConfigHolder.GetBackupDir ENTER"  ' [ADR-0100]
    GetBackupDir = GetValueOrDefault("backup_dir", vbNullString)
End Function

Public Function GetFormatDir() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0952] modConfigHolder.GetFormatDir ENTER"  ' [ADR-0100]
    GetFormatDir = GetValueOrDefault("format_dir", vbNullString)
End Function

Public Function GetUiDir() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0953] modConfigHolder.GetUiDir ENTER"  ' [ADR-0100]
    GetUiDir = GetValueOrDefault("ui_dir", vbNullString)
End Function

Public Function GetDebugLevelStr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0956] modConfigHolder.GetDebugLevelStr ENTER"  ' [ADR-0100]
    GetDebugLevelStr = GetValueOrDefault("debugLevel", "ERROR")
End Function

' ================================================================
' Change-3: required-UI-file fallback (ADR-0134, revised 2026-06-23)
' On Workbook_Open and on re-render, verify every screen the role
' renders has its UI definition file under ui_dir\<role>\. The set of
' screen-id list comes from the per-role SSOT (clsSetupOrchestrator
' SHEETS_*), not a hard-coded list (ADR-0090). System ids (e.g. LOG)
' have no UI file and are skipped. If any file is missing,
' callers skip the re-render and keep the prior screen. Only UI
' definition files are in scope; format/data/backup are not.
' EnsureStorageAccessible is a compat shim: returns "" when all
' required UI files are present, else the missing file name.
' ================================================================
Public Function EnsureStorageAccessible(ByVal xlsmName As String) As String
    On Error GoTo EH
    If RequiredUiFilesAvailable(xlsmName) Then
        EnsureStorageAccessible = ""
    Else
        EnsureStorageAccessible = m_lastMissingUiFile
    End If
    Exit Function
EH:
    EnsureStorageAccessible = ""
End Function

' C3 core: True iff every screen of roleName has its UI definition file
' (ui_dir\<role>\<screenId>.txt). On False the missing file name is
' exposed via LastMissingUiFile. The screen-id list comes from the
' per-role SSOT (clsSetupOrchestrator.ScreenIdsCsv over SHEETS_*), not a
' hard-coded list (ADR-0090); ids without a UI file (e.g. LOG) are skipped,
' and an intentionally disabled file (<id>.txt.disabled) counts as present.
Public Function RequiredUiFilesAvailable(ByVal roleName As String) As Boolean
    On Error GoTo EH
    m_lastMissingUiFile = ""
    Dim uiBase As String
    uiBase = GetUiDir()
    ' Fail-open when ui_dir is not loaded yet (e.g. Workbook_Open calls this
    ' before LoadConfig): an empty path must NOT be reported as files-missing,
    ' else the C3 MsgBox fires on every interactive open and setup is skipped.
    If Len(uiBase) = 0 Then RequiredUiFilesAvailable = True: Exit Function
    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    Dim ids() As String
    ids = Split(orch.ScreenIdsCsv(roleName), "|")
    Dim i As Long
    Dim sid As String
    For i = LBound(ids) To UBound(ids)
        sid = Trim$(ids(i))
        If IsScreenSheet(sid) Then
            ' present = active file OR intentionally disabled file
            Dim okF As Boolean
            okF = modFileIO.FileExists(uiBase & sid & ".txt")
            If Not okF Then okF = modFileIO.FileExists(uiBase & sid & ".txt.disabled")
            If Not okF Then
                m_lastMissingUiFile = sid & ".txt"
                RequiredUiFilesAvailable = False
                Exit Function
            End If
        End If
    Next i
    RequiredUiFilesAvailable = True
    Exit Function
EH:
    ' fail-open: never block startup / re-render on a checker error
    RequiredUiFilesAvailable = True
End Function

' Missing UI definition file from the last RequiredUiFilesAvailable call
' (empty when none was missing).
Public Function LastMissingUiFile() As String
    LastMissingUiFile = m_lastMissingUiFile
End Function

' A screen id has a UI definition file; system ids (LOG, etc.) do not.
' Screen ids follow the "M-<digits>" naming convention. This is a
' convention, not a hard-coded screen list (ADR-0090): the id set comes
' from the orchestrator SSOT and is filtered here by shape.
Private Function IsScreenSheet(ByVal nm As String) As Boolean
    On Error GoTo EH
    If Len(nm) < 3 Then Exit Function
    If Left$(nm, 2) <> "M-" Then Exit Function
    Dim i As Long
    Dim ch As String
    For i = 3 To Len(nm)
        ch = Mid$(nm, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsScreenSheet = True
    Exit Function
EH:
    IsScreenSheet = False
End Function

' Notify that a required UI definition file is missing and the previous
' screen is being kept. Headless: progress-log only (no modal dialog).
Public Sub NotifyStorageInaccessible(ByVal p As String)
    On Error Resume Next
    If modCommon.IsHeadless() Then
        modCommon.AppendProgressLog "[STORAGE-INACCESSIBLE] skip re-render keep-prior missing=" & p
        Exit Sub
    End If
    MsgBox MsgStorageA() & p & MsgStorageB(), vbExclamation, MsgStorageTitle()
End Sub

Private Function MsgStorageA() As String
    MsgStorageA = ChrW(&H5B9A) & ChrW(&H7FA9) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H304C) & ChrW(&H898B) & ChrW(&H3064) & ChrW(&H304B) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3A) & ChrW(&H20)
End Function

Private Function MsgStorageB() As String
    MsgStorageB = ChrW(&H3002) & ChrW(&H524D) & ChrW(&H56DE) & ChrW(&H8868) & ChrW(&H793A) & ChrW(&H306E) & ChrW(&H307E) & ChrW(&H307E) & ChrW(&H3067) & ChrW(&H7D99) & ChrW(&H7D9A) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H3002)
End Function

Private Function MsgStorageTitle() As String
    MsgStorageTitle = ChrW(&H5B9A) & ChrW(&H7FA9) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
End Function

Public Sub Reset()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0957] modConfigHolder.Reset ENTER"  ' [ADR-0100]
    Set m_config = Nothing
    m_isInitialized = False
End Sub

' ================================================================
' v2.3 Phase A: SetConfigKeys (partial merge API)
' ----------------------------------------------------------------
' Function: SetConfigKeys
' Summary:  Merge keyValues into the existing m_config dictionary.
'           Unlike SetConfig (which fully REPLACES m_config), this
'           overlays only the keys present in keyValues and leaves
'           every other key untouched.
'           Designed to be called immediately after a successful
'           modConfigLoader.SaveConfigKeys so the in-memory holder
'           stays in sync with the on-disk config.txt.
' Param:    ByVal keyValues As Object - Scripting.Dictionary
'                                       key => new value (string)
' Note:     If the holder has never been initialized, this call
'           initializes it with just the supplied keys (so a
'           SaveConfigKeys -> SetConfigKeys flow is safe even on
'           a cold start). Existing SetConfig API is unchanged.
' ================================================================
Public Sub SetConfigKeys(ByVal keyValues As Object)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0958] modConfigHolder.SetConfigKeys ENTER"  ' [ADR-0100]
    If keyValues Is Nothing Then Exit Sub

    ' Bootstrap holder on cold start so we never read from Nothing.
    If m_config Is Nothing Then
        Set m_config = CreateObject("Scripting.Dictionary")
    End If

    Dim k As Variant
    For Each k In keyValues.Keys
        m_config(CStr(k)) = CStr(keyValues(k))
    Next k

    m_isInitialized = True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0959] modConfigHolder.SetConfigKeys EXIT-OK"  ' [ADR-0100]
End Sub
```
