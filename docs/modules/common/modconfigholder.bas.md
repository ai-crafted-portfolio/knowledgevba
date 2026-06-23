---
title: modConfigHolder.bas
description: modConfigHolder.bas のソースコード（コピペ用）
---

# modConfigHolder.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-22 22:53 JST

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
' Change-3: storage-access fallback (ADR-0134)
' EnsureStorageAccessible returns "" when every configured storage
' folder (data / format / ui<role> / backup) is reachable, else the
' first failing path. "Reachable" = the folder exists OR its parent
' does (leaf folders are created lazily, so a not-yet-created leaf
' under a live base must NOT be treated as inaccessible). Fail-open.
' ================================================================
Public Function EnsureStorageAccessible(ByVal xlsmName As String) As String
    On Error GoTo EH
    Dim p As String
    p = GetDataDir()
    If Not DirAccessible(p) Then EnsureStorageAccessible = p: Exit Function
    p = GetFormatDir()
    If Not DirAccessible(p) Then EnsureStorageAccessible = p: Exit Function
    p = GetUiDir() & xlsmName & "\"
    If Not DirAccessible(p) Then EnsureStorageAccessible = p: Exit Function
    p = GetBackupDir()
    If Not DirAccessible(p) Then EnsureStorageAccessible = p: Exit Function
    EnsureStorageAccessible = ""
    Exit Function
EH:
    EnsureStorageAccessible = ""
End Function

Private Function DirAccessible(ByVal p As String) As Boolean
    If Len(Trim(p)) = 0 Then DirAccessible = True: Exit Function
    If modFileIO.FolderExists(p) Then DirAccessible = True: Exit Function
    DirAccessible = modFileIO.FolderExists(ParentOf(p))
End Function

Private Function ParentOf(ByVal p As String) As String
    Dim s As String
    s = p
    Do While Len(s) > 0 And Right(s, 1) = "\"
        s = Left(s, Len(s) - 1)
    Loop
    Dim pos As Long
    pos = InStrRev(s, "\")
    If pos > 1 Then
        ParentOf = Left(s, pos - 1)
    Else
        ParentOf = s
    End If
End Function

' Notify that a storage folder is unreachable and the previous screen
' is being kept. Headless: progress-log only (no modal dialog).
Public Sub NotifyStorageInaccessible(ByVal p As String)
    On Error Resume Next
    If modCommon.IsHeadless() Then
        modCommon.AppendProgressLog "[STORAGE-INACCESSIBLE] skip re-render keep-prior path=" & p
        Exit Sub
    End If
    MsgBox MsgStorageA() & p & MsgStorageB(), vbExclamation, MsgStorageTitle()
End Sub

Private Function MsgStorageA() As String
    MsgStorageA = ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30EB) & ChrW(&H30C0) & ChrW(&H306B) & ChrW(&H30A2) & ChrW(&H30AF) & ChrW(&H30BB) & ChrW(&H30B9) & ChrW(&H3067) & ChrW(&H304D) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3A) & ChrW(&H20)
End Function

Private Function MsgStorageB() As String
    MsgStorageB = ChrW(&H3002) & ChrW(&H524D) & ChrW(&H56DE) & ChrW(&H8868) & ChrW(&H793A) & ChrW(&H306E) & ChrW(&H307E) & ChrW(&H307E) & ChrW(&H3067) & ChrW(&H7D99) & ChrW(&H7D9A) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H3002)
End Function

Private Function MsgStorageTitle() As String
    MsgStorageTitle = ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC)
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
