---
title: modConfigHolder.bas
description: modConfigHolder.bas 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# modConfigHolder.bas

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 讓呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`modConfigHolder.bas`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
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
' Depends: nothing
' Related: ADR-0053 section 5.2 / 7.3 N6
'          Q7  (debugLevel default ERROR)
'          Q17 (GetDebugLevel As Long, enum return)
'          Q22 (path default root C:\KnowledgeMgr\)
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
    GetDataDir = GetValueOrDefault("data_dir", "C:\KnowledgeMgr\data\")
End Function

Public Function GetBackupDir() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0951] modConfigHolder.GetBackupDir ENTER"  ' [ADR-0100]
    GetBackupDir = GetValueOrDefault("backup_dir", "C:\KnowledgeMgr\backup\")
End Function

Public Function GetFormatDir() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0952] modConfigHolder.GetFormatDir ENTER"  ' [ADR-0100]
    GetFormatDir = GetValueOrDefault("format_dir", "C:\KnowledgeMgr\formats\")
End Function

Public Function GetUiDir() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0953] modConfigHolder.GetUiDir ENTER"  ' [ADR-0100]
    GetUiDir = GetValueOrDefault("ui_dir", "C:\KnowledgeMgr\ui\")
End Function

Public Function GetLogDir() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0954] modConfigHolder.GetLogDir ENTER"  ' [ADR-0100]
    GetLogDir = GetValueOrDefault("log_dir", "C:\KnowledgeMgr\log\")
End Function

Public Function GetConfigDir() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0955] modConfigHolder.GetConfigDir ENTER"  ' [ADR-0100]
    GetConfigDir = GetValueOrDefault("config_dir", "C:\KnowledgeMgr\config\")
End Function

Public Function GetDebugLevelStr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0956] modConfigHolder.GetDebugLevelStr ENTER"  ' [ADR-0100]
    GetDebugLevelStr = GetValueOrDefault("debugLevel", "ERROR")
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