---
title: modConfigHolder.bas
---

# modConfigHolder.bas

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 3 ブック共通 |
| 役割 | 設定値をメモリに保持し、各層へ取得用のメソッドで提供する |
| 行数 | 153 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modConfigHolder.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modConfigHolder"
' ================================================================
' ?ｿｽ?ｿｽ?ｿｽW?ｿｽ?ｿｽ?ｿｽ[?ｿｽ?ｿｽ: modConfigHolder?ｿｽiPhase 2 task 2.3 / ?ｿｽ?ｿｽ?ｿｽ[?ｿｽe?ｿｽB?ｿｽ?ｿｽ?ｿｽe?ｿｽB?ｿｽw?ｿｽj
' ?ｿｽT?ｿｽv:       config.txt ?ｿｽ?ｿｽ in-memory holder
'             ADR-0053 ?ｿｽ?ｿｽ7.3 N6 NM-2 ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽimodConfigLoader ?ｿｽ?ｿｽ config ?ｿｽ?ｿｽﾇ搾ｿｽ?ｿｽﾝ、
'             ?ｿｽ{ holder ?ｿｽ?ｿｽ SetConfig ?ｿｽ?ｿｽ ?ｿｽ?ｿｽ?ｿｽw?ｿｽ?ｿｽ GetValue ?ｿｽﾅ托ｿｽ?ｿｽ?ｿｽ?ｿｽ謫ｾ?ｿｽj
' Version:    v2.1?ｿｽi2026-05-16 EOD?ｿｽAQ1-Q57 ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽf?ｿｽj
' ?ｿｽﾋ托ｿｽ?ｿｽ?ｿｽ:     ?ｿｽﾈゑｿｽ
' ?ｿｽﾖ連:       ADR-0053 ?ｿｽ?ｿｽ5.2 / ?ｿｽ?ｿｽ7.3 N6
'             Q7?ｿｽidebugLevel ?ｿｽ?ｿｽ?ｿｽ?ｿｽ ERROR?ｿｽj
'             Q17?ｿｽiGetDebugLevel As Long?ｿｽAenum ?ｿｽl?ｿｽﾔ却?ｿｽj
'             Q22?ｿｽi?ｿｽp?ｿｽX?ｿｽ?ｿｽ?ｿｽ?ｿｽl C:\KnowledgeMgr\ ?ｿｽz?ｿｽ?ｿｽ?ｿｽj
' ?ｿｽd?ｿｽv?ｿｽK?ｿｽ?ｿｽ:   ?ｿｽN?ｿｽ?ｿｽ?ｿｽ?ｿｽ 1 ?ｿｽ?ｿｽ read + holder ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ謫ｾ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
'             clsLogger.RefreshDebugLevel ?ｿｽ?ｿｽ?ｿｽ{ holder ?ｿｽ?ｿｽ GetDebugLevel ?ｿｽ?ｿｽ?ｿｽg?ｿｽ?ｿｽ
' ================================================================
Option Explicit

' v2.1: debugLevel enum 6 ?ｿｽl?ｿｽimodCommon ?ｿｽﾉゑｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾌ定数?ｿｽ?ｿｽ?ｿｽ?ｿｽA?ｿｽ{?ｿｽ?ｿｽ?ｿｽﾅは参?ｿｽﾆのゑｿｽ?ｿｽﾟ重?ｿｽ?ｿｽ?ｿｽ?ｿｽ`?ｿｽj
Public Const DEBUG_LEVEL_OFF As Long = 0
Public Const DEBUG_LEVEL_ERROR As Long = 1
Public Const DEBUG_LEVEL_WARN As Long = 2
Public Const DEBUG_LEVEL_INFO As Long = 3
Public Const DEBUG_LEVEL_DEBUG As Long = 4
Public Const DEBUG_LEVEL_TRACE As Long = 5

Private m_config As Object              ' Scripting.Dictionary
Private m_isInitialized As Boolean

' ----------------------------------------------------------------
' Public I/F
' ----------------------------------------------------------------

' ================================================================
' ?ｿｽﾖ撰ｿｽ?ｿｽ?ｿｽ: SetConfig
' ?ｿｽT?ｿｽv:   modConfigLoader ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾌ全 config ?ｿｽ?ｿｽ?ｿｽｯ趣ｿｽ?ｿｽ holder ?ｿｽﾉ保趣ｿｽ
' ?ｿｽ?ｿｽ?ｿｽ?ｿｽ:   ByVal configDict As Object - Scripting.Dictionary?ｿｽikey?ｿｽ?ｿｽvalue?ｿｽj
' ================================================================
Public Sub SetConfig(ByVal configDict As Object)
    Set m_config = configDict
    m_isInitialized = True
End Sub

' ================================================================
' ?ｿｽﾖ撰ｿｽ?ｿｽ?ｿｽ: GetValue
' ?ｿｽT?ｿｽv:   key ?ｿｽﾉ対会ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ value ?ｿｽ?ｿｽ?ｿｽ謫ｾ?ｿｽiin-memory ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽj
' ?ｿｽ?ｿｽ?ｿｽ?ｿｽ:   ByVal key As String - config key
' ?ｿｽﾟゑｿｽl: String - ?ｿｽl?ｿｽi?ｿｽ?ｿｽ?ｿｽ?ｿｽ` / ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽﾈゑｿｽ vbNullString?ｿｽj
' ================================================================
Public Function GetValue(ByVal key As String) As String
    If Not m_isInitialized Then
        GetValue = vbNullString
        Exit Function
    End If
    If m_config Is Nothing Then
        GetValue = vbNullString
        Exit Function
    End If
    If m_config.Exists(key) Then
        GetValue = CStr(m_config(key))
    Else
        GetValue = vbNullString
    End If
End Function

' ================================================================
' ?ｿｽﾖ撰ｿｽ?ｿｽ?ｿｽ: GetValueOrDefault
' ?ｿｽT?ｿｽv:   GetValue + ?ｿｽf?ｿｽt?ｿｽH?ｿｽ?ｿｽ?ｿｽg?ｿｽl?ｿｽt?ｿｽH?ｿｽ[?ｿｽ?ｿｽ?ｿｽo?ｿｽb?ｿｽN
' ================================================================
Public Function GetValueOrDefault(ByVal key As String, ByVal defaultValue As String) As String
    Dim v As String
    v = GetValue(key)
    If Len(v) = 0 Then
        GetValueOrDefault = defaultValue
    Else
        GetValueOrDefault = v
    End If
End Function

' ================================================================
' ?ｿｽﾖ撰ｿｽ?ｿｽ?ｿｽ: HasKey
' ?ｿｽT?ｿｽv:   key ?ｿｽ?ｿｽ?ｿｽ?ｿｽ`?ｿｽ?ｿｽ?ｿｽ?ｿｽﾄゑｿｽ?ｿｽ驍ｩ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
' ================================================================
Public Function HasKey(ByVal key As String) As Boolean
    If Not m_isInitialized Or m_config Is Nothing Then
        HasKey = False
    Else
        HasKey = m_config.Exists(key)
    End If
End Function

' ================================================================
' ?ｿｽﾖ撰ｿｽ?ｿｽ?ｿｽ: IsInitialized
' ?ｿｽT?ｿｽv:   SetConfig ?ｿｽﾏみゑｿｽ?ｿｽ?ｿｽ?ｿｽ?ｿｽ
' ================================================================
Public Function IsInitialized() As Boolean
    IsInitialized = m_isInitialized
End Function

' ================================================================
' ?ｿｽﾖ撰ｿｽ?ｿｽ?ｿｽ: GetDebugLevel?ｿｽiQ17 ?ｿｽm?ｿｽ?ｿｽA?ｿｽﾟゑｿｽl Long?ｿｽAenum ?ｿｽl?ｿｽj
' ?ｿｽT?ｿｽv:   debugLevel ?ｿｽ?ｿｽ Long enum ?ｿｽl?ｿｽﾅ取得
'         ?ｿｽ?ｿｽ?ｿｽ?ｿｽl = ERROR?ｿｽiQ7 ?ｿｽm?ｿｽ?ｿｽA?ｿｽ{?ｿｽﾔ運?ｿｽp?ｿｽﾅ擾ｿｽ?ｿｽ\?ｿｽ?ｿｽ?ｿｽj
' ?ｿｽﾟゑｿｽl: Long - DEBUG_LEVEL_OFF/ERROR/WARN/INFO/DEBUG/TRACE ?ｿｽﾌゑｿｽ?ｿｽ?ｿｽ?ｿｽ黷ｩ
' ================================================================
Public Function GetDebugLevel() As Long
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
End Function

' ================================================================
' Wrapper accessors for common dir keys
' ================================================================
Public Function GetDataDir() As String
    GetDataDir = GetValueOrDefault("data_dir", "C:\KnowledgeMgr\data\")
End Function

Public Function GetBackupDir() As String
    GetBackupDir = GetValueOrDefault("backup_dir", "C:\KnowledgeMgr\backup\")
End Function

Public Function GetFormatDir() As String
    GetFormatDir = GetValueOrDefault("format_dir", "C:\KnowledgeMgr\formats\")
End Function

Public Function GetUiDir() As String
    GetUiDir = GetValueOrDefault("ui_dir", "C:\KnowledgeMgr\ui\")
End Function

Public Function GetLogDir() As String
    GetLogDir = GetValueOrDefault("log_dir", "C:\KnowledgeMgr\log\")
End Function

Public Function GetConfigDir() As String
    GetConfigDir = GetValueOrDefault("config_dir", "C:\KnowledgeMgr\config\")
End Function

Public Function GetDebugLevelStr() As String
    GetDebugLevelStr = GetValueOrDefault("debugLevel", "ERROR")
End Function

Public Sub Reset()
    Set m_config = Nothing
    m_isInitialized = False
End Sub
```
