---
title: modConfigLoader.bas
---

# modConfigLoader.bas

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 3 ブック共通 |
| 役割 | 起動時に <ブック名>_config.txt を読み込み modConfigHolder へ渡す（読み取り専用） |
| 行数 | 189 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modConfigLoader.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modConfigLoader"
' ================================================================
' Module:  modConfigLoader (Phase 2 task 2.3 / utility layer)
' Summary: Reads xxxx_config.txt at startup and feeds modConfigHolder
'          via SetConfig.
'          ADR-0053 section 2.9 external I/O isolation rule compliant.
'          Implements the 5-key load fallback of v2_config_schema.md.
' Version: v2.1 (2026-05-16 EOD, Q1-Q57 resolved)
' Depends: modStanzaIO (parser), modConfigHolder (holder)
' Related: ADR-0053 section 2.9 / 7.3 N6, ADR-0054, ADR-0049
'          Q7  (debugLevel default ERROR, config.txt edited directly,
'               modConfigLoader is read-only)
'          Q8  (xlsm name -> 3 config files)
'          Q17 (public I/F: LoadConfig / Reload, no Save API)
'          Q22 (default path root C:\KnowledgeMgr\)
' Note:    Japanese block comments were lost (no backup) during a
'          v2.1 bug-fix edit and restored here in ASCII. Logic unchanged.
' ================================================================
Option Explicit

' ----------------------------------------------------------------
' Constants (v2.1 default values, per Q7 / Q22)
' ----------------------------------------------------------------
Private Const DEFAULT_DEBUG_LEVEL As String = "ERROR"
Private Const DEFAULT_LOG_ROTATION_ROWS As String = "10000"
Private Const DEFAULT_DATA_DIR As String = "C:\KnowledgeMgr\data\"
Private Const DEFAULT_FORMAT_DIR As String = "C:\KnowledgeMgr\formats\"
Private Const DEFAULT_UI_DIR As String = "C:\KnowledgeMgr\ui\"
Private Const DEFAULT_BACKUP_DIR As String = "C:\KnowledgeMgr\backup\"
Private Const DEFAULT_LOG_DIR As String = "C:\KnowledgeMgr\log\"
Private Const DEFAULT_CONFIG_DIR As String = "C:\KnowledgeMgr\config\"
Private Const DEFAULT_UI_SCHEMA_FAIL_MODE As String = "safeDefault"
Private Const DEFAULT_SYSTEM_SHEET_VISIBILITY As String = "Hidden"
Private Const DEFAULT_AUTO_RELOAD_ON_STARTUP As String = "TRUE"
Private Const DEFAULT_MIGRATE_BACKUP_ENABLED As String = "TRUE"

' v2.1: last loaded path kept for Reload
Private m_currentConfigPath As String

' ----------------------------------------------------------------
' Public I/F (Q17: of the 8 functions, the loader exposes only
' LoadConfig + Reload. GetValue / GetValueOrDefault / GetDataDir
' belong to modConfigHolder.)
' ----------------------------------------------------------------

' ================================================================
' Function: LoadConfig
' Summary:  Reads the config.txt for xlsmName and feeds modConfigHolder
'           via SetConfig. v2.1 (Q8): xlsmName maps to 3 config files.
' Param:    ByVal xlsmName As String
' Return:   Boolean - TRUE on success / FALSE when the file is absent
'           or no key could be parsed (caller emits WARN).
' ================================================================
Public Function LoadConfig(ByVal xlsmName As String) As Boolean
    On Error GoTo ErrHandler

    Dim configFilePath As String
    configFilePath = ResolveConfigDir() & xlsmName & "_config.txt"
    m_currentConfigPath = configFilePath

    LoadConfig = LoadConfigFromPath(configFilePath)
    Exit Function

ErrHandler:
    ' ret=False exit must NOT pollute the holder; seed defensive
    ' defaults only when the holder has not been initialized
    ' (architecture OQ-07 safeDefault).
    SeedHolderDefaultsIfEmpty
    LoadConfig = False
End Function

' ----------------------------------------------------------------
' Private: ResolveConfigDir
'   Resolves the config-file directory. When the holder is already
'   initialized and carries a config_dir value, that value wins;
'   otherwise fall back to DEFAULT_CONFIG_DIR (Q22).
'   This also lets callers / tests redirect the config directory
'   without abort (architecture OQ-07 safeDefault).
' ----------------------------------------------------------------
Private Function ResolveConfigDir() As String
    Dim dir As String
    dir = vbNullString
    If modConfigHolder.IsInitialized() Then
        dir = modConfigHolder.GetValue("config_dir")
    End If
    If Len(dir) = 0 Then
        dir = DEFAULT_CONFIG_DIR
    End If
    ResolveConfigDir = dir
End Function

' ----------------------------------------------------------------
' Private: SeedHolderDefaultsIfEmpty
'   Seeds the holder with defensive defaults ONLY when SetConfig
'   has not yet been called. An already-initialized holder is left
'   untouched, so a failed load never overwrites a good config and
'   never leaves loader-internal partial state behind (architecture
'   OQ-07 safeDefault, holder-pollution prevention).
' ----------------------------------------------------------------
Private Sub SeedHolderDefaultsIfEmpty()
    If modConfigHolder.IsInitialized() Then Exit Sub
    Dim fallbackDict As Object
    Set fallbackDict = CreateObject("Scripting.Dictionary")
    LoadDefaults fallbackDict
    modConfigHolder.SetConfig fallbackDict
End Sub

' ================================================================
' Function: Reload (Q17, Q7 dynamic debugLevel update)
' Summary:  Re-reads the current config path (clsLogger 5s polling).
' Return:   Boolean
' ================================================================
Public Function Reload() As Boolean
    If Len(m_currentConfigPath) = 0 Then
        Reload = False
        Exit Function
    End If
    Reload = LoadConfigFromPath(m_currentConfigPath)
End Function

' ----------------------------------------------------------------
' Private
' ----------------------------------------------------------------

Private Function LoadConfigFromPath(ByVal configFilePath As String) As Boolean
    On Error GoTo ErrHandler

    Dim configDict As Object
    Set configDict = CreateObject("Scripting.Dictionary")

    LoadDefaults configDict

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(configFilePath) Then
        ' File absent: do not write the loader-internal configDict into
        ' the holder. Seed defaults only when the holder is not yet
        ' initialized (architecture OQ-07 safeDefault / holder-pollution
        ' prevention).
        SeedHolderDefaultsIfEmpty
        LoadConfigFromPath = False
        Exit Function
    End If

    Dim sections As Collection
    Set sections = modStanzaIO.ParseStanzaFile(configFilePath)

    Dim sec As Object
    Dim i As Long
    For i = 1 To sections.Count
        Set sec = sections.Item(i)
        Dim kv As Object
        Set kv = sec.KeyValues
        Dim k As Variant
        For Each k In kv.Keys
            configDict(CStr(k)) = CStr(kv(k))
        Next k
    Next i

    modConfigHolder.SetConfig configDict
    LoadConfigFromPath = True
    Exit Function

ErrHandler:
    ' Error mid-parse: do NOT write the partially-built configDict into
    ' the holder. Seed defensive defaults only when the holder has not
    ' been initialized (architecture OQ-07 safeDefault).
    SeedHolderDefaultsIfEmpty
    LoadConfigFromPath = False
End Function

' ----------------------------------------------------------------
' Private: LoadDefaults
'   Fills configDict with the default values (per Q7 / Q22).
' ----------------------------------------------------------------
Private Sub LoadDefaults(ByVal configDict As Object)
    configDict("debugLevel") = DEFAULT_DEBUG_LEVEL
    configDict("logRotationRows") = DEFAULT_LOG_ROTATION_ROWS
    configDict("data_dir") = DEFAULT_DATA_DIR
    configDict("format_dir") = DEFAULT_FORMAT_DIR
    configDict("ui_dir") = DEFAULT_UI_DIR
    configDict("backup_dir") = DEFAULT_BACKUP_DIR
    configDict("log_dir") = DEFAULT_LOG_DIR
    configDict("config_dir") = DEFAULT_CONFIG_DIR
    configDict("uiSchemaFailMode") = DEFAULT_UI_SCHEMA_FAIL_MODE
    configDict("systemSheetVisibility") = DEFAULT_SYSTEM_SHEET_VISIBILITY
    configDict("autoReloadOnStartup") = DEFAULT_AUTO_RELOAD_ON_STARTUP
    configDict("migrateBackupEnabled") = DEFAULT_MIGRATE_BACKUP_ENABLED
End Sub
```
