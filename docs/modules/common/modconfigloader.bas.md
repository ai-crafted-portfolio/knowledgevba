---
title: modConfigLoader.bas
description: modConfigLoader.bas のソースコード（コピペ用）
---

# modConfigLoader.bas

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: 標準モジュール

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
' Module:  modConfigLoader (Phase 2 task 2.3 / utility layer)
' Summary: Reads xxxx_config.txt at startup and feeds modConfigHolder
'          via SetConfig.
'          ADR-0053 section 2.9 external I/O isolation rule compliant.
'          Implements the 5-key load fallback of v2_config_schema.md.
' Version: v2.1 (2026-05-16 EOD, Q1-Q57 resolved)
'          v2.3 Phase A: added SaveConfigKeys (partial writeback).
' Depends: modStanzaIO (parser), modConfigHolder (holder), clsLogger
' Related: ADR-0053 section 2.9 / 7.3 N6, ADR-0054, ADR-0049
'          Q7  (debugLevel default ERROR, config.txt edited directly,
'               modConfigLoader is read-only)
'          Q8  (xlsm name -> 3 config files)
'          Q17 (public I/F: LoadConfig / Reload, no Save API)
'               -> v2.3: added SaveConfigKeys partial-writeback API
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
Private Const DEFAULT_CONFIG_DIR As String = "C:\KnowledgeMgr\"
Private Const DEFAULT_UI_SCHEMA_FAIL_MODE As String = "safeDefault"
Private Const DEFAULT_SYSTEM_SHEET_VISIBILITY As String = "Hidden"
Private Const DEFAULT_AUTO_RELOAD_ON_STARTUP As String = "TRUE"
Private Const DEFAULT_MIGRATE_BACKUP_ENABLED As String = "TRUE"

' v2.3 Phase A: SaveConfigKeys (partial writeback) constants. Declared at
' module top because VBA requires module-level Const before any Sub/Function.
Private Const SAVE_TMP_SUFFIX As String = ".tmp"
Private Const SAVE_DEFAULT_SECTION As String = "added_by_setup"

' v2.1: last loaded path kept for Reload
Private m_currentConfigPath As String

' ----------------------------------------------------------------
' Public I/F (Q17: of the 8 functions, the loader exposes only
' LoadConfig + Reload. GetValue / GetValueOrDefault / GetDataDir
' belong to modConfigHolder.)
' v2.3 Phase A: SaveConfigKeys added (partial writeback).
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

' ================================================================
' v2.3 Phase A: SaveConfigKeys (partial writeback API)
' ----------------------------------------------------------------
' design ref: migration_plan_v1_to_v2_rev2 / sekkeisho_v23_accepted
' Adds partial writeback so M-11 / M-10 settings screens (and
' future admin features) can persist edits to <xlsmName>_config.txt
' without losing comments, section headers or unrelated keys.
' Atomic by writing to a sibling .tmp file and renaming on success.
' CP932 + CRLF preserved (CP932 = "Shift_JIS" via ADODB.Stream).
' Logging is best-effort: failures Debug.Print and return 0.
' ================================================================

' (SAVE_TMP_SUFFIX / SAVE_DEFAULT_SECTION moved to module top - VBA requires module-level Const before any Sub/Function)

' ================================================================
' Function: SaveConfigKeys
' Summary:  Partial writeback into <xlsmName>_config.txt.
'           - lines of shape 'key=value' whose key appears in
'             keyValues are rewritten with the new value
'           - blank lines, '#' comments, '[section]' headers are
'             left untouched
'           - keys present in keyValues but NOT in the existing
'             file are appended at the end under a marker section
'             '[added_by_setup]'
'           - write goes to <path>.tmp then is renamed onto path
' Param:    xlsmName  As String  - logical name ("kanri" etc),
'                                  same key as LoadConfig uses
'           keyValues As Object  - Scripting.Dictionary
'                                  key => new value (string)
' Return:   Long - count of keys actually written (updated+added).
'                  0 on any I/O error (logged, see Note).
' Note:     If LoadConfig has not been called yet, ResolveConfigDir
'           falls back to DEFAULT_CONFIG_DIR (C:\KnowledgeMgr\),
'           same as LoadConfig. Caller is responsible for invoking
'           modConfigHolder.SetConfigKeys to sync in-memory state.
' ================================================================
Public Function SaveConfigKeys(ByVal xlsmName As String, ByVal keyValues As Object) As Long
    On Error GoTo ErrHandler

    If keyValues Is Nothing Then
        SaveConfigKeys = 0
        Exit Function
    End If

    Dim configFilePath As String
    configFilePath = ResolveConfigDir() & xlsmName & "_config.txt"

    Dim tmpPath As String
    tmpPath = configFilePath & SAVE_TMP_SUFFIX

    ' (1) Read current file into memory (if it exists). Lines are
    '     preserved as-is so we keep blank lines, comments, headers.
    Dim lines As Collection
    Set lines = ReadAllLines(configFilePath)

    ' (2) Rewrite lines whose key is in keyValues, track which keys
    '     were consumed so we can append the rest at the end.
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1 ' vbTextCompare = case-insensitive

    Dim updatedCount As Long
    updatedCount = RewriteMatchingLines(lines, keyValues, seen)

    ' (3) Append keys that did not appear in the file.
    Dim addedCount As Long
    addedCount = AppendMissingKeys(lines, keyValues, seen)

    ' (4) Write to .tmp then rename onto the real path (atomic-ish).
    WriteAllLines tmpPath, lines
    PromoteTmpToFinal tmpPath, configFilePath

    SaveConfigKeys = updatedCount + addedCount
    Exit Function

ErrHandler:
    LogSaveError "SaveConfigKeys", xlsmName, Err.Number, Err.Description
    ' Try to clean up tmp file (best effort, ignore errors)
    On Error Resume Next
    CleanupTmp configFilePath & SAVE_TMP_SUFFIX
    On Error GoTo 0
    SaveConfigKeys = 0
End Function

' ----------------------------------------------------------------
' Private: ReadAllLines
'   Returns Collection of String. If file missing, returns empty.
'   Uses ADODB.Stream with Shift_JIS (CP932) charset.
' ----------------------------------------------------------------
Private Function ReadAllLines(ByVal filePath As String) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Set ReadAllLines = result
        Exit Function
    End If

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "Shift_JIS"
    stream.Open
    stream.LoadFromFile filePath
    Dim allText As String
    allText = stream.ReadText
    stream.Close
    Set stream = Nothing

    ' Normalize line endings: split on vbCrLf first, then on vbLf
    ' to be robust against tools that stripped CR.
    Dim parts() As String
    parts = SplitLinesPreserving(allText)

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        result.Add parts(i)
    Next i

    Set ReadAllLines = result
End Function

' ----------------------------------------------------------------
' Private: SplitLinesPreserving
'   Splits on CRLF (preferred) or LF, returning array of lines
'   without trailing line-terminator characters.
' ----------------------------------------------------------------
Private Function SplitLinesPreserving(ByVal s As String) As String()
    Dim normalized As String
    normalized = Replace(s, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    SplitLinesPreserving = Split(normalized, vbLf)
End Function

' ----------------------------------------------------------------
' Private: RewriteMatchingLines
'   For each line of shape 'key=value', if key is in keyValues,
'   replace the value portion with the new value (keeping any
'   leading whitespace). Records consumed keys into 'seen'.
' Return: count of lines actually modified.
' ----------------------------------------------------------------
Private Function RewriteMatchingLines(ByVal lines As Collection, _
                                       ByVal keyValues As Object, _
                                       ByVal seen As Object) As Long
    Dim modified As Long
    modified = 0

    Dim i As Long
    For i = 1 To lines.Count
        Dim line As String
        line = lines.Item(i)

        Dim parsed As Object
        Set parsed = ParseKeyValueLine(line)
        If parsed Is Nothing Then GoTo NextLine

        Dim k As String
        k = parsed("key")
        If Not keyValues.Exists(k) Then GoTo NextLine

        ' Found a key to update. Build new line preserving leading
        ' whitespace; pick up the new value from keyValues.
        Dim newLine As String
        newLine = parsed("leading") & k & "=" & CStr(keyValues(k))

        ' Replace by deleting and re-inserting at the same position.
        lines.Remove i
        If i > lines.Count Then
            lines.Add newLine
        Else
            lines.Add newLine, Before:=i
        End If
        seen(k) = True
        modified = modified + 1
NextLine:
    Next i

    RewriteMatchingLines = modified
End Function

' ----------------------------------------------------------------
' Private: ParseKeyValueLine
'   Returns Scripting.Dictionary with keys "leading", "key" if the
'   line is a key=value line, otherwise Nothing.
'   Comment lines ('#') and section headers ('[xxx]') return Nothing.
' ----------------------------------------------------------------
Private Function ParseKeyValueLine(ByVal line As String) As Object
    Set ParseKeyValueLine = Nothing

    Dim trimmed As String
    trimmed = Trim(line)
    If Len(trimmed) = 0 Then Exit Function
    If Left(trimmed, 1) = "#" Then Exit Function
    If Left(trimmed, 1) = "[" Then Exit Function ' section header

    Dim eqPos As Long
    eqPos = InStr(line, "=")
    If eqPos <= 1 Then Exit Function

    ' Split into leading-whitespace + key
    Dim leading As String
    Dim keyPart As String
    Dim keyEnd As Long
    keyEnd = eqPos - 1

    ' Extract leading whitespace from the original line
    Dim p As Long
    p = 1
    Do While p <= Len(line)
        Dim ch As String
        ch = Mid(line, p, 1)
        If ch <> " " And ch <> vbTab Then Exit Do
        p = p + 1
    Loop
    leading = Left(line, p - 1)
    keyPart = Mid(line, p, keyEnd - p + 1)
    keyPart = RTrim(keyPart)

    ' Reject if key contains whitespace inside (defensive)
    If InStr(keyPart, " ") > 0 Then Exit Function
    If Len(keyPart) = 0 Then Exit Function

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("leading") = leading
    d("key") = keyPart
    Set ParseKeyValueLine = d
End Function

' ----------------------------------------------------------------
' Private: AppendMissingKeys
'   For each key in keyValues that was NOT seen during rewrite,
'   append 'key=value' under a marker section at file end.
'   If the marker section header is not present yet, add it.
' Return: count of keys appended.
' ----------------------------------------------------------------
Private Function AppendMissingKeys(ByVal lines As Collection, _
                                    ByVal keyValues As Object, _
                                    ByVal seen As Object) As Long
    Dim missingKeys As Collection
    Set missingKeys = New Collection
    Dim k As Variant
    For Each k In keyValues.Keys
        If Not seen.Exists(CStr(k)) Then
            missingKeys.Add CStr(k)
        End If
    Next k
    If missingKeys.Count = 0 Then
        AppendMissingKeys = 0
        Exit Function
    End If

    ' Ensure a blank separator + marker section header.
    If lines.Count > 0 Then
        Dim lastLine As String
        lastLine = lines.Item(lines.Count)
        If Len(Trim(lastLine)) > 0 Then
            lines.Add ""
        End If
    End If
    lines.Add "[" & SAVE_DEFAULT_SECTION & "]"

    Dim i As Long
    For i = 1 To missingKeys.Count
        Dim ky As String
        ky = missingKeys.Item(i)
        lines.Add ky & "=" & CStr(keyValues(ky))
    Next i
    AppendMissingKeys = missingKeys.Count
End Function

' ----------------------------------------------------------------
' Private: WriteAllLines
'   Writes the collection of lines to filePath as CP932 + CRLF.
' ----------------------------------------------------------------
Private Sub WriteAllLines(ByVal filePath As String, ByVal lines As Collection)
    Dim buf As String
    Dim i As Long
    For i = 1 To lines.Count
        If i > 1 Then buf = buf & vbCrLf
        buf = buf & lines.Item(i)
    Next i

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "Shift_JIS"
    stream.Open
    stream.WriteText buf
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing
End Sub

' ----------------------------------------------------------------
' Private: PromoteTmpToFinal
'   Atomic-ish rename of tmpPath onto finalPath. Uses FSO so we can
'   overwrite. If finalPath exists, it is deleted first; then tmp is
'   renamed. On Windows VBA there is no true atomic, but this brief
'   window is acceptable for our config use-case.
' ----------------------------------------------------------------
Private Sub PromoteTmpToFinal(ByVal tmpPath As String, ByVal finalPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(tmpPath) Then
        Err.Raise vbObjectError + 5001, _
            "modConfigLoader.PromoteTmpToFinal", _
            "tmp not written: " & tmpPath
    End If
    If fso.FileExists(finalPath) Then
        fso.DeleteFile finalPath, True
    End If
    fso.MoveFile tmpPath, finalPath
End Sub

' ----------------------------------------------------------------
' Private: CleanupTmp
'   Best-effort delete of leftover .tmp file. Used by ErrHandler.
' ----------------------------------------------------------------
Private Sub CleanupTmp(ByVal tmpPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(tmpPath) Then fso.DeleteFile tmpPath, True
End Sub

' ----------------------------------------------------------------
' Private: LogSaveError
'   Best-effort logging via clsLogger on the LOG sheet of the
'   current workbook. If logging itself fails, fall back to
'   Debug.Print. Always swallows the error so the caller can
'   return 0.
' ----------------------------------------------------------------
Private Sub LogSaveError(ByVal func As String, _
                          ByVal xlsmName As String, _
                          ByVal errNum As Long, _
                          ByVal errDesc As String)
    On Error Resume Next
    Dim logSheet As Object
    Set logSheet = ThisWorkbook.Worksheets("LOG")
    If Not logSheet Is Nothing Then
        Dim oLogger As clsLogger
        Set oLogger = New clsLogger
        oLogger.Init logSheet
        oLogger.LogError "modConfigLoader", func, _
            "SaveConfigKeys failed: " & errNum & " " & errDesc, _
            xlsmName, "SAVECFG-ERR-EE-001"
    End If
    Debug.Print "[ERR] modConfigLoader." & func & "(" & xlsmName & "): " & errNum & " " & errDesc
    On Error GoTo 0
End Sub
```
