---
title: modKnowledgeFileIO.bas
description: modKnowledgeFileIO.bas のソースコード（コピペ用）
---

# modKnowledgeFileIO.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-11 15:52 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modKnowledgeFileIO.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modKnowledgeFileIO"
' ================================================================
' Module:   modKnowledgeFileIO (Phase 2 / utility layer)
' Summary:  Knowledge .txt read/write/delete/backup
'           v2.1: dedicated parser for ###xxx### stanza (Q47)
'           ADR-0053 2.9 external I/O isolation
'           write enforce: register-only (ThisWorkbook.Name, architecture 3.6)
' Version:  v2.1 (2026-05-19, Q1-Q57 all reflected)
' Deps:     modConfigHolder, modCommon
' ================================================================
Option Explicit

' xlsm gate enforce (architecture 3.6): SaveKnowledge runs only from
' the register workbook. Symmetric with modFormatLoader.SaveFormat (Q19).
' v2.2: M-12 field-reflection per-run backup subfolder prefix.
Private Const REFLECT_PREFIX As String = "reflect_"

' ----------------------------------------------------------------
' Public I/F: read
' ----------------------------------------------------------------
Public Function LoadKnowledge(ByVal knowledgeNo As String) As Object
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1262] modKnowledgeFileIO.LoadKnowledge ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = modConfigHolder.GetDataDir() & knowledgeNo & ".txt"

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Set LoadKnowledge = result
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1263] modKnowledgeFileIO.LoadKnowledge EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim content As String
    content = ReadAllTextShiftJIS(filePath)

    Set result = ParseKnowledgeStanza(content)
    Set LoadKnowledge = result
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1264] modKnowledgeFileIO.LoadKnowledge EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1265] modKnowledgeFileIO.LoadKnowledge EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set LoadKnowledge = CreateObject("Scripting.Dictionary")
End Function

' ================================================================
' GetKnowledgeTimestamp (Q46 optimistic lock)
' ================================================================
Public Function GetKnowledgeTimestamp(ByVal knowledgeNo As String) As Date
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1266] modKnowledgeFileIO.GetKnowledgeTimestamp ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = modConfigHolder.GetDataDir() & knowledgeNo & ".txt"

    If Len(Dir(filePath)) > 0 Then
        GetKnowledgeTimestamp = FileDateTime(filePath)
    Else
        GetKnowledgeTimestamp = 0
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1267] modKnowledgeFileIO.GetKnowledgeTimestamp EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1268] modKnowledgeFileIO.GetKnowledgeTimestamp EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    GetKnowledgeTimestamp = 0
End Function

' ----------------------------------------------------------------
' Public I/F: write (optimistic lock Q46)
' ----------------------------------------------------------------
Public Function SaveKnowledge( _
    ByVal knowledgeNo As String, _
    ByVal knowledgeDict As Object, _
    ByVal originalTimestamp As Date _
) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1269] modKnowledgeFileIO.SaveKnowledge ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    ' xlsm write gate (architecture 3.6): knowledge save is permitted
    ' only from the register workbook. Symmetric with SaveFormat (Q19).
    ' Return 3 = rejected by xlsm gate (0=ok / 1=conflict / 2=CP932).
    If Not IsRegisterWorkbook() Then
        SaveKnowledge = 3
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1270] modKnowledgeFileIO.SaveKnowledge EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim filePath As String
    filePath = modConfigHolder.GetDataDir() & knowledgeNo & ".txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(modConfigHolder.GetDataDir()) Then
        fso.CreateFolder modConfigHolder.GetDataDir()
    End If

    If originalTimestamp <> 0 Then
        If fso.FileExists(filePath) Then
            Dim currentTs As Date
            currentTs = FileDateTime(filePath)
            If Abs(DateDiff("s", originalTimestamp, currentTs)) > 2 Then
                SaveKnowledge = 1  ' conflict
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1271] modKnowledgeFileIO.SaveKnowledge EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
        End If
    End If

    Dim content As String
    content = SerializeKnowledgeStanza(knowledgeDict)

    If Not IsCP932Compatible(content) Then
        SaveKnowledge = 2  ' CP932 incompatible
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1272] modKnowledgeFileIO.SaveKnowledge EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    ' 2026-06-06: reject control chars (Chr(0)-Chr(31) except Tab/LF/CR) -- ADR-0112
    If Not IsControlCharSafe(content) Then
        SaveKnowledge = 2
        Exit Function
    End If

    WriteAllTextShiftJIS filePath, content
    SaveKnowledge = 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1273] modKnowledgeFileIO.SaveKnowledge EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1274] modKnowledgeFileIO.SaveKnowledge EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    SaveKnowledge = 2
End Function

' ================================================================
' Private helpers
' ================================================================

' Q47 ###xxx### format serializer
Private Function SerializeKnowledgeStanza(ByVal d As Object) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1275] modKnowledgeFileIO.SerializeKnowledgeStanza ENTER"  ' [ADR-0100]
    Dim sb As String
    sb = ""
    Dim k As Variant
    For Each k In d.Keys
        sb = sb & "###" & CStr(k) & "###" & vbCrLf
        sb = sb & CStr(d(k)) & vbCrLf
    Next k
    SerializeKnowledgeStanza = sb
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1276] modKnowledgeFileIO.SerializeKnowledgeStanza EXIT-OK"  ' [ADR-0100]
End Function

' xlsm write gate (architecture 3.6): True only when the caller is the
' register workbook. Mirrors modFormatLoader.SaveFormat Q19 enforce.
' Decision order:
'   1. config token xlsmName (register / touroku) -> allowed
'   2. config token xlsmName set to any other value -> rejected
'   3. xlsmName unset -> fall back to ThisWorkbook.Name match
Private Function IsRegisterWorkbook() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1277] modKnowledgeFileIO.IsRegisterWorkbook ENTER"  ' [ADR-0100]
    On Error GoTo Bad
    Dim wbName As String
    wbName = LCase(modConfigHolder.GetValueOrDefault("xlsmName", ""))
    If wbName = "register" Or wbName = "touroku" Then
        IsRegisterWorkbook = True
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1278] modKnowledgeFileIO.IsRegisterWorkbook EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If Len(wbName) > 0 Then
        ' xlsmName is explicitly set to a non-register role -> reject
        IsRegisterWorkbook = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1279] modKnowledgeFileIO.IsRegisterWorkbook EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    ' xlsmName unset: identify by the running workbook file name
    IsRegisterWorkbook = (ThisWorkbook.Name = TOUROKU_XLSM_NAME)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1280] modKnowledgeFileIO.IsRegisterWorkbook EXIT-OK"  ' [ADR-0100]
    Exit Function
Bad:
    IsRegisterWorkbook = False
End Function

' Check if string is fully CP932-encodable round-trip
Private Function IsCP932Compatible(ByVal s As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1281] modKnowledgeFileIO.IsCP932Compatible ENTER"  ' [ADR-0100]
    On Error GoTo Bad
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2
    ado.Charset = "Shift_JIS"
    ado.Open
    ado.WriteText s
    ado.Position = 0
    Dim back As String
    back = ado.ReadText
    ado.Close
    IsCP932Compatible = (back = s)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1282] modKnowledgeFileIO.IsCP932Compatible EXIT-OK"  ' [ADR-0100]
    Exit Function
Bad:
    IsCP932Compatible = False
End Function

' Write text as Shift_JIS (CP932) with CRLF
Private Sub WriteAllTextShiftJIS(ByVal filePath As String, ByVal content As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1283] modKnowledgeFileIO.WriteAllTextShiftJIS ENTER"  ' [ADR-0100]
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2
    ado.Charset = "Shift_JIS"
    ado.Open
    ado.WriteText content
    ado.SaveToFile filePath, 2
    ado.Close
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1284] modKnowledgeFileIO.WriteAllTextShiftJIS EXIT-OK"  ' [ADR-0100]
End Sub

' Read text from Shift_JIS file
Private Function ReadAllTextShiftJIS(ByVal filePath As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1285] modKnowledgeFileIO.ReadAllTextShiftJIS ENTER"  ' [ADR-0100]
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2
    ado.Charset = "Shift_JIS"
    ado.Open
    ado.LoadFromFile filePath
    ReadAllTextShiftJIS = ado.ReadText
    ado.Close
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1286] modKnowledgeFileIO.ReadAllTextShiftJIS EXIT-OK"  ' [ADR-0100]
End Function

' Parse ###Key### / value blocks into Dictionary
Private Function ParseKnowledgeStanza(ByVal content As String) As Object
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1287] modKnowledgeFileIO.ParseKnowledgeStanza ENTER"  ' [ADR-0100]
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    ' [BUG-B7 2026-06-11] data stanza keys vary in case across generations
    ' (###FormatID### vs ###formatId###, UpdatedAt vs updatedAt). Consumers
    ' (modEntrySearch.SafeStr etc.) look keys up case-sensitively, so force
    ' TextCompare BEFORE any key is added.
    d.CompareMode = vbTextCompare
    Dim lines() As String
    Dim norm As String
    norm = Replace(content, vbCrLf, vbLf)
    norm = Replace(norm, vbCr, vbLf)
    lines = Split(norm, vbLf)
    Dim i As Long, currentKey As String, body As String
    currentKey = ""
    body = ""
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = lines(i)
        If Len(ln) >= 6 And Left$(ln, 3) = "###" And Right$(ln, 3) = "###" Then
            If Len(currentKey) > 0 Then
                If Right$(body, 1) = vbLf Then body = Left$(body, Len(body) - 1)
                d(currentKey) = body
            End If
            currentKey = Mid$(ln, 4, Len(ln) - 6)
            body = ""
        Else
            If Len(currentKey) > 0 Then
                If Len(body) > 0 Then body = body & vbLf
                body = body & ln
            End If
        End If
    Next i
    If Len(currentKey) > 0 Then
        If Right$(body, 1) = vbLf Then body = Left$(body, Len(body) - 1)
        d(currentKey) = body
    End If
    Set ParseKnowledgeStanza = d
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1288] modKnowledgeFileIO.ParseKnowledgeStanza EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' v2.1 minimal impl (2026-05-19): replaces empty stubs so legacy
' E2E tests (_e2e_touroku / _e2e_kensaku / _e2e_kanri) can verify
' real semantics. Scope: enumerate data_dir for *.txt knowledge files.
' ================================================================

' ListAllKnowledges: return Collection of knowledgeNo (file basename minus .txt)
Public Function ListAllKnowledges() As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1289] modKnowledgeFileIO.ListAllKnowledges ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dir As String
    dir = modConfigHolder.GetDataDir()
    If Not fso.FolderExists(dir) Then
        Set ListAllKnowledges = result
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1290] modKnowledgeFileIO.ListAllKnowledges EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim folder As Object
    Set folder = fso.GetFolder(dir)
    Dim f As Object
    For Each f In folder.Files
        If LCase(fso.GetExtensionName(f.Name)) = "txt" Then
            result.Add fso.GetBaseName(f.Name)
        End If
    Next f
    Set ListAllKnowledges = result
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1291] modKnowledgeFileIO.ListAllKnowledges EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1292] modKnowledgeFileIO.ListAllKnowledges EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set ListAllKnowledges = New Collection
End Function

' ListKnowledgesByFormat: filter by FormatID with name-prefix fast path.
Public Function ListKnowledgesByFormat(ByVal formatId As String) As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1293] modKnowledgeFileIO.ListKnowledgesByFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection
    Dim all As Collection
    Set all = ListAllKnowledges()
    Dim prefix As String
    prefix = formatId & "-"
    Dim id As Variant
    For Each id In all
        Dim sId As String
        sId = CStr(id)
        If Len(sId) >= Len(prefix) And Left$(sId, Len(prefix)) = prefix Then
            result.Add sId
        Else
            Dim d As Object
            Set d = LoadKnowledge(sId)
            If d.Exists("FormatID") Then
                If CStr(d("FormatID")) = formatId Then result.Add sId
            End If
        End If
    Next id
    Set ListKnowledgesByFormat = result
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1294] modKnowledgeFileIO.ListKnowledgesByFormat EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1295] modKnowledgeFileIO.ListKnowledgesByFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set ListKnowledgesByFormat = New Collection
End Function

' DeleteKnowledge: remove the .txt file. Returns True on success.
Public Function DeleteKnowledge(ByVal knowledgeNo As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1296] modKnowledgeFileIO.DeleteKnowledge ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath As String
    filePath = modConfigHolder.GetDataDir() & knowledgeNo & ".txt"
    If fso.FileExists(filePath) Then
        fso.DeleteFile filePath
        DeleteKnowledge = True
    Else
        DeleteKnowledge = False
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1297] modKnowledgeFileIO.DeleteKnowledge EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1298] modKnowledgeFileIO.DeleteKnowledge EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    DeleteKnowledge = False
End Function

' BackupKnowledgeFile: copy <knowledgeNo>.txt to backup_dir with ts suffix.
Public Function BackupKnowledgeFile(ByVal knowledgeNo As String, ByVal ts As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1299] modKnowledgeFileIO.BackupKnowledgeFile ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim srcPath As String
    srcPath = modConfigHolder.GetDataDir() & knowledgeNo & ".txt"
    If Not fso.FileExists(srcPath) Then
        BackupKnowledgeFile = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1300] modKnowledgeFileIO.BackupKnowledgeFile EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim bdir As String
    bdir = modConfigHolder.GetBackupDir()
    If Not fso.FolderExists(bdir) Then fso.CreateFolder bdir
    Dim dstPath As String
    dstPath = bdir & knowledgeNo & "_" & ts & ".txt"
    fso.CopyFile srcPath, dstPath, True
    BackupKnowledgeFile = True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1301] modKnowledgeFileIO.BackupKnowledgeFile EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1302] modKnowledgeFileIO.BackupKnowledgeFile EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    BackupKnowledgeFile = False
End Function

' CleanupOldBackups (Q34): delete *.txt in backup_dir older than 90 days.
' Returns deleted count; -1 on error. (impl 2026-05-19)

' CleanupOldBackups (Q34): delete *.txt in backup_dir older than 90 days.
' Returns deleted count; -1 on error. (impl 2026-05-19)
Public Function CleanupOldBackups() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1303] modKnowledgeFileIO.CleanupOldBackups ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim backupDir As String
    backupDir = modConfigHolder.GetBackupDir()
    If Not fso.FolderExists(backupDir) Then
        CleanupOldBackups = 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1304] modKnowledgeFileIO.CleanupOldBackups EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim threshold As Date
    threshold = DateAdd("d", -90, Now())
    Dim deleted As Long
    deleted = 0
    Dim folder As Object
    Set folder = fso.GetFolder(backupDir)
    Dim f As Object
    For Each f In folder.Files
        If LCase(fso.GetExtensionName(f.Name)) = "txt" Then
            If f.DateLastModified < threshold Then
                fso.DeleteFile f.Path, True
                deleted = deleted + 1
            End If
        End If
    Next f
    CleanupOldBackups = deleted
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1305] modKnowledgeFileIO.CleanupOldBackups EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1306] modKnowledgeFileIO.CleanupOldBackups EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    CleanupOldBackups = -1
End Function

' ----------------------------------------------------------------
' v2.2: reflect-folder backup / restore (M-12 field reflection)
'   A migration run backs up its pre-reflection knowledge files into
'   backup_dir\<reflectDir>\ (reflectDir = reflect_YYYYMMDD_HHMM,
'   one folder per run; screen_design_v2 2.11.4). Btn_RestoreBackup
'   restores the latest reflect_* folder back into data_dir by an
'   overwrite copy. These are plain file-system copies (no delete),
'   matching the M-12 special exception that allows the admin book to
'   write data\ (screen_design_v2 2.11.8).
' ----------------------------------------------------------------

' BackupKnowledgeToReflect: copy data_dir\<knowledgeNo>.txt into
' backup_dir\<reflectDir>\ . Creates backup_dir and the reflect
' subfolder when missing. Returns True on success.
Public Function BackupKnowledgeToReflect(ByVal knowledgeNo As String, _
                                         ByVal reflectDir As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1307] modKnowledgeFileIO.BackupKnowledgeToReflect ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim srcPath As String
    srcPath = modConfigHolder.GetDataDir() & knowledgeNo & ".txt"
    If Not fso.FileExists(srcPath) Then
        BackupKnowledgeToReflect = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1308] modKnowledgeFileIO.BackupKnowledgeToReflect EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim bdir As String
    bdir = modConfigHolder.GetBackupDir()
    If Not fso.FolderExists(bdir) Then fso.CreateFolder bdir
    Dim rdir As String
    rdir = bdir & reflectDir & "\"
    If Not fso.FolderExists(rdir) Then fso.CreateFolder rdir
    fso.CopyFile srcPath, rdir & knowledgeNo & ".txt", True
    BackupKnowledgeToReflect = True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1309] modKnowledgeFileIO.BackupKnowledgeToReflect EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1310] modKnowledgeFileIO.BackupKnowledgeToReflect EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    BackupKnowledgeToReflect = False
End Function

' GetLatestReflectFolder: return the name of the most recent reflect_*
' subfolder under backup_dir (lexically greatest, since the names carry
' a timestamp). Returns "" when no reflect_* folder exists.
Public Function GetLatestReflectFolder() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1311] modKnowledgeFileIO.GetLatestReflectFolder ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim bdir As String
    bdir = modConfigHolder.GetBackupDir()
    GetLatestReflectFolder = ""
    If Not fso.FolderExists(bdir) Then Exit Function
    Dim latest As String
    latest = ""
    Dim subFolder As Object
    For Each subFolder In fso.GetFolder(bdir).SubFolders
        If LCase(Left$(subFolder.Name, Len(REFLECT_PREFIX))) = REFLECT_PREFIX Then
            If subFolder.Name > latest Then latest = subFolder.Name
        End If
    Next subFolder
    GetLatestReflectFolder = latest
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1312] modKnowledgeFileIO.GetLatestReflectFolder EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1313] modKnowledgeFileIO.GetLatestReflectFolder EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    GetLatestReflectFolder = ""
End Function

' RestoreKnowledgeFromReflect: copy every *.txt in backup_dir\<reflectDir>\
' back into data_dir (overwrite copy; nothing is deleted). Returns the
' restored file count, or -1 on error / missing folder.
Public Function RestoreKnowledgeFromReflect(ByVal reflectDir As String) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1314] modKnowledgeFileIO.RestoreKnowledgeFromReflect ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim rdir As String
    rdir = modConfigHolder.GetBackupDir() & reflectDir & "\"
    If Not fso.FolderExists(rdir) Then
        RestoreKnowledgeFromReflect = -1
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1315] modKnowledgeFileIO.RestoreKnowledgeFromReflect EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim ddir As String
    ddir = modConfigHolder.GetDataDir()
    If Not fso.FolderExists(ddir) Then fso.CreateFolder ddir
    Dim restored As Long
    restored = 0
    Dim f As Object
    For Each f In fso.GetFolder(rdir).Files
        If LCase(fso.GetExtensionName(f.Name)) = "txt" Then
            fso.CopyFile f.Path, ddir & f.Name, True
            restored = restored + 1
        End If
    Next f
    RestoreKnowledgeFromReflect = restored
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1316] modKnowledgeFileIO.RestoreKnowledgeFromReflect EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1317] modKnowledgeFileIO.RestoreKnowledgeFromReflect EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    RestoreKnowledgeFromReflect = -1
End Function

' ADR-0006/0090/0094 JP literal removal:
Private Property Get TOUROKU_XLSM_NAME() As String
    TOUROKU_XLSM_NAME = ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63) & ".xlsm"
End Property

' Reject control characters (Chr(0)-Chr(31)) except Tab(9), LF(10), CR(13).
' 2026-06-06: added for ADR-0112 validation strengthening.
Private Function IsControlCharSafe(ByVal s As String) As Boolean
    On Error GoTo ErrHandler
    Dim i As Long
    For i = 1 To Len(s)
        Dim a As Long: a = AscW(Mid(s, i, 1))
        If a >= 0 And a < 32 Then
            If a <> 9 And a <> 10 And a <> 13 Then
                IsControlCharSafe = False
                Exit Function
            End If
        End If
    Next i
    IsControlCharSafe = True
    Exit Function
ErrHandler:
    IsControlCharSafe = False
End Function
```
