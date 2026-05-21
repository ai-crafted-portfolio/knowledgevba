---
title: modKnowledgeFileIO.bas
---

# modKnowledgeFileIO.bas

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 3 ブック共通 |
| 役割 | ナレッジ .txt の読み書き・バックアップ（書き込みは登録修正.xlsm のみ） |
| 行数 | 379 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modKnowledgeFileIO.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
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
Private Const TOUROKU_XLSM_NAME As String = "登録修正.xlsm"

' ----------------------------------------------------------------
' Public I/F: read
' ----------------------------------------------------------------
Public Function LoadKnowledge(ByVal knowledgeNo As String) As Object
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = modConfigHolder.GetDataDir() & knowledgeNo & ".txt"

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Set LoadKnowledge = result
        Exit Function
    End If

    Dim content As String
    content = ReadAllTextShiftJIS(filePath)

    Set result = ParseKnowledgeStanza(content)
    Set LoadKnowledge = result
    Exit Function

ErrHandler:
    Set LoadKnowledge = CreateObject("Scripting.Dictionary")
End Function

' ================================================================
' GetKnowledgeTimestamp (Q46 optimistic lock)
' ================================================================
Public Function GetKnowledgeTimestamp(ByVal knowledgeNo As String) As Date
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = modConfigHolder.GetDataDir() & knowledgeNo & ".txt"

    If Len(Dir(filePath)) > 0 Then
        GetKnowledgeTimestamp = FileDateTime(filePath)
    Else
        GetKnowledgeTimestamp = 0
    End If
    Exit Function

ErrHandler:
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
    On Error GoTo ErrHandler

    ' xlsm write gate (architecture 3.6): knowledge save is permitted
    ' only from the register workbook. Symmetric with SaveFormat (Q19).
    ' Return 3 = rejected by xlsm gate (0=ok / 1=conflict / 2=CP932).
    If Not IsRegisterWorkbook() Then
        SaveKnowledge = 3
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
                Exit Function
            End If
        End If
    End If

    Dim content As String
    content = SerializeKnowledgeStanza(knowledgeDict)

    If Not IsCP932Compatible(content) Then
        SaveKnowledge = 2  ' CP932 incompatible
        Exit Function
    End If

    WriteAllTextShiftJIS filePath, content
    SaveKnowledge = 0
    Exit Function

ErrHandler:
    SaveKnowledge = 2
End Function

' ================================================================
' Private helpers
' ================================================================

' Q47 ###xxx### format serializer
Private Function SerializeKnowledgeStanza(ByVal d As Object) As String
    Dim sb As String
    sb = ""
    Dim k As Variant
    For Each k In d.Keys
        sb = sb & "###" & CStr(k) & "###" & vbCrLf
        sb = sb & CStr(d(k)) & vbCrLf
    Next k
    SerializeKnowledgeStanza = sb
End Function

' xlsm write gate (architecture 3.6): True only when the caller is the
' register workbook. Mirrors modFormatLoader.SaveFormat Q19 enforce.
' Decision order:
'   1. config token xlsmName (register / touroku) -> allowed
'   2. config token xlsmName set to any other value -> rejected
'   3. xlsmName unset -> fall back to ThisWorkbook.Name match
Private Function IsRegisterWorkbook() As Boolean
    On Error GoTo Bad
    Dim wbName As String
    wbName = LCase(modConfigHolder.GetValueOrDefault("xlsmName", ""))
    If wbName = "register" Or wbName = "touroku" Then
        IsRegisterWorkbook = True
        Exit Function
    End If
    If Len(wbName) > 0 Then
        ' xlsmName is explicitly set to a non-register role -> reject
        IsRegisterWorkbook = False
        Exit Function
    End If
    ' xlsmName unset: identify by the running workbook file name
    IsRegisterWorkbook = (ThisWorkbook.Name = TOUROKU_XLSM_NAME)
    Exit Function
Bad:
    IsRegisterWorkbook = False
End Function

' Check if string is fully CP932-encodable round-trip
Private Function IsCP932Compatible(ByVal s As String) As Boolean
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
    Exit Function
Bad:
    IsCP932Compatible = False
End Function

' Write text as Shift_JIS (CP932) with CRLF
Private Sub WriteAllTextShiftJIS(ByVal filePath As String, ByVal content As String)
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2
    ado.Charset = "Shift_JIS"
    ado.Open
    ado.WriteText content
    ado.SaveToFile filePath, 2
    ado.Close
End Sub

' Read text from Shift_JIS file
Private Function ReadAllTextShiftJIS(ByVal filePath As String) As String
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2
    ado.Charset = "Shift_JIS"
    ado.Open
    ado.LoadFromFile filePath
    ReadAllTextShiftJIS = ado.ReadText
    ado.Close
End Function

' Parse ###Key### / value blocks into Dictionary
Private Function ParseKnowledgeStanza(ByVal content As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
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
End Function

' ================================================================
' v2.1 minimal impl (2026-05-19): replaces empty stubs so legacy
' E2E tests (_e2e_touroku / _e2e_kensaku / _e2e_kanri) can verify
' real semantics. Scope: enumerate data_dir for *.txt knowledge files.
' ================================================================

' ListAllKnowledges: return Collection of knowledgeNo (file basename minus .txt)
Public Function ListAllKnowledges() As Collection
    On Error GoTo ErrHandler
    Dim result As Collection
    Set result = New Collection
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dir As String
    dir = modConfigHolder.GetDataDir()
    If Not fso.FolderExists(dir) Then
        Set ListAllKnowledges = result
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
    Exit Function
ErrHandler:
    Set ListAllKnowledges = New Collection
End Function

' ListKnowledgesByFormat: filter by FormatID with name-prefix fast path.
Public Function ListKnowledgesByFormat(ByVal formatId As String) As Collection
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
    Exit Function
ErrHandler:
    Set ListKnowledgesByFormat = New Collection
End Function

' DeleteKnowledge: remove the .txt file. Returns True on success.
Public Function DeleteKnowledge(ByVal knowledgeNo As String) As Boolean
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
    Exit Function
ErrHandler:
    DeleteKnowledge = False
End Function

' BackupKnowledgeFile: copy <knowledgeNo>.txt to backup_dir with ts suffix.
Public Function BackupKnowledgeFile(ByVal knowledgeNo As String, ByVal ts As String) As Boolean
    On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim srcPath As String
    srcPath = modConfigHolder.GetDataDir() & knowledgeNo & ".txt"
    If Not fso.FileExists(srcPath) Then
        BackupKnowledgeFile = False
        Exit Function
    End If
    Dim bdir As String
    bdir = modConfigHolder.GetBackupDir()
    If Not fso.FolderExists(bdir) Then fso.CreateFolder bdir
    Dim dstPath As String
    dstPath = bdir & knowledgeNo & "_" & ts & ".txt"
    fso.CopyFile srcPath, dstPath, True
    BackupKnowledgeFile = True
    Exit Function
ErrHandler:
    BackupKnowledgeFile = False
End Function

' CleanupOldBackups (Q34): delete *.txt in backup_dir older than 90 days.
' Returns deleted count; -1 on error. (impl 2026-05-19)

' CleanupOldBackups (Q34): delete *.txt in backup_dir older than 90 days.
' Returns deleted count; -1 on error. (impl 2026-05-19)
Public Function CleanupOldBackups() As Long
    On Error GoTo ErrHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim backupDir As String
    backupDir = modConfigHolder.GetBackupDir()
    If Not fso.FolderExists(backupDir) Then
        CleanupOldBackups = 0
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
    Exit Function
ErrHandler:
    CleanupOldBackups = -1
End Function
```
