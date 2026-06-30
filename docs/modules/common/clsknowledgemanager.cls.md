---
title: clsKnowledgeManager.cls
description: clsKnowledgeManager.cls のソースコード（コピペ用）
---

# clsKnowledgeManager.cls

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: クラスモジュール
**更新日**: 2026-06-30 14:44 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsKnowledgeManager.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "clsKnowledgeManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' Class:   clsKnowledgeManager (v2.3, BUG-4 fix)
' Summary: knowledge .txt CRUD (v2.1 ###xxx### format) + optimistic lock
'          + registration form build / load-for-edit / update wrappers
'          for modEntryKnowledge.bas caller signatures.
' Version: v2.3 (2026-05-30, BUG-4 caller API alignment)
' Deps:    modKnowledgeFileIO (Q47 new format),
'          modFormatLoader (Q19),
'          clsFieldMigrator (Q50 sync migrate)
'          ADR-0053 2.3 / 2.9 / 2.10
' ================================================================
Option Explicit

Private m_logger As Object        ' clsLogger or Nothing
Private m_formatMgr As Object     ' clsFormatManager or Nothing
Private m_dataFolder As String    ' optional override for data dir; empty = use modConfigHolder.GetDataDir()

' ----------------------------------------------------------------
' Init
' ----------------------------------------------------------------

' BUG-4 fix: caller modEntryKnowledge.bas calls
'   knwMgr.Init logger, formatMgr, GetDataFolder()
' All three args are Optional so the legacy 1-arg call still works.
Public Sub Init( _
    Optional ByVal logger As Object = Nothing, _
    Optional ByVal formatMgr As Object = Nothing, _
    Optional ByVal dataFolder As String = "")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0500] clsKnowledgeManager.Init ENTER"  ' [ADR-0100]
    Set m_logger = logger
    Set m_formatMgr = formatMgr
    m_dataFolder = dataFolder
End Sub

' ----------------------------------------------------------------
' Public API: Load
' ----------------------------------------------------------------

Public Function LoadKnowledge( _
    ByVal knowledgeNo As String, _
    Optional ByRef originalTimestamp As Date _
) As Object
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0501] clsKnowledgeManager.LoadKnowledge ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Set LoadKnowledge = modKnowledgeFileIO.LoadKnowledge(knowledgeNo)
    originalTimestamp = modKnowledgeFileIO.GetKnowledgeTimestamp(knowledgeNo)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0502] clsKnowledgeManager.LoadKnowledge EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0503] clsKnowledgeManager.LoadKnowledge EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "LoadKnowledge", Err.Description, "", "LOG-KNW-LOAD-ERR"
    End If
    Set LoadKnowledge = CreateObject("Scripting.Dictionary")
End Function

' ----------------------------------------------------------------
' Public API: Number issue (Q33, <FormatID>-<4 digit sequence>)
' ----------------------------------------------------------------

Public Function BuildKnowledgeNumber(ByVal formatId As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0504] clsKnowledgeManager.BuildKnowledgeNumber ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    ' 2026-06-06: reject non-existent format -- production must not mint numbers for missing formats.
    If Not modFormatLoader.FormatExists(formatId) Then
        BuildKnowledgeNumber = ""
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0504b] clsKnowledgeManager.BuildKnowledgeNumber EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim existingList As Collection
    Set existingList = modKnowledgeFileIO.ListKnowledgesByFormat(formatId)

    Dim maxSeq As Long
    maxSeq = 0
    Dim knwNo As Variant
    For Each knwNo In existingList
        Dim seq As Long
        seq = ExtractSequence(CStr(knwNo), formatId)
        If seq > maxSeq Then maxSeq = seq
    Next knwNo

    Dim nextSeq As Long
    nextSeq = maxSeq + 1
    If nextSeq > 9999 Then
        Err.Raise vbObjectError + 200, "clsKnowledgeManager", _
                  "KnowledgeNo overflow: format " & formatId
    End If

    BuildKnowledgeNumber = formatId & "-" & Format$(nextSeq, "0000")
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0505] clsKnowledgeManager.BuildKnowledgeNumber EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0506] clsKnowledgeManager.BuildKnowledgeNumber EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "BuildKnowledgeNumber", Err.Description, "", "LOG-KNW-BLDNO-ERR"
    End If
    BuildKnowledgeNumber = ""
End Function

Private Function ExtractSequence(ByVal knwNo As String, ByVal formatId As String) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0507] clsKnowledgeManager.ExtractSequence ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim prefix As String
    prefix = formatId & "-"
    If Left(knwNo, Len(prefix)) <> prefix Then
        ExtractSequence = 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0508] clsKnowledgeManager.ExtractSequence EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim seqStr As String
    seqStr = Mid(knwNo, Len(prefix) + 1)
    ExtractSequence = CLng(seqStr)
    If Err.Number <> 0 Then
        ExtractSequence = 0
        Err.Clear
    End If
End Function

' ----------------------------------------------------------------
' Public API: Save (Q46 optimistic lock)
' ----------------------------------------------------------------

Public Function SaveKnowledge( _
    ByVal knowledgeNo As String, _
    ByVal knowledgeDict As Object, _
    ByVal originalTimestamp As Date _
) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0509] clsKnowledgeManager.SaveKnowledge ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    SaveKnowledge = modKnowledgeFileIO.SaveKnowledge(knowledgeNo, knowledgeDict, originalTimestamp)

    If SaveKnowledge = 0 And Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "SaveKnowledge", "saved: " & knowledgeNo, knowledgeNo, "LOG-KNW-SAVE-OK"
    ElseIf SaveKnowledge = 1 And Not m_logger Is Nothing Then
        m_logger.LogWarn "clsKnowledgeManager", "SaveKnowledge", "conflict: " & knowledgeNo, knowledgeNo, "LOG-KNW-CONFLICT"
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0510] clsKnowledgeManager.SaveKnowledge EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0511] clsKnowledgeManager.SaveKnowledge EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "SaveKnowledge", Err.Description, "", "LOG-KNW-SAVE-ERR"
    End If
    SaveKnowledge = 2
End Function

' ----------------------------------------------------------------
' Public API: Delete (Q45 M-03)
' ----------------------------------------------------------------

Public Function DeleteKnowledge(ByVal knowledgeNo As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0512] clsKnowledgeManager.DeleteKnowledge ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim ts As String
    ts = Format$(Now(), "yyyymmdd_hhnnss")
    If Not modKnowledgeFileIO.BackupKnowledgeFile(knowledgeNo, ts) Then
        If Not m_logger Is Nothing Then
            m_logger.LogError "clsKnowledgeManager", "DeleteKnowledge", _
                              "backup failed: " & knowledgeNo, knowledgeNo, "LOG-KNW-DEL-NOTFOUND"
        End If
        DeleteKnowledge = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0513] clsKnowledgeManager.DeleteKnowledge EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    DeleteKnowledge = modKnowledgeFileIO.DeleteKnowledge(knowledgeNo)

    If DeleteKnowledge And Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "DeleteKnowledge", _
                         "deleted: " & knowledgeNo & " (backup_ts=" & ts & ")", _
                         knowledgeNo, "LOG-KNW-DEL-OK"
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0514] clsKnowledgeManager.DeleteKnowledge EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0515] clsKnowledgeManager.DeleteKnowledge EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "DeleteKnowledge", Err.Description, "", "LOG-KNW-DEL-ERR"
    End If
    DeleteKnowledge = False
End Function

' ----------------------------------------------------------------
' Public API: Lists
' ----------------------------------------------------------------

Public Function ListAllKnowledges() As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0516] clsKnowledgeManager.ListAllKnowledges ENTER"  ' [ADR-0100]
    Set ListAllKnowledges = modKnowledgeFileIO.ListAllKnowledges()
End Function

Public Function ListKnowledgesByFormat(ByVal formatId As String) As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0517] clsKnowledgeManager.ListKnowledgesByFormat ENTER"  ' [ADR-0100]
    Set ListKnowledgesByFormat = modKnowledgeFileIO.ListKnowledgesByFormat(formatId)
End Function

' ================================================================
' BUG-4 fix: caller signature wrappers
' ================================================================

' ----------------------------------------------------------------
' BuildRegistrationForm
' Summary: load format definition and lay out the registration form
'          on the SHEET_KNW_SAVE sheet (label column, type column,
'          value column, etc.).
' Caller:  modEntryKnowledge.Btn_LoadKnowledgeFormat (no return value used)
' ----------------------------------------------------------------
Public Sub BuildRegistrationForm(ByVal formatId As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0518] clsKnowledgeManager.BuildRegistrationForm ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    If m_formatMgr Is Nothing Then
        Err.Raise vbObjectError + 210, "clsKnowledgeManager", _
                  "formatMgr not initialized (call Init logger, formatMgr, dataFolder)"
    End If

    Dim sections As Collection
    Set sections = m_formatMgr.LoadFormat(formatId)
    If sections Is Nothing Then Set sections = New Collection

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_SAVE)

    ' Header: format id at (KS_ROW_FMT_ID, KS_COL_FMT_ID_VAL) = (6, 3)
    ws.Cells(KS_ROW_FMT_ID, KS_COL_FMT_ID_VAL).Value = formatId

    ' Clear previous form rows (label/type/value)
    Dim i As Long
    For i = KS_FORM_START_ROW To KS_FORM_START_ROW + 200
        ws.Cells(i, KS_FIELD_COL_NAME).Value = ""
        ws.Cells(i, KS_FIELD_COL_NAME + 1).Value = ""   ' type column (D)
        ws.Cells(i, KS_FIELD_COL_VALUE).Value = ""
    Next i

    ' Emit FIELD sections
    Dim targetRow As Long
    targetRow = KS_FORM_START_ROW

    Dim sec As Object
    For Each sec In sections
        If sec.SectionName = "FIELD" Then
            Dim fieldName As String
            Dim fieldType As String
            fieldName = ""
            fieldType = ""
            If sec.HasKey("FieldName") Then fieldName = sec.GetValue("FieldName")
            If sec.HasKey("FieldType") Then fieldType = sec.GetValue("FieldType")
            ws.Cells(targetRow, KS_FIELD_COL_NAME).Value = fieldName
            ws.Cells(targetRow, KS_FIELD_COL_NAME + 1).Value = fieldType
            ws.Cells(targetRow, KS_FIELD_COL_VALUE).Value = ""
            targetRow = targetRow + 1
        End If
    Next sec

    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "BuildRegistrationForm", _
                         "fields=" & CStr(targetRow - KS_FORM_START_ROW), _
                         formatId, "LOG-KNW-BUILDFORM-OK"
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0519] clsKnowledgeManager.BuildRegistrationForm EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0520] clsKnowledgeManager.BuildRegistrationForm EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "BuildRegistrationForm", Err.Description, formatId, "LOG-KNW-BUILDFORM-ERR"
    End If
End Sub

' ----------------------------------------------------------------
' SaveNewKnowledge
' Summary: read field values from SHEET_KNW_SAVE, allocate new
'          knowledgeNo via BuildKnowledgeNumber, persist via
'          SaveKnowledge. Returns new knowledgeNo on success, ""
'          on failure (missing format id, empty form, save error).
' Caller:  modEntryKnowledge.Btn_SaveKnowledge
' ----------------------------------------------------------------
Public Function SaveNewKnowledge() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0521] clsKnowledgeManager.SaveNewKnowledge ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_SAVE)

    Dim formatId As String
    formatId = Trim(CStr(ws.Cells(KS_ROW_FMT_ID, KS_COL_FMT_ID_VAL).Value))
    If formatId = "" Then
        SaveNewKnowledge = ""
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0522] clsKnowledgeManager.SaveNewKnowledge EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    ' Collect fields from the form into a Dictionary keyed by FieldName.
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim anyField As Boolean
    anyField = False
    For i = KS_FORM_START_ROW To KS_FORM_START_ROW + 200
        Dim fName As String
        fName = Trim(CStr(ws.Cells(i, KS_FIELD_COL_NAME).Value))
        If fName = "" Then Exit For
        Dim fValue As String
        fValue = CStr(ws.Cells(i, KS_FIELD_COL_VALUE).Value)
        dict(fName) = fValue
        anyField = True
    Next i

    If Not anyField Then
        SaveNewKnowledge = ""
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0523] clsKnowledgeManager.SaveNewKnowledge EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim newId As String
    newId = BuildKnowledgeNumber(formatId)
    If newId = "" Then
        SaveNewKnowledge = ""
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0524] clsKnowledgeManager.SaveNewKnowledge EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim rc As Long
    rc = SaveKnowledge(newId, dict, CDate(0))
    If rc <> 0 Then
        SaveNewKnowledge = ""
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0525] clsKnowledgeManager.SaveNewKnowledge EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    SaveNewKnowledge = newId
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0526] clsKnowledgeManager.SaveNewKnowledge EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0527] clsKnowledgeManager.SaveNewKnowledge EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "SaveNewKnowledge", Err.Description, "", "LOG-KNW-SAVENEW-ERR"
    End If
    SaveNewKnowledge = ""
End Function

' ----------------------------------------------------------------
' LoadForEdit
' Summary: load knowledge by id and lay its fields into the edit
'          sheet (SHEET_KNW_EDIT). Returns True on success, False
'          if the knowledge cannot be loaded (not found / empty).
' Caller:  modEntryKnowledge.Btn_LoadKnowledge
' ----------------------------------------------------------------
Public Function LoadForEdit(ByVal knowledgeNo As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0528] clsKnowledgeManager.LoadForEdit ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim ts As Date
    Dim dict As Object
    Set dict = modKnowledgeFileIO.LoadKnowledge(knowledgeNo)
    ts = modKnowledgeFileIO.GetKnowledgeTimestamp(knowledgeNo)

    If dict Is Nothing Then
        LoadForEdit = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0529] clsKnowledgeManager.LoadForEdit EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If dict.Count = 0 Then
        LoadForEdit = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0530] clsKnowledgeManager.LoadForEdit EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)

    ' Knowledge No goes in the header cell at (KE_ROW_FMT_ID,
    ' KE_COL_KNW_NO) = (1, 3). The edit sheet form starts at the
    ' same row layout as the save sheet (form rows from row 8).
    ws.Cells(1, 3).Value = knowledgeNo

    Dim row As Long
    row = KS_FORM_START_ROW

    Dim k As Variant
    For Each k In dict.Keys
        ws.Cells(row, KS_FIELD_COL_NAME).Value = CStr(k)
        ws.Cells(row, KS_FIELD_COL_VALUE).Value = CStr(dict(k))
        row = row + 1
    Next k

    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsKnowledgeManager", "LoadForEdit", _
                         "loaded fields=" & CStr(dict.Count), _
                         knowledgeNo, "LOG-KNW-LOADEDIT-OK"
    End If
    LoadForEdit = True
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0531] clsKnowledgeManager.LoadForEdit EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0532] clsKnowledgeManager.LoadForEdit EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "LoadForEdit", Err.Description, knowledgeNo, "LOG-KNW-LOADEDIT-ERR"
    End If
    LoadForEdit = False
End Function

' ----------------------------------------------------------------
' UpdateKnowledge
' Summary: read current form values from SHEET_KNW_EDIT and write
'          them back to the given knowledgeNo. Returns True on
'          successful save (rc=0), False otherwise.
' Caller:  modEntryKnowledge.Btn_UpdateKnowledge
' ----------------------------------------------------------------
Public Function UpdateKnowledge(ByVal knowledgeNo As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0533] clsKnowledgeManager.UpdateKnowledge ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_EDIT)

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim anyField As Boolean
    anyField = False
    For i = KS_FORM_START_ROW To KS_FORM_START_ROW + 200
        Dim fName As String
        fName = Trim(CStr(ws.Cells(i, KS_FIELD_COL_NAME).Value))
        If fName = "" Then Exit For
        Dim fValue As String
        fValue = CStr(ws.Cells(i, KS_FIELD_COL_VALUE).Value)
        dict(fName) = fValue
        anyField = True
    Next i

    If Not anyField Then
        UpdateKnowledge = False
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0534] clsKnowledgeManager.UpdateKnowledge EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim origTs As Date
    origTs = modKnowledgeFileIO.GetKnowledgeTimestamp(knowledgeNo)

    Dim rc As Long
    rc = SaveKnowledge(knowledgeNo, dict, origTs)
    UpdateKnowledge = (rc = 0)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0535] clsKnowledgeManager.UpdateKnowledge EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0536] clsKnowledgeManager.UpdateKnowledge EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsKnowledgeManager", "UpdateKnowledge", Err.Description, knowledgeNo, "LOG-KNW-UPD-ERR"
    End If
    UpdateKnowledge = False
End Function
```
