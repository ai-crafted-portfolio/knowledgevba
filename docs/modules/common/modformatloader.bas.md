---
title: modFormatLoader.bas
description: modFormatLoader.bas のソースコード（コピペ用）
---

# modFormatLoader.bas

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modFormatLoader.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modFormatLoader"
' ================================================================
' ・ｽ・ｽ・ｽW・ｽ・ｽ・ｽ[・ｽ・ｽ: modFormatLoader・ｽiPhase 4 / ・ｽ・ｽ・ｽ[・ｽe・ｽB・ｽ・ｽ・ｽe・ｽB・ｽw・ｽj
' ・ｽT・ｽv:       ・ｽt・ｽH・ｽ[・ｽ}・ｽb・ｽg・ｽ・ｽ` .txt ・ｽﾇ搾ｿｽ・ｽE・ｽ・ｽ・ｽ・ｽ・ｽE・ｽ尞・
'             write/delete ・ｽﾍ管暦ｿｽ.xlsm ・ｽﾌみ（ThisWorkbook.Name ・ｽ・ｽ enforce・ｽAQ19・ｽj
'             ADR-0053 ・ｽ・ｽ2.3 / ・ｽ・ｽ2.9 ・ｽ・ｽ・ｽ・ｽ
'             v2_format_stanza_schema.md ・ｽ・ｽ・ｽ・ｽ・ｽi[FORMAT]+[FIELD]+[MIGRATE_RULE]・ｽAkey=value ・ｽ`・ｽ・ｽ・ｽj
' Version:    v2.1・ｽi2026-05-16 EOD・ｽAQ1-Q57 ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽ・ｽf・ｽj
' ・ｽﾋ托ｿｽ・ｽ・ｽ:     modStanzaIO・ｽi・ｽﾄ用 [SECTION]+key=value parser・ｽj, modConfigHolder, modKnowledgeFileIO・ｽiQ55 ・ｽ尞懶ｿｽO check ・ｽp・ｽj, ClsStanzaSection
' ・ｽﾖ連:       Q6・ｽiformat schema ・ｽm・ｽ・ｽA5 ・ｽ・ｽ FieldType・ｽj
'             Q19・ｽi・ｽ・ｽ・ｽJ I/F 5 ・ｽﾖ撰ｿｽ・ｽm・ｽ・ｽ + ・ｽﾇ暦ｿｽ.xlsm ・ｽ・ｽ・ｽ・ｽj
'             Q43・ｽiM-07 ・ｽ尞懶ｿｽAM-08 ・ｽﾅ全・ｽ・ｽ・ｽ・ｽ + ・ｽ尞懶ｿｽ@・ｽ\・ｽW・ｽ・ｽj
'             Q50・ｽiformat ・ｽﾛ托ｿｽ・ｽ・ｽ・ｽﾉ奇ｿｽ・ｽ・ｽ knowledge ・ｽ・ｽ・ｽ・ｽ migrate・ｽj
'             Q55・ｽiformat ・ｽ尞懶ｿｽO・ｽ・ｽ knowledge ・ｽ・ｽ・ｽ・ｽ check・ｽA1 ・ｽ・ｽ・ｽﾅゑｿｽ・ｽ・ｽ・ｽ・ｽ・ｽ reject・ｽj
' ================================================================
Option Explicit

' xlsm ・ｽ・ｽ enforce ・ｽp・ｽi・ｽﾇ暦ｿｽ.xlsm ・ｽ・ｽ・ｽ・ｽ・ｽ write/delete・ｽj
' iter18b: KANRI_XLSM_NAME via ChrW (the CP932 literal was mojibake-corrupted
' through an Edit/Write round-trip, leading SaveFormat to reject every write
' against the in-process workbook because the comparison always failed).
Private Function KANRI_XLSM_NAME() As String
    KANRI_XLSM_NAME = ChrW(&H7BA1) & ChrW(&H7406) & ".xlsm"
End Function

' ----------------------------------------------------------------
' Public I/F・ｽiQ19 ・ｽm・ｽ・ｽA5 ・ｽﾖ撰ｿｽ・ｽj
' ----------------------------------------------------------------

' ================================================================
' ・ｽﾖ撰ｿｽ・ｽ・ｽ: LoadFormat
' ・ｽT・ｽv:   formatId ・ｽ・ｽ・ｽ・ｽ format .txt ・ｽ・ｽﾇ搾ｿｽ・ｽAClsStanzaSection ・ｽ・ｽ Collection ・ｽ・ｽﾔゑｿｽ
' ・ｽ・ｽ・ｽ・ｽ:   ByVal formatId As String - ・ｽ・ｽ "TICKET"
' ・ｽﾟゑｿｽl: Collection - ・ｽv・ｽf・ｽ・ｽ ClsStanzaSection ([FORMAT], [FIELD] N ・ｽ・ｽ, [MIGRATE_RULE] N ・ｽ・ｽ)
'         ・ｽt・ｽ@・ｽC・ｽ・ｽ・ｽs・ｽﾝ趣ｿｽ・ｽﾍ具ｿｽ Collection
' ================================================================
Public Function LoadFormat(ByVal formatId As String) As Collection
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = modConfigHolder.GetFormatDir() & formatId & ".txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Set LoadFormat = New Collection
        Exit Function
    End If

    Set LoadFormat = modStanzaIO.ParseStanzaFile(filePath)
    Exit Function

ErrHandler:
    LogError "modFormatLoader", "LoadFormat", "formatId=" & formatId & " Err=" & Err.Description
    Set LoadFormat = New Collection
End Function

' ================================================================
' ・ｽﾖ撰ｿｽ・ｽ・ｽ: LoadFormatList
' ・ｽT・ｽv:   format_dir ・ｽz・ｽ・ｽ・ｽﾌ全 format ・ｽﾉつゑｿｽ・ｽ・ｽ FormatID + Description ・ｽ・ｽﾔゑｿｽ
'         M-02 ・ｽh・ｽ・ｽ・ｽb・ｽv・ｽ_・ｽE・ｽ・ｽ・ｽiQ28・ｽj・ｽp・ｽAListAllFormats ・ｽﾆの違い・ｽ・ｽ Description ・ｽ・ｽ・ｽ・ｽ
' ・ｽﾟゑｿｽl: Collection - ・ｽv・ｽf・ｽ・ｽ Dictionary {FormatID, Description}
' ================================================================
Public Function LoadFormatList() As Collection
    On Error GoTo ErrHandler

    Dim result As Collection
    Set result = New Collection

    Dim allIds As Collection
    Set allIds = ListAllFormats()

    Dim id As Variant
    For Each id In allIds
        Dim sections As Collection
        Set sections = LoadFormat(CStr(id))

        Dim desc As String
        desc = ExtractDescription(sections)

        Dim entry As Object
        Set entry = CreateObject("Scripting.Dictionary")
        entry("FormatID") = CStr(id)
        entry("Description") = desc

        result.Add entry
    Next id

    Set LoadFormatList = result
    Exit Function

ErrHandler:
    LogError "modFormatLoader", "LoadFormatList", "Err=" & Err.Description
    Set LoadFormatList = New Collection
End Function

' ================================================================
' Function: ListAllFormats - format_dir ・ｽz・ｽ・ｽ・ｽﾌ全 *.txt ・ｽ・ｽ basename ・ｽ・ｽﾔゑｿｽ
' ================================================================
Public Function ListAllFormats() As Collection
    On Error GoTo ErrHandler
    Dim result As New Collection
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dir As String: dir = modConfigHolder.GetFormatDir()
    If Not fso.FolderExists(dir) Then
        Set ListAllFormats = result
        Exit Function
    End If
    Dim folder As Object: Set folder = fso.GetFolder(dir)
    Dim f As Object
    For Each f In folder.Files
        If LCase(fso.GetExtensionName(f.Name)) = "txt" Then
            result.Add fso.GetBaseName(f.Name)
        End If
    Next f
    Set ListAllFormats = result
    Exit Function
ErrHandler:
    Set ListAllFormats = New Collection
End Function

' Extract Description from sections (looks in [FORMAT] section)
Private Function ExtractDescription(ByVal sections As Collection) As String
    If sections Is Nothing Then Exit Function
    If sections.Count = 0 Then Exit Function
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To sections.Count
        Set sec = sections.Item(i)
        If sec.SectionName = "FORMAT" Then
            If sec.HasKey("Description") Then ExtractDescription = sec.GetValue("Description")
            Exit Function
        End If
    Next i
End Function

' FormatExists: check if .txt exists
Public Function FormatExists(ByVal formatId As String) As Boolean
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    FormatExists = fso.FileExists(modConfigHolder.GetFormatDir() & formatId & ".txt")
End Function

' SaveFormat: enforce ・ｽﾇ暦ｿｽ.xlsm only (Q19), then write via modStanzaIO.WriteStanzaFile
Public Function SaveFormat(ByVal formatId As String, ByVal sections As Collection) As Long
    On Error GoTo ErrHandler
    ' Q19 enforce: xlsmName check
    Dim wbName As String
    wbName = modConfigHolder.GetValueOrDefault("xlsmName", "")
    If LCase(wbName) <> "kanri" Then
        Dim actualName As String
        actualName = ThisWorkbook.Name
        If actualName <> KANRI_XLSM_NAME And LCase(wbName) <> "kanri" Then
            SaveFormat = 2 ' rejected by Q19
            Exit Function
        End If
    End If
    Dim filePath As String
    filePath = modConfigHolder.GetFormatDir() & formatId & ".txt"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(modConfigHolder.GetFormatDir()) Then
        fso.CreateFolder modConfigHolder.GetFormatDir()
    End If
    modStanzaIO.WriteStanzaFile filePath, sections
    SaveFormat = 0
    Exit Function
ErrHandler:
    SaveFormat = 3
End Function

' DeleteFormat: enforce + delete file
' Q55: reject (ret=2) if any knowledge of this format still exists.
' Return codes: 0=success / 1=error / 2=rejected by Q55
Public Function DeleteFormat(ByVal formatId As String) As Long
    On Error GoTo ErrHandler
    ' Q55 check: any knowledge using this format?
    Dim usingKnw As Collection
    Set usingKnw = modKnowledgeFileIO.ListKnowledgesByFormat(formatId)
    If Not usingKnw Is Nothing Then
        If usingKnw.Count > 0 Then
            DeleteFormat = 2  ' Q55 reject
            Exit Function
        End If
    End If
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath As String
    filePath = modConfigHolder.GetFormatDir() & formatId & ".txt"
    If fso.FileExists(filePath) Then fso.DeleteFile filePath
    DeleteFormat = 0
    Exit Function
ErrHandler:
    DeleteFormat = 1
End Function

' Private LogError stub
Private Sub LogError(ByVal moduleName As String, ByVal funcName As String, ByVal msg As String)
    Debug.Print "[ERR] " & moduleName & "." & funcName & ": " & msg
End Sub
```
