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
' モジュール: modFormatLoader（Phase 4 / ユーティリティ層）
' 概要:       フォーマット定義 .txt の読み書き・列挙・削除
'             write/delete は管理.xlsm のみ（ThisWorkbook.Name で enforce、Q19）
'             ADR-0053 第2.3 / 第2.9 を実装
'             v2_format_stanza_schema.md に準拠（[FORMAT]+[FIELD]+[MIGRATE_RULE]、key=value 形式）
' Version:    v2.1（2026-05-16 EOD、Q1-Q57 全件反映）
' 依存:       modStanzaIO（汎用 [SECTION]+key=value parser）, modConfigHolder, modKnowledgeFileIO（Q55 整合性 check 用）, ClsStanzaSection
' 関連:       Q6（format schema 確定、5 種 FieldType）
'             Q19（公開 I/F 5 関数確定 + 管理.xlsm のみ write）
'             Q43（M-07 廃止、M-08 で全機能 + 編集機能集約）
'             Q50（format 保存時に既存 knowledge を自動 migrate）
'             Q55（format 削除前に knowledge 整合性 check、1 件でも残れば reject）
' ================================================================
Option Explicit

' xlsm 名 enforce 用（管理.xlsm のみが write/delete を許可）
' iter18b: KANRI_XLSM_NAME via ChrW (the CP932 literal was mojibake-corrupted
' through an Edit/Write round-trip, leading SaveFormat to reject every write
' against the in-process workbook because the comparison always failed).
Private Function KANRI_XLSM_NAME() As String
    KANRI_XLSM_NAME = ChrW(&H7BA1) & ChrW(&H7406) & ".xlsm"
End Function

' ----------------------------------------------------------------
' Public I/F（Q19 確定、5 関数）
' ----------------------------------------------------------------

' ================================================================
' 関数名: LoadFormat
' 概要:   formatId を受けて format .txt を読み込み、ClsStanzaSection の Collection を返す
' 引数:   ByVal formatId As String - 例 "TICKET"
' 戻り値: Collection - 要素は ClsStanzaSection ([FORMAT], [FIELD] N 件, [MIGRATE_RULE] N 件)
'         ファイル不在時は空 Collection
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
' 関数名: LoadFormatList
' 概要:   format_dir 配下の全 format について FormatID + Description を返す
'         M-02 ドロップダウン（Q28）用、ListAllFormats との違いは Description を付与する点
' 戻り値: Collection - 要素は Dictionary {FormatID, Description}
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
' Function: ListAllFormats - format_dir 配下の全 *.txt の basename を返す
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

' SaveFormat: 管理.xlsm のみ enforce (Q19) のうえ modStanzaIO.WriteStanzaFile で書き出し
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
