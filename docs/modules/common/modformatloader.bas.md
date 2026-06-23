---
title: modFormatLoader.bas
description: modFormatLoader.bas のソースコード（コピペ用）
---

# modFormatLoader.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-23 09:12 JST

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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1238] modFormatLoader.KANRI_XLSM_NAME ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1239] modFormatLoader.LoadFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = modConfigHolder.GetFormatDir() & formatId & ".txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Set LoadFormat = New Collection
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1240] modFormatLoader.LoadFormat EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Set LoadFormat = modStanzaIO.ParseStanzaFile(filePath)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1241] modFormatLoader.LoadFormat EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1242] modFormatLoader.LoadFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1243] modFormatLoader.LoadFormatList ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim result As Collection
    Set result = New Collection

    Dim allIds As Collection
    Set allIds = ListAllFormats()

    Dim id As Variant
    For Each id In allIds
        Dim sections As Collection
        Set sections = LoadFormat(CStr(id))

        ' [USER-REQ 2026-06-09] Skip formats whose [FORMAT].Status = "無効".
        If IsFormatDisabled(sections) Then GoTo NextId

        Dim desc As String
        desc = ExtractDescription(sections)

        Dim entry As Object
        Set entry = CreateObject("Scripting.Dictionary")
        entry("FormatID") = CStr(id)
        entry("Description") = desc

        result.Add entry
NextId:
    Next id

    Set LoadFormatList = result
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1244] modFormatLoader.LoadFormatList EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1245] modFormatLoader.LoadFormatList EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    LogError "modFormatLoader", "LoadFormatList", "Err=" & Err.Description
    Set LoadFormatList = New Collection
End Function

' ================================================================
' Function: ListAllFormats - format_dir 配下の全 *.txt の basename を返す
' ================================================================
Public Function ListAllFormats() As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1246] modFormatLoader.ListAllFormats ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim result As New Collection
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dir As String: dir = modConfigHolder.GetFormatDir()
    If Not fso.FolderExists(dir) Then
        Set ListAllFormats = result
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1247] modFormatLoader.ListAllFormats EXIT-OK"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1248] modFormatLoader.ListAllFormats EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1249] modFormatLoader.ListAllFormats EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set ListAllFormats = New Collection
End Function

' Extract Description from sections (looks in [FORMAT] section)
Private Function ExtractDescription(ByVal sections As Collection) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1250] modFormatLoader.ExtractDescription ENTER"  ' [ADR-0100]
    If sections Is Nothing Then Exit Function
    If sections.Count = 0 Then Exit Function
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To sections.Count
        Set sec = sections.Item(i)
        If sec.SectionName = "FORMAT" Then
            ' [USER-REQ 2026-06-09] Prefer FormatName, fallback to Description.
            If sec.HasKey("FormatName") Then
                ExtractDescription = sec.GetValue("FormatName")
            ElseIf sec.HasKey("Description") Then
                ExtractDescription = sec.GetValue("Description")
            End If
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1251] modFormatLoader.ExtractDescription EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next i
End Function

' FormatExists: check if .txt exists
Public Function FormatExists(ByVal formatId As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1252] modFormatLoader.FormatExists ENTER"  ' [ADR-0100]
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    FormatExists = fso.FileExists(modConfigHolder.GetFormatDir() & formatId & ".txt")
End Function

' SaveFormat: 管理.xlsm のみ enforce (Q19) のうえ modStanzaIO.WriteStanzaFile で書き出し
Public Function SaveFormat(ByVal formatId As String, ByVal sections As Collection) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1253] modFormatLoader.SaveFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' Q19 enforce: xlsmName check
    Dim wbName As String
    wbName = modConfigHolder.GetValueOrDefault("xlsmName", "")
    If LCase(wbName) <> "kanri" Then
        Dim actualName As String
        actualName = ThisWorkbook.Name
        If actualName <> KANRI_XLSM_NAME And LCase(wbName) <> "kanri" Then
            SaveFormat = 2 ' rejected by Q19
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1254] modFormatLoader.SaveFormat EXIT-OK"  ' [ADR-0100]
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
    ' [Change-5] ensure the per-format data folder exists alongside the format
    If Not fso.FolderExists(modConfigHolder.GetDataDir()) Then fso.CreateFolder modConfigHolder.GetDataDir()
    Dim ddir As String
    ddir = modConfigHolder.GetDataDir() & formatId & "\"
    If Not fso.FolderExists(ddir) Then fso.CreateFolder ddir
    SaveFormat = 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1255] modFormatLoader.SaveFormat EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1256] modFormatLoader.SaveFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    SaveFormat = 3
End Function

' DeleteFormat: enforce + delete file
' Q55: reject (ret=2) if any knowledge of this format still exists.
' Return codes: 0=success / 1=error / 2=rejected by Q55
Public Function DeleteFormat(ByVal formatId As String) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1257] modFormatLoader.DeleteFormat ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    ' Q55 check: any knowledge using this format?
    Dim usingKnw As Collection
    Set usingKnw = modKnowledgeFileIO.ListKnowledgesByFormat(formatId)
    If Not usingKnw Is Nothing Then
        If usingKnw.Count > 0 Then
            DeleteFormat = 2  ' Q55 reject
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1258] modFormatLoader.DeleteFormat EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    End If
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath As String
    filePath = modConfigHolder.GetFormatDir() & formatId & ".txt"
    If fso.FileExists(filePath) Then fso.DeleteFile filePath
    ' [Change-5] remove the now-orphan per-format data folder when empty
    ' (Q55 above already guarantees no knowledge of this format remains).
    Dim ddir As String
    ddir = modConfigHolder.GetDataDir() & formatId & "\"
    If fso.FolderExists(ddir) Then
        Dim df As Object
        Set df = fso.GetFolder(ddir)
        If df.Files.Count = 0 And df.SubFolders.Count = 0 Then df.Delete True
    End If
    DeleteFormat = 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1259] modFormatLoader.DeleteFormat EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1260] modFormatLoader.DeleteFormat EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    DeleteFormat = 1
End Function

' Private LogError stub
Private Sub LogError(ByVal moduleName As String, ByVal funcName As String, ByVal msg As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1261] modFormatLoader.LogError ENTER"  ' [ADR-0100]
    Debug.Print "[ERR] " & moduleName & "." & funcName & ": " & msg
End Sub

' [USER-REQ 2026-06-09] Returns True when the [FORMAT].Status value matches
' the disabled marker (CP932 "????" = 無効, or the JP literal directly).
Private Function IsFormatDisabled(ByVal sections As Collection) As Boolean
    On Error Resume Next
    If sections Is Nothing Then Exit Function
    Dim sec As ClsStanzaSection
    Dim i As Long
    Dim s As String
    Dim disabled As String
    disabled = ChrW(&H7121) & ChrW(&H52B9)
    For i = 1 To sections.Count
        Set sec = sections.Item(i)
        If sec.SectionName = "FORMAT" Then
            If sec.HasKey("Status") Then
                s = sec.GetValue("Status")
                If s = disabled Or s = "????" Then
                    IsFormatDisabled = True
                    Exit Function
                End If
            End If
            Exit For
        End If
    Next i
End Function
```
