---
title: modEntryFormat.bas
---

# modEntryFormat.bas

| 項目 | 内容 |
|---|---|
| 層 | エントリポイント層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 管理.xlsm |
| 役割 | フォーマット設計画面のボタン処理。フォーマット定義セルとデータの相互変換、保存・読込ワークフロー |
| 行数 | 346 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modEntryFormat.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modEntryFormat"
' ============================================================
' modEntryFormat (Phase 3 / slim, SRP split, ASCII-only)
' Bridge + Workflow + Btn_, helpers extracted to clsCellAddrHelper / clsCellIO / clsGridIO
' ============================================================
Option Explicit

Private Const FIELDS_KEY As String = "Fields"

Public Function BuildFormatDictFromCells(ByVal target As Object, ByVal uiSections As Collection) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Set result(FIELDS_KEY) = New Collection
    If uiSections Is Nothing Then
        Set BuildFormatDictFromCells = result
        Exit Function
    End If
    Dim mode As String
    mode = LCase(modConfigHolder.GetValueOrDefault("uiSchemaFailMode", "warn"))
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To uiSections.Count
        Set sec = uiSections.Item(i)
        Select Case sec.SectionName
            Case "INPUT"
                Dim cellAddr As String, fieldName As String
                cellAddr = sec.GetValue("Cell")
                fieldName = sec.GetValue("Name")
                If Len(fieldName) > 0 And Len(cellAddr) > 0 Then
                    result(fieldName) = clsCellIO.ReadCellValue(target, cellAddr)
                ElseIf mode = "strict" Then
                    Err.Raise vbObjectError + 9201, "modEntryFormat.BuildFormatDictFromCells", _
                              "INPUT stanza missing Name or Cell key"
                End If
            Case "GRID"
                Set result(FIELDS_KEY) = clsGridIO.ReadGridFields(target, sec, mode)
        End Select
    Next i
    Set BuildFormatDictFromCells = result
End Function

Public Sub WriteFormatDictToCells(ByVal target As Object, ByVal uiSections As Collection, ByVal formatDict As Object)
    If uiSections Is Nothing Or formatDict Is Nothing Then Exit Sub
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To uiSections.Count
        Set sec = uiSections.Item(i)
        Select Case sec.SectionName
            Case "INPUT"
                Dim cellAddr As String, fieldName As String
                cellAddr = sec.GetValue("Cell")
                fieldName = sec.GetValue("Name")
                If Len(fieldName) > 0 And Len(cellAddr) > 0 Then
                    If formatDict.Exists(fieldName) Then
                        Dim v As Variant
                        v = formatDict(fieldName)
                        If IsNull(v) Then
                            clsCellIO.WriteCellValue target, cellAddr, ""
                        Else
                            clsCellIO.WriteCellValue target, cellAddr, CStr(v)
                        End If
                    End If
                End If
            Case "GRID"
                If formatDict.Exists(FIELDS_KEY) Then
                    clsGridIO.WriteGridFields target, sec, formatDict(FIELDS_KEY)
                End If
        End Select
    Next i
End Sub

Public Function SaveFormat_Workflow(ByVal target As Object, ByVal uiSections As Collection) As String
    On Error GoTo ErrHandler
    Dim fmtDict As Object
    Set fmtDict = BuildFormatDictFromCells(target, uiSections)
    Dim fmtId As String
    If fmtDict.Exists("FormatID") Then fmtId = CStr(fmtDict("FormatID"))
    If Len(fmtId) = 0 Then
        SaveFormat_Workflow = ""
        Exit Function
    End If
    Dim sections As Collection
    Set sections = New Collection
    Dim fmtSec As ClsStanzaSection
    Set fmtSec = New ClsStanzaSection
    fmtSec.Init "FORMAT", 1
    fmtSec.SetValue "FormatID", fmtId
    If fmtDict.Exists("FormatName") Then fmtSec.SetValue "FormatName", CStr(fmtDict("FormatName"))
    If fmtDict.Exists("FormatVersion") Then fmtSec.SetValue "FormatVersion", CStr(fmtDict("FormatVersion"))
    sections.Add fmtSec
    If fmtDict.Exists(FIELDS_KEY) Then
        Dim fields As Collection
        Set fields = fmtDict(FIELDS_KEY)
        Dim i As Long
        For i = 1 To fields.Count
            Dim row As Object
            Set row = fields.Item(i)
            Dim fldSec As ClsStanzaSection
            Set fldSec = New ClsStanzaSection
            fldSec.Init "FIELD", 1
            If row.Exists("Name") Then fldSec.SetValue "FieldName", CStr(row("Name"))
            If row.Exists("Type") Then fldSec.SetValue "FieldType", CStr(row("Type"))
            If row.Exists("Required") Then fldSec.SetValue "Required", CStr(row("Required"))
            sections.Add fldSec
        Next i
    End If
    Dim ret As Long
    ret = modFormatLoader.SaveFormat(fmtId, sections)
    If ret = 0 Then
        SaveFormat_Workflow = fmtId
        On Error Resume Next
        Dim oLogger006 As clsLogger
        Set oLogger006 = New clsLogger
        oLogger006.Init ThisWorkbook.Worksheets("LOG")
        oLogger006.LogInfo "modEntryFormat", "SaveFormat_Workflow", "Format save complete: " & fmtId, fmtId, "SAVE-EXIT-OK-II-006"
        On Error GoTo 0
    Else
        SaveFormat_Workflow = ""
    End If
    Exit Function
ErrHandler:
    On Error Resume Next
    Dim oLogger032 As clsLogger
    Set oLogger032 = New clsLogger
    oLogger032.Init ThisWorkbook.Worksheets("LOG")
    oLogger032.LogWarn "modEntryFormat", "SaveFormat_Workflow", "Format save failed: " & Err.Description, fmtId, "VALIDATE-WARN-WW-032"
    On Error GoTo 0
    SaveFormat_Workflow = ""
End Function

Public Function LoadFormat_Workflow(ByVal target As Object, ByVal uiSections As Collection, ByVal formatId As String) As Boolean
    On Error GoTo ErrHandler
    If Len(formatId) = 0 Then
        LoadFormat_Workflow = False
        Exit Function
    End If
    Dim loaded As Collection
    Set loaded = modFormatLoader.LoadFormat(formatId)
    If loaded.Count = 0 Then
        LoadFormat_Workflow = False
        Exit Function
    End If
    Dim formatDict As Object
    Set formatDict = CreateObject("Scripting.Dictionary")
    Set formatDict(FIELDS_KEY) = New Collection
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To loaded.Count
        Set sec = loaded.Item(i)
        Select Case sec.SectionName
            Case "FORMAT"
                If sec.HasKey("FormatID") Then formatDict("FormatID") = sec.GetValue("FormatID")
                If sec.HasKey("FormatName") Then formatDict("FormatName") = sec.GetValue("FormatName")
                If sec.HasKey("FormatVersion") Then formatDict("FormatVersion") = sec.GetValue("FormatVersion")
            Case "FIELD"
                Dim row As Object
                Set row = CreateObject("Scripting.Dictionary")
                If sec.HasKey("FieldName") Then row("Name") = sec.GetValue("FieldName")
                If sec.HasKey("FieldType") Then row("Type") = sec.GetValue("FieldType")
                If sec.HasKey("Required") Then row("Required") = sec.GetValue("Required")
                formatDict(FIELDS_KEY).Add row
        End Select
    Next i
    WriteFormatDictToCells target, uiSections, formatDict
    LoadFormat_Workflow = True
    Exit Function
ErrHandler:
    LoadFormat_Workflow = False
End Function

Public Sub Btn_SaveFormat_v21()
    On Error GoTo ErrHandler
    Dim ui As Collection: Set ui = modUILoader.LoadUiDefinition("Kanri", "M-03")
    If ui.Count = 0 Then Exit Sub
    SaveFormat_Workflow ActiveSheet, ui
    On Error Resume Next
    Dim oLogger007 As clsLogger
    Set oLogger007 = New clsLogger
    oLogger007.Init ThisWorkbook.Worksheets("LOG")
    oLogger007.LogInfo "modEntryFormat", "Btn_SaveFormat_v21", "Btn_SaveFormat_v21 done", "", "SAVE-EXIT-OK-II-007"
    On Error GoTo 0
    Exit Sub
ErrHandle