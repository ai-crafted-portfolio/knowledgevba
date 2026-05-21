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
ErrHandler:
    Debug.Print "[ERR] Btn_SaveFormat_v21: " & Err.Number & " " & Err.Description
End Sub

Public Sub Btn_LoadFormat_v21()
    On Error GoTo ErrHandler
    Dim ui As Collection: Set ui = modUILoader.LoadUiDefinition("Kanri", "M-03")
    If ui.Count = 0 Then Exit Sub
    Dim tempDict As Object: Set tempDict = BuildFormatDictFromCells(ActiveSheet, ui)
    Dim fmtId As String
    If tempDict.Exists("FormatID") Then fmtId = CStr(tempDict("FormatID"))
    If Len(fmtId) = 0 Then Exit Sub
    LoadFormat_Workflow ActiveSheet, ui, fmtId
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Btn_LoadFormat_v21: " & Err.Number & " " & Err.Description
End Sub

Public Sub Btn_BackToMain_v21()
    On Error GoTo ErrHandler
    Dim hasInput As Boolean
    hasInput = HasUnsavedInput_M03(ActiveSheet)
    If hasInput Then
        On Error Resume Next
        Dim oLogger031v As clsLogger
        Set oLogger031v = New clsLogger
        oLogger031v.Init ThisWorkbook.Worksheets("LOG")
        oLogger031v.LogWarn "modEntryFormat", "Btn_BackToMain_v21", "M-03 unsaved input detected, confirm dialog shown", "", "VALIDATE-WARN-WW-031"
        On Error GoTo ErrHandler
        ' Headless guard: under COM automation skip the modal dialog and
        ' fall through to the default (discard and navigate to main).
        Dim r As VbMsgBoxResult
        If IsHeadless() Then
            r = vbNo
        Else
            r = MsgBox("Unsaved input. Save before back to main?" & vbCrLf & _
                       "Yes=Save and back / No=Discard and back / Cancel=Stay editing", _
                       vbYesNoCancel + vbExclamation, "Confirm")
        End If
        Select Case r
            Case vbYes
                Btn_SaveFormat_v21
                NavigateToMain_v21
            Case vbNo
                NavigateToMain_v21
            Case vbCancel
                Exit Sub
        End Select
    Else
        NavigateToMain_v21
    End If
    Exit Sub
ErrHandler:
    On Error Resume Next
    Dim oLogger032b As clsLogger
    Set oLogger032b = New clsLogger
    oLogger032b.Init ThisWorkbook.Worksheets("LOG")
    oLogger032b.LogError "modEntryFormat", "Btn_BackToMain_v21", "Back to main transition error: " & Err.Description, "", "BACKTOMAIN-ERR-EE-032"
    On Error GoTo 0
    If Not IsHeadless() Then
        MsgBox "Btn_BackToMain_v21 error: " & Err.Description, vbCritical, "Error"
    Else
        Debug.Print "[ERR] Btn_BackToMain_v21: " & Err.Number & " " & Err.Description
    End If
End Sub

Private Function HasUnsavedInput_M03(ByVal ws As Object) As Boolean
    On Error GoTo ErrHandler
    Dim v1 As String, v2 As String, v3 As String
    v1 = CStr(ws.Range("C3").Value)
    v2 = CStr(ws.Range("C4").Value)
    v3 = CStr(ws.Range("C5").Value)
    If Len(v1) > 0 Or Len(v2) > 0 Or Len(v3) > 0 Then
        HasUnsavedInput_M03 = True
        Exit Function
    End If
    If Len(CStr(ws.Range("B7").Value)) > 0 Then
        HasUnsavedInput_M03 = True
        Exit Function
    End If
    HasUnsavedInput_M03 = False
    Exit Function
ErrHandler:
    HasUnsavedInput_M03 = False
End Function

Private Sub NavigateToMain_v21()
    On Error Resume Next
    ThisWorkbook.Worksheets("Main").Activate
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Worksheets("M-01").Activate
    End If
    On Error GoTo 0
End Sub

Public Sub Btn_ConfirmDiff_v21()
    On Error GoTo ErrHandler
    Dim ws As Object: Set ws = ActiveSheet
    Dim ui As Collection: Set ui = modUILoader.LoadUiDefinition("Kanri", "M-03")
    If ui.Count = 0 Then Exit Sub

    Dim tempDict As Object: Set tempDict = BuildFormatDictFromCells(ws, ui)
    Dim fmtId As String
    If tempDict.Exists("FormatID") Then fmtId = CStr(tempDict("FormatID"))
    If Len(fmtId) = 0 Then Exit Sub

    Dim newVer As Long
    newVer = 2
    If tempDict.Exists("FormatVersion") Then
        Dim verStr As String
        verStr = CStr(tempDict("FormatVersion"))
        If Len(verStr) > 0 And IsNumeric(verStr) Then newVer = CLng(verStr)
    End If

    Dim affected As Collection
    Set affected = modKnowledgeFileIO.ListKnowledgesByFormat(fmtId)
    Dim affCount As Long
    If Not affected Is Nothing Then affCount = affected.Count

    Dim msg As String
    msg = "FormatID=" & fmtId & vbCrLf
    msg = msg & "NewVersion=" & newVer & vbCrLf
    msg = msg & "Affected knowledge count: " & affCount & vbCrLf
    msg = msg & "Migrate now?" & vbCrLf
    msg = msg & "(Yes=migrate / No=cancel)"

    ' Headless guard: when driven by COM automation, skip the modal confirm
    ' dialog and proceed with the default (migrate). An unguarded MsgBox blocks
    ' indefinitely under headless run.
    ' NOTE (2026-05-20 hang fix): Application.Interactive is TRUE by default in
    ' a COM-automated Excel even when Visible=False, so the old single-property
    ' guard let the MsgBox through and blocked forever. IsHeadless() now checks
    ' multiple signals (Interactive AND Visible).
    Dim r As VbMsgBoxResult
    If IsHeadless() Then
        r = vbYes
    Else
        r = MsgBox(msg, vbYesNo + vbQuestion, "Confirm format diff")
    End If

    If r = vbYes Then
        Dim mig As Object: Set mig = New clsFieldMigrator
        Dim ret As Long
        ret = mig.MigrateAllByFormat(fmtId, newVer)
        If Not IsHeadless() Then
            MsgBox "Migration done. ret=" & ret, vbInformation, "Result"
        End If
    End If
    Exit Sub
ErrHandler:
    If Not IsHeadless() Then
        MsgBox "Btn_ConfirmDiff_v21 error: " & Err.Description, vbCritical, "Error"
    Else
        Debug.Print "[ERR] Btn_ConfirmDiff_v21: " & Err.Number & " " & Err.Description
    End If
End Sub

' ------------------------------------------------------------
' IsHeadless: TRUE when running under COM automation with no
' interactive UI. Used to suppress modal dialogs (MsgBox / InputBox)
' that would otherwise block a headless harness indefinitely.
' Checks two independent signals because Application.Interactive
' alone defaults to TRUE under 
```
