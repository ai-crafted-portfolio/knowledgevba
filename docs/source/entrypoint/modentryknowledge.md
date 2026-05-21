---
title: modEntryKnowledge.bas
---

# modEntryKnowledge.bas

| 項目 | 内容 |
|---|---|
| 層 | エントリポイント層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 登録修正.xlsm |
| 役割 | ナレッジ登録・修正画面のボタン処理。入力セルとナレッジデータの相互変換、保存・読込ワークフロー |
| 行数 | 208 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modEntryKnowledge.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modEntryKnowledge"
' ============================================================
' modEntryKnowledge (Phase 3 / slim, Sprint2 SRP split, ASCII-only)
' Bridge + Workflow + Btn_, helpers extracted to clsCellIO
' ============================================================
Option Explicit

Public Function BuildDictFromCells(ByVal target As Object, ByVal uiSections As Collection) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    If uiSections Is Nothing Then
        Set BuildDictFromCells = result
        Exit Function
    End If
    Dim mode As String
    mode = GetUiFailMode()
    Dim inputCount As Long
    inputCount = 0
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To uiSections.Count
        Set sec = uiSections.Item(i)
        If sec.SectionName = "INPUT" Then
            inputCount = inputCount + 1
            Dim cellAddr As String, fieldName As String
            cellAddr = sec.GetValue("Cell")
            fieldName = sec.GetValue("Name")
            If Len(fieldName) = 0 Then
                If mode = "strict" Then
                    Err.Raise vbObjectError + 9001, "modEntryKnowledge.BuildDictFromCells", _
                              "INPUT stanza missing Name key"
                End If
            ElseIf Len(cellAddr) = 0 Then
                If mode = "strict" Then
                    Err.Raise vbObjectError + 9002, "modEntryKnowledge.BuildDictFromCells", _
                              "INPUT stanza Cell key empty for Name=" & fieldName
                End If
            Else
                On Error GoTo CellErrHandler
                result(fieldName) = clsCellIO.ReadCellValue(target, cellAddr)
                On Error GoTo 0
            End If
        End If
NextStanza:
    Next i
    If inputCount = 0 Then
        If mode = "strict" Then
            Err.Raise vbObjectError + 9010, "modEntryKnowledge.BuildDictFromCells", _
                      "No INPUT stanza found"
        End If
    End If
    Set BuildDictFromCells = result
    Exit Function
CellErrHandler:
    If mode = "strict" Then
        Err.Raise vbObjectError + 9003, "modEntryKnowledge.BuildDictFromCells", _
                  "Invalid cell address: " & cellAddr
    Else
        Resume NextStanza
    End If
End Function

Public Sub WriteDictToCells(ByVal target As Object, ByVal uiSections As Collection, ByVal dict As Object)
    If uiSections Is Nothing Or dict Is Nothing Then Exit Sub
    Dim nameCellMap As Object
    Set nameCellMap = CreateObject("Scripting.Dictionary")
    Dim sec As ClsStanzaSection
    Dim i As Long
    For i = 1 To uiSections.Count
        Set sec = uiSections.Item(i)
        If sec.SectionName = "INPUT" Then
            Dim nm As String, ca As String
            nm = sec.GetValue("Name")
            ca = sec.GetValue("Cell")
            If Len(nm) > 0 And Len(ca) > 0 Then nameCellMap(nm) = ca
        End If
    Next i
    Dim keys As Variant
    keys = dict.Keys
    Dim k As Long
    For k = 0 To dict.Count - 1
        Dim keyName As String
        keyName = CStr(keys(k))
        If nameCellMap.Exists(keyName) Then
            Dim cellAddr As String
            cellAddr = CStr(nameCellMap(keyName))
            If IsObject(dict(keyName)) Then
                Err.Raise vbObjectError + 9021, "modEntryKnowledge.WriteDictToCells", _
                          "Cannot write object to single cell, key=" & keyName
            End If
            Dim v As Variant
            v = dict(keyName)
            If IsArray(v) Then
                Err.Raise vbObjectError + 9020, "modEntryKnowledge.WriteDictToCells", _
                          "Cannot write array to single cell, key=" & keyName
            End If
            Dim valueStr As String
            If IsNull(v) Then
                valueStr = ""
            ElseIf VarType(v) = vbDate Then
                valueStr = Format(v, "yyyy-mm-dd")
            Else
                valueStr = CStr(v)
            End If
            clsCellIO.WriteCellValue target, cellAddr, valueStr
        End If
    Next k
End Sub

Public Function SaveKnowledge_Workflow(ByVal target As Object, ByVal uiSections As Collection, ByVal knowledgeNo As String) As String
    On Error GoTo ErrHandler
    Dim dict As Object
    Set dict = BuildDictFromCells(target, uiSections)
    If dict.Count = 0 Then
        SaveKnowledge_Workflow = ""
        Exit Function
    End If
    Dim knwNo As String
    If Len(knowledgeNo) > 0 Then
        knwNo = knowledgeNo
    ElseIf dict.Exists("KnowledgeNo") Then
        knwNo = CStr(dict("KnowledgeNo"))
    Else
        knwNo = ""
    End If
    If Len(knwNo) = 0 Then
        SaveKnowledge_Workflow = ""
        Exit Function
    End If
    dict("KnowledgeNo") = knwNo
    If Not dict.Exists("UpdatedAt") Then
        dict("UpdatedAt") = Format(Now(), "yyyy-mm-dd hh:nn:ss")
    End If
    Dim ret As Long
    ret = modKnowledgeFileIO.SaveKnowledge(knwNo, dict, 0)
    If ret = 0 Then
        SaveKnowledge_Workflow = knwNo
        ' S5-LOG-02: SAVE-EXIT-OK-II-004 (Knowledge save success, before exit)
        On Error Resume Next
        Dim oLogger004 As clsLogger
        Set oLogger004 = New clsLogger
        oLogger004.Init ThisWorkbook.Worksheets("LOG")
        oLogger004.LogInfo "modEntryKnowledge", "SaveKnowledge_Workflow", "Knowledge 保存完了: " & knwNo, knwNo, "SAVE-EXIT-OK-II-004"
        On Error GoTo 0
    Else
        SaveKnowledge_Workflow = ""
    End If
    Exit Function
ErrHandler:
    SaveKnowledge_Workflow = ""
End Function

Public Function LoadKnowledge_Workflow(ByVal target As Object, ByVal uiSections As Collection, ByVal knowledgeNo As String) As Boolean
    On Error GoTo ErrHandler
    If Len(knowledgeNo) = 0 Then
        LoadKnowledge_Workflow = False
        Exit Function
    End If
    Dim dict As Object
    Set dict = modKnowledgeFileIO.LoadKnowledge(knowledgeNo)
    If dict.Count = 0 Then
        LoadKnowledge_Workflow = False
        Exit Function
    End If
    WriteDictToCells target, uiSections, dict
    LoadKnowledge_Workflow = True
    Exit Function
ErrHandler:
    LoadKnowledge_Workflow = False
End Function

Public Sub Btn_SaveKnowledge_v21()
    On Error GoTo ErrHandler
    Dim ui As Collection: Set ui = modUILoader.LoadUiDefinition("Touroku", "M-05")
    If ui.Count = 0 Then Exit Sub
    Dim knwNo As String
    knwNo = SaveKnowledge_Workflow(ActiveSheet, ui, "")
    ' S5-LOG-02: SAVE-EXIT-OK-II-005 (Btn_SaveKnowledge_v21 button handler exit)
    On Error Resume Next
    Dim oLogger005 As clsLogger
    Set oLogger005 = New clsLogger
    oLogger005.Init ThisWorkbook.Worksheets("LOG")
    oLogger005.LogInfo "modEntryKnowledge", "Btn_SaveKnowledge_v21", "Btn_SaveKnowledge_v21 完了 knwNo=" & knwNo, knwNo, "SAVE-EXIT-OK-II-005"
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Btn_SaveKnowledge_v21: " & Err.Number & " " & Err.Description
End Sub

Public Sub Btn_LoadKnowledge_v21()
    On Error GoTo ErrHandler
    Dim ui As Collection: Set ui = modUILoader.LoadUiDefinition("Touroku", "M-09")
    If ui.Count = 0 Then Exit Sub
    Dim tempDict As Object: Set tempDict = BuildDictFromCells(ActiveSheet, ui)
    Dim knwNo As String
    If tempDict.Exists("KnowledgeNo") Then knwNo = CStr(tempDict("KnowledgeNo"))
    If Len(knwNo) = 0 Then Exit Sub
    LoadKnowledge_Workflow ActiveSheet, ui, knwNo
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] Btn_LoadKnowledge_v21: " & Err.Number & " " & Err.Description
End Sub

Private Function GetUiFailMode() As String
    Dim s As String
    s = LCase(modConfigHolder.GetValueOrDefault("uiSchemaFailMode", "warn"))
    If s = "strict" Then GetUiFailMode = "strict" Else GetUiFailMode = "warn"
End Function
```
