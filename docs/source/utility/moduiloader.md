---
title: modUILoader.bas
---

# modUILoader.bas

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 3 ブック共通 |
| 役割 | UI 定義 .txt を読み込み、シートにレイアウトを適用する |
| 行数 | 322 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modUILoader.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modUILoader"
' ================================================================
' ���W���[��: modUILoader�iPhase 3 / ���[�e�B���e�B�w�j
' �T�v:       UI ��` .txt �Ǎ� + ApplyUiToSheet�iv2.1 �Ŋ֐�����{���j
'             ADR-0053 ��2.9 �O�� I/O �Ɨ����K�񏅎�
'             v2_ui_stanza_schema.md ��4 �K�p��������
' Version:    v2.1�i2026-05-16 EOD�AQ1-Q57 �e���f���f�j
' �ˑ���:     modStanzaIO�iparser�j, modConfigHolder�iui_dir�j, ClsStanzaSection
' �֘A:       ADR-0053 ��2.3 / ��2.9�AADR-0056�imockup �� UI �X�^���U���o�j
'             Q5�iUI ��` .txt �� read-only�ASave / Write API �Ȃ��j
'             Q20�i���J I/F 3 �֐��FLoadUiDefinition / LoadUiList / ApplyUiToSheet�j
'             Q24�i`#` �R�����g�s�e�F�A�s���̂݁j
'             Q29�i[NOTE] �Z�N�V�����FPosition / Text / FontSize / Italic�j
' ================================================================
Option Explicit

' ----------------------------------------------------------------
' �萔�i�K�p���Av2_ui_stanza_schema.md ��4 �����j
' ----------------------------------------------------------------
' [SHEET] �� [COLUMN] �� [ROW] �� [FONT] �� [HEADER (merge ��s)] ��
' [LABEL] �� [INPUT] �� [NOTE] �� [GRID] �� [BUTTON] �� [BUTTON_TEMPLATE]

' ----------------------------------------------------------------
' Public I/F
' ----------------------------------------------------------------

' ================================================================
' �֐���: LoadUiDefinition�iQ20�j
' �T�v:   xlsm �� + screen ID ���� UI ��` .txt ��ǂݍ���� Collection ��Ԃ�
' ����:   ByVal xlsmName As String  - "�o�^�E�C��" / "�ݒ�" / "�Ǘ�"
'         ByVal screenId As String  - "M-02" ���iv2.1 �m�� 8 ��ʂ̂����ꂩ�j
' �߂�l: Collection - �v�f�� ClsStanzaSection
' ================================================================
Public Function LoadUiDefinition(ByVal xlsmName As String, ByVal screenId As String) As Collection
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = modConfigHolder.GetUiDir() & xlsmName & "\" & screenId & ".txt"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Set LoadUiDefinition = New Collection
        Exit Function
    End If

    Set LoadUiDefinition = modStanzaIO.ParseStanzaFile(filePath)
    Exit Function

ErrHandler:
    Set LoadUiDefinition = New Collection
End Function

' ================================================================
' �֐���: LoadUiList�iQ20�j
' �T�v:   ui_dir/<xlsmName>/ �z���� screenId �ꗗ��Ԃ��i�g���q�����j
' ================================================================
Public Function LoadUiList(ByVal xlsmName As String) As Collection
    On Error GoTo ErrHandler

    Dim result As Collection
    Set result = New Collection

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folderPath As String
    folderPath = modConfigHolder.GetUiDir() & xlsmName & "\"

    If Not fso.FolderExists(folderPath) Then
        Set LoadUiList = result
        Exit Function
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    Dim file As Object
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
            result.Add fso.GetBaseName(file.Name)
        End If
    Next file

    Set LoadUiList = result
    Exit Function

ErrHandler:
    Set LoadUiList = New Collection
End Function

' ================================================================
' �֐���: ApplyUiToSheet�iQ20�Av2.1 �Ŋ֐�����{���j
' �T�v:   xlsm �� + screen ID ���� UI ��` .txt ��Ǎ��݁A�ΏۃV�[�g�ɕ��� UI ��K�p
' ����:   ByVal xlsmName As String  - "�o�^�E�C��" / "�ݒ�" / "�Ǘ�"
'         ByVal screenId As String  - "M-02" ���iv2.1 �m�� 8 ��ʁj
'         ByVal ws As Worksheet     - �K�p��V�[�g
' �߂�l: Boolean - ���� TRUE / ���s FALSE�iuiSchemaFailMode �ɏ]���j
' ================================================================
Public Function ApplyUiToSheet( _
    ByVal xlsmName As String, _
    ByVal screenId As String, _
    ByVal ws As Worksheet _
) As Boolean
    On Error GoTo ErrHandler

    ' uiSchemaFailMode (v2_ui_stanza_schema.md 5): safeDefault vs abort.
    ' canonical config uses warn/strict tokens; abort==strict, safeDefault==warn.
    Dim failMode As String
    failMode = GetUiSchemaFailMode()

    Dim filePath As String
    filePath = modConfigHolder.GetUiDir() & xlsmName & "\" & screenId & ".txt"

    ' �t�@�C�����݃`�F�b�N
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Debug.Print "[UILoader] file not found: " & filePath
        If failMode = "abort" Then
            Err.Raise vbObjectError + 9301, "modUILoader.ApplyUiToSheet", _
                "UI definition file not found (abort mode): " & filePath
        End If
        ApplyUiToSheet = False
        Exit Function
    End If

    ' parse
    Dim sections As Collection
    Set sections = modStanzaIO.ParseStanzaFile(filePath)
    If sections.Count = 0 Then
        ' grammar error / empty parse result (schema 5 error condition)
        Debug.Print "[UILoader] parse failed: " & filePath
        If failMode = "abort" Then
            Err.Raise vbObjectError + 9302, "modUILoader.ApplyUiToSheet", _
                "UI definition grammar error (abort mode): " & filePath
        End If
        ApplyUiToSheet = False
        Exit Function
    End If

    ' �K�p�����Ƃɏ����imerge �Փˉ���A��4 �����j
    ApplySectionsByName ws, sections, "SHEET"
    ApplySectionsByName ws, sections, "COLUMN"
    ApplySectionsByName ws, sections, "ROW"
    ApplySectionsByName ws, sections, "FONT"
    ApplySectionsByName ws, sections, "HEADER"
    ApplySectionsByName ws, sections, "LABEL"
    ApplySectionsByName ws, sections, "INPUT"
    ApplySectionsByName ws, sections, "NOTE"
    ApplySectionsByName ws, sections, "GRID"
    ApplySectionsByName ws, sections, "BUTTON"
    ApplySectionsByName ws, sections, "BUTTON_TEMPLATE"

    ApplyUiToSheet = True
    Exit Function

ErrHandler:
    Debug.Print "[UILoader] " & screenId & " - " & Err.Description
    ' abort mode: re-raise so the caller can halt all-screen startup.
    ' safeDefault mode: swallow, return False (this screen only fails).
    If GetUiSchemaFailMode() = "abort" Then
        Err.Raise Err.Number, "modUILoader.ApplyUiToSheet", Err.Description
    End If
    ApplyUiToSheet = False
End Function

' ================================================================
' �֐���: GetUiSchemaFailMode
' �T�v:   config �� uiSchemaFailMode ��Ǎ��݁A"abort" / "safeDefault" ��Ԃ�
'         schema 5 �F safeDefault�i���j / abort�B
'         canonical config �g�[�N�� warn/strict �Ƃ̑Ή��F
'         strict==abort, warn==safeDefault�ischema �K�b�`�ɂ��ǌ��F�j
' �߂�l: String - "abort" �܂��� "safeDefault"
' ================================================================
Private Function GetUiSchemaFailMode() As String
    Dim s As String
    s = LCase(modConfigHolder.GetValueOrDefault("uiSchemaFailMode", "safeDefault"))
    If s = "abort" Or s = "strict" Then
        GetUiSchemaFailMode = "abort"
    Else
        GetUiSchemaFailMode = "safeDefault"
    End If
End Function

' ----------------------------------------------------------------
' Private: �Z�N�V������ʂ��Ƃ̓K�p�f�B�X�p�b�`
' ----------------------------------------------------------------

Private Sub ApplySectionsByName( _
    ByVal ws As Worksheet, _
    ByVal sections As Collection, _
    ByVal targetName As String _
)
    Dim sec As ClsStanzaSection
    For Each sec In sections
        If sec.SectionName = targetName Then
            Select Case targetName
                Case "SHEET":  ApplySheet ws, sec
                Case "COLUMN": ApplyColumn ws, sec
                Case "ROW":    ApplyRow ws, sec
                Case "FONT":   ApplyFont ws, sec
                Case "HEADER": ApplyHeader ws, sec
                Case "LABEL":  ApplyLabel ws, sec
                Case "INPUT":  ApplyInput ws, sec
                Case "NOTE":   ApplyNote ws, sec
                Case "GRID":   ApplyGrid ws, sec
                Case "BUTTON": ApplyButton ws, sec
                Case "BUTTON_TEMPLATE": ApplyButtonTemplate ws, sec
            End Select
        End If
    Next sec
End Sub

' ----------------------------------------------------------------
' Private: �ʃZ�N�V�����K�p�i���`�APhase 3 �m�莞�ɏڍ׎����j
' ----------------------------------------------------------------

Private Sub ApplySheet(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Dim tabColor As String
    tabColor = sec.GetValue("TabColor")
    If Len(tabColor) > 0 Then
        ws.Tab.Color = HexToRgbLong(tabColor)
    End If
    ' SheetName �͓K�p�ϑO��imockup_reference �\�z���ɃV�[�g����v�j
End Sub

Private Sub ApplyColumn(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Dim key As Variant
    Dim colLetter As String
    Dim widthVal As Double
    For Each key In sec.KeyValues.Keys
        If Left(CStr(key), 12) = "ColumnWidth_" Then
            colLetter = Mid(CStr(key), 13)
            widthVal = CDbl(sec.GetValue(CStr(key)))
            ws.Columns(colLetter).ColumnWidth = widthVal
        End If
    Next key
End Sub

Private Sub ApplyRow(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Dim key As Variant
    Dim rowNum As Long
    Dim heightVal As Double
    For Each key In sec.KeyValues.Keys
        If Left(CStr(key), 10) = "RowHeight_" Then
            rowNum = CLng(Mid(CStr(key), 11))
            heightVal = CDbl(sec.GetValue(CStr(key)))
            ws.Rows(rowNum).RowHeight = heightVal
        End If
    Next key
End Sub

Private Sub ApplyFont(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Dim defaultFont As String
    Dim defaultSize As String
    defaultFont = sec.GetValue("DefaultFont")
    defaultSize = sec.GetValue("DefaultFontSize")
    If Len(defaultFont) > 0 Then
        ws.Cells.Font.Name = defaultFont
    End If
    If Len(defaultSize) > 0 Then
        ws.Cells.Font.Size = CDbl(defaultSize)
    End If
End Sub

Private Sub ApplyHeader(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Dim cellAddr As String
    cellAddr = sec.GetValue("Cell")
    If Len(cellAddr) = 0 Then Exit Sub

    Dim targetRange As Range
    Set targetRange = ws.Range(cellAddr)

    ' merge�i�͈͎w��̏ꍇ�j
    If InStr(cellAddr, ":") > 0 Then
        On Error Resume Next
        targetRange.Merge
        On Error GoTo 0
    End If

    ' ���e�ݒ�
    targetRange.Cells(1, 1).Value = UnescapeStanzaValue(sec.GetValue("Text"))
    If sec.HasKey("BackColor") Then
        targetRange.Interior.Color = HexToRgbLong(sec.GetValue("BackColor"))
    End If
    If sec.HasKey("ForeColor") Then
        targetRange.Font.Color = HexToRgbLong(sec.GetValue("ForeColor"))
    End If
    If sec.HasKey("FontSize") Then
        targetRange.Font.Size = CDbl(sec.GetValue("FontSize"))
    End If
    If sec.GetValue("Bold") = "TRUE" Then
        targetRange.Font.Bold = True
    End If
End Sub

Private Sub ApplyLabel(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Dim cellAddr As String
    cellAddr = sec.GetValue("Cell")
    If Len(cellAddr) = 0 Then Exit Sub
    Dim r As Range
    Set r = ws.Range(cellAddr)
    r.Value = modStanzaIO.UnescapeStanzaValue(sec.GetValue("Text"))
End Sub

' Q5=B full body: input cell color / border / data validation per schema 3.6.
Private Sub ApplyInput(ByVal ws As Worksheet, ByVal sec As ClsStanzaSection)
    Dim cellAddr As String
    cellAddr = sec.GetValue("Cell")
    If Len(cellAddr) = 0 Then Exit Sub
    Dim r As Range
    Set r = ws.Range(cellAddr)

    ' input background color: explicit InputColor, else schema 3.6 default #FFFFCC
    If sec.HasKey("InputColor") Then
        r.Interior.Color = HexToRgbLong(sec.GetValue("InputColor"))
    Else
        r.Interior.Color = HexToRgbLong("#FFFFCC")
    End If

    ' border (schema 3.6: none / thin / medium / thick, default thin)
    Dim borderVal As String
    If sec.H
```
