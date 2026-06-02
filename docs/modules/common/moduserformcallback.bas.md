---
title: modUserFormCallback.bas
description: modUserFormCallback.bas のソースコード（コピペ用）
---

# modUserFormCallback.bas

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modUserFormCallback.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modUserFormCallback"
' ================================================================
' modUserFormCallback (v2.3 Phase N-3, 2026-05-27)
' Purpose:
'   Bridge from dynamically-generated UserForm button click handlers
'   back to the calling clsUserFormRenderer instance. The injected
'   handlers in the dynamic form call Application.Run "modUserFormCallback.OnXxx"
'   which routes back to the singleton renderer instance.
' Design:
'   The renderer instance is stashed in modUserFormCallback at ShowForm
'   start, then handlers (OnRegister/OnUpdate/OnDelete/OnEdit/OnClear)
'   look it up to write the return value or trigger an action.
' ================================================================
Option Explicit

Private m_renderer As clsUserFormRenderer

Public Sub SetRenderer(ByVal r As clsUserFormRenderer)
    Set m_renderer = r
End Sub

' Called from dynamic form's btnRegister_Click
Public Sub OnRegister()
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Sub
    Dim newId As String
    newId = PersistFromActiveForm("register")
    m_renderer.SetReturnId newId
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] OnRegister: " & Err.Number & " " & Err.Description
End Sub

' Called from dynamic form's btnUpdate_Click
Public Sub OnUpdate()
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Sub
    Dim id As String
    id = PersistFromActiveForm("edit")
    m_renderer.SetReturnId id
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] OnUpdate: " & Err.Number & " " & Err.Description
End Sub

' Called from dynamic form's btnDelete_Click
Public Sub OnDelete()
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Sub
    Dim kid As String
    kid = m_renderer.GetKnowledgeId()
    If Len(kid) = 0 Then Exit Sub
    Dim ok As Boolean
    ok = modKnowledgeFileIO.DeleteKnowledge(kid)
    If ok Then m_renderer.SetReturnId "DELETED"
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] OnDelete: " & Err.Number & " " & Err.Description
End Sub

' Called from dynamic form's btnEdit_Click (M-09 view -> M-06 edit)
Public Sub OnEdit()
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Sub
    Dim kid As String
    kid = m_renderer.GetKnowledgeId()
    m_renderer.SetReturnId "EDIT:" & kid
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] OnEdit: " & Err.Number & " " & Err.Description
End Sub

' Called from dynamic form's btnClear_Click (M-05 register)
Public Sub OnClear()
    On Error GoTo ErrHandler
    Dim frm As Object
    Set frm = GetActiveDynForm()
    If frm Is Nothing Then Exit Sub
    Dim ctl As Object
    For Each ctl In frm.Controls
        Dim nm As String
        nm = ctl.Name
        ' Phase R-2: never clear the format selector or readonly header fields.
        If nm = "txtFormatId" Or nm = "cboFormatId" Or Left$(nm, 7) = "hdrctl_" Then
            ' skip
        ElseIf TypeName(ctl) = "TextBox" Then
            ctl.Text = ""
        ElseIf TypeName(ctl) = "ComboBox" Then
            ctl.Value = ""
        End If
    Next ctl
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] OnClear: " & Err.Number & " " & Err.Description
End Sub

' Phase R-2 F-3: called from dynamic form's cboFormatId_Change. Reads the newly
' selected format and, if it differs from the current one, unloads the form so
' ShowForm's loop rebuilds it with the new format's field set.
Public Sub OnFormatChange()
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Sub
    Dim frm As Object
    Set frm = GetActiveDynForm()
    If frm Is Nothing Then Exit Sub
    Dim newFmt As String
    On Error Resume Next
    newFmt = CStr(frm.Controls("cboFormatId").Value)
    On Error GoTo ErrHandler
    Dim before As String
    before = m_renderer.GetFormatId()
    m_renderer.RequestReformat newFmt
    If m_renderer.GetFormatId() <> before Then Unload frm
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] OnFormatChange: " & Err.Number & " " & Err.Description
End Sub

' R-3-a (2026-05-28): called from dynamic form's btnLoad_Click (M-06 edit).
' Reads the entered knowledge number and asks the renderer to reload + re-render.
Public Sub OnLoad()
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Sub
    Dim frm As Object
    Set frm = GetActiveDynForm()
    If frm Is Nothing Then Exit Sub
    Dim no As String
    On Error Resume Next
    no = CStr(FindCtl(frm, "txtKnowledgeNo").Text)   ' R-3-χ-4: frScroll 内
    On Error GoTo ErrHandler
    If Len(Trim$(no)) = 0 Then Exit Sub
    m_renderer.RequestLoad no
    Unload frm
    Exit Sub
ErrHandler:
    Debug.Print "[ERR] OnLoad: " & Err.Number & " " & Err.Description
End Sub

' Phase R-3-χ-4: field 系 controls は frScroll Frame 内にある。frm.Controls は直下のみ
' なので、frame 子も含めて探索/列挙するヘルパを使う(1 段ネスト = frScroll)。
Private Function FindCtl(ByVal frm As Object, ByVal nm As String) As Object
    On Error Resume Next
    Dim c As Object
    Set c = frm.Controls(nm)
    If Not c Is Nothing Then
        Set FindCtl = c
        Exit Function
    End If
    Dim f As Object
    For Each f In frm.Controls
        If TypeName(f) = "Frame" Then
            Dim cc As Object
            Set cc = Nothing
            Set cc = f.Controls(nm)
            If Not cc Is Nothing Then
                Set FindCtl = cc
                Exit Function
            End If
        End If
    Next f
End Function

Private Function AllFormControls(ByVal frm As Object) As Collection
    ' R-3-χ-4/γ-2: frm.Controls の挙動(直下のみ / frame子も flatten 込み)は環境差があるため、
    ' frm.Controls と frame の子を両方走査しつつ name で重複排除する(二重列挙=二重保存を防ぐ)。
    Dim col As Collection
    Set col = New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim c As Object, cc As Object
    For Each c In frm.Controls
        If Not seen.Exists(c.Name) Then
            col.Add c
            seen(c.Name) = True
        End If
        If TypeName(c) = "Frame" Then
            For Each cc In c.Controls
                If Not seen.Exists(cc.Name) Then
                    col.Add cc
                    seen(cc.Name) = True
                End If
            Next cc
        End If
    Next c
    Set AllFormControls = col
End Function

' Read field values from the active dynamic form and persist as a
' knowledge file. Returns the knowledgeId on success, "" on failure.
' mode = "register" -> generate next knowledgeId
' mode = "edit"     -> reuse current knowledgeId from renderer
Private Function PersistFromActiveForm(ByVal mode As String) As String
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Function
    Dim frm As Object
    Set frm = GetActiveDynForm()
    If frm Is Nothing Then Exit Function

    Dim fmtId As String
    fmtId = m_renderer.GetFormatId()
    ' For register: fmtId may have been changed in the form. Phase R-2 F-3:
    ' read whichever selector control exists (textbox or dropdown).
    On Error Resume Next
    Dim ctlFmt As Object
    Set ctlFmt = Nothing
    Set ctlFmt = frm.Controls("txtFormatId")
    If Not ctlFmt Is Nothing Then fmtId = CStr(ctlFmt.Text)
    Dim ctlCbo As Object
    Set ctlCbo = Nothing
    Set ctlCbo = frm.Controls("cboFormatId")
    If Not ctlCbo Is Nothing Then
        If Len(CStr(ctlCbo.Value)) > 0 Then fmtId = CStr(ctlCbo.Value)
    End If
    On Error GoTo ErrHandler

    Dim knowledgeId As String
    If mode = "edit" Then
        knowledgeId = m_renderer.GetKnowledgeId()
    Else
        knowledgeId = GenerateKnowledgeId(fmtId)
    End If

    ' Build content as ### header ### body for each field control
    Dim sb As String
    sb = "###KnowledgeNo###" & vbCrLf & knowledgeId & vbCrLf
    sb = sb & "###FormatID###" & vbCrLf & fmtId & vbCrLf

    ' Phase O-2: look up fieldName via renderer's ctlName -> fieldName map
    ' so multi-byte JP field names (???? / ???????? / etc) persist correctly.
    ' txtFormatId is the format selector and is NOT a knowledge field, skip.
    Dim ctl As Object
    For Each ctl In AllFormControls(frm)   ' R-3-χ-4: frScroll frame 内の field も列挙
        Dim nm As String
        nm = ctl.Name
        If Left$(nm, 4) = "ctl_" And nm <> "txtFormatId" Then
            Dim fldName As String
            fldName = m_renderer.GetFieldNameForCtl(nm)
            If Len(fldName) = 0 Then fldName = Mid$(nm, 5)
            Dim val As String
            If TypeName(ctl) = "TextBox" Then
                ' Phase R-2 F-4: a field still showing its placeholder = empty.
                If CStr(ctl.Tag) = "PLACEHOLDER" Then
                    val = ""
                Else
                    val = CStr(ctl.Text)
                End If
            ElseIf TypeName(ctl) = "ComboBox" Then
                val = CStr(ctl.Value)
            Else
                val = ""
            End If
            sb = sb & "###" & fldName & "###" & vbCrLf & val & vbCrLf
        End If
    Next ctl

    ' Write to data_dir using FileSystemObject so we control encoding.
    Dim filePath As String
    filePath = modConfigHolder.GetDataDir() & knowledgeId & ".txt"
    Dim adoStream As Object
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = 2
    adoStream.Charset = "Shift_JIS"
    adoStream.Open
    adoStream.WriteText sb
    adoStream.SaveToFile filePath, 2  ' overwrite
    adoStream.Close

    PersistFromActiveForm = knowledgeId
    Exit Function
ErrHandler:
    Debug.Print "[ERR] PersistFromActiveForm: " & Err.Number & " " & Err.Description
    PersistFromActiveForm = ""
End Function

' Test-mode public wrapper around GenerateKnowledgeId (Phase N-5 E2E).
Public Function TestGenerateKnowledgeId(ByVal formatId As String) As String
    TestGenerateKnowledgeId = GenerateKnowledgeId(formatId)
End Function

' Generate a knowledge id: <formatId>-<YYYY>-<NNNN>. NNNN scans data_dir
' for the largest existing suffix and uses +1.
Private Function GenerateKnowledgeId(ByVal formatId As String) As String
    Dim yr As String
    yr = Format$(Year(Now()), "0000")
    Dim prefix As String
    prefix = formatId & "-" & yr & "-"
    Dim maxNo As Long
    maxNo = 0
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim dir As String
    dir = modConfigHolder.GetDataDir()
    If Not fso.FolderExists(dir) Then
        GenerateKnowledgeId = prefix & "0001"
        Exit Function
    End If
    Dim f As Object
    For Each f In fso.GetFolder(dir).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "txt" Then
            Dim base As String
            base = fso.GetBaseName(f.Name)
            If InStr(1, base, prefix, vbTextCompare) = 1 Then
                Dim tail As String
                tail = Mid$(base, Len(prefix) + 1)
                If IsNumeric(tail) Then
                    If CLng(tail) > maxNo Then maxNo = CLng(tail)
                End If
            End If
        End If
    Next f
    GenerateKnowledgeId = prefix & Format$(maxNo + 1, "0000")
End Function

' Look up the currently shown dynamic UserForm via the renderer's stashed name.
Private Function GetActiveDynForm() As Object
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Function
    Dim nm As String
    nm = m_renderer.GetDynFormName()
    Dim uf As Object
    For Each uf In VBA.UserForms
        If uf.Name = nm Then
            Set GetActiveDynForm = uf
            Exit Function
        End If
    Next uf
    Set GetActiveDynForm = Nothing
    Exit Function
ErrHandler:
    Set GetActiveDynForm = Nothing
End Function
```
