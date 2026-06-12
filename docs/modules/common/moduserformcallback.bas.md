---
title: modUserFormCallback.bas
description: modUserFormCallback.bas のソースコード（コピペ用）
---

# modUserFormCallback.bas

**配置先**: 共通モジュール（3 ブック共通）
**種類**: 標準モジュール
**更新日**: 2026-06-12 08:51

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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1580] modUserFormCallback.SetRenderer ENTER"  ' [ADR-0100]
    Set m_renderer = r
End Sub

' Called from dynamic form's btnRegister_Click
Public Sub OnRegister()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1581] modUserFormCallback.OnRegister ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Sub
    Dim newId As String
    newId = PersistFromActiveForm("register")
    m_renderer.SetReturnId newId
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1582] modUserFormCallback.OnRegister EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1583] modUserFormCallback.OnRegister EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] OnRegister: " & Err.Number & " " & Err.Description
End Sub

' [BUG-B12 2026-06-11] Register with completion message + form kept open
' (spec M-05). Called from the dynamic form's btnRegister_Click. Returns
' True when a knowledge file was created (newId non-empty).
Public Function OnRegisterV2() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1581b] modUserFormCallback.OnRegisterV2 ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    OnRegisterV2 = False
    If m_renderer Is Nothing Then Exit Function
    ' [BUG-B16 2026-06-11] enforce required fields before persisting.
    Dim missMsg As String
    missMsg = MissingRequiredMessage()
    If Len(missMsg) > 0 Then
        If Not modCommon.IsHeadless() Then MsgBox missMsg, vbExclamation, MsgRegisterTitle()
        Exit Function
    End If
    ' [B32 2026-06-12] block invalid date values (user decision Q4)
    Dim badDateMsg As String
    badDateMsg = InvalidDateMessage()
    If Len(badDateMsg) > 0 Then
        If Not modCommon.IsHeadless() Then MsgBox badDateMsg, vbExclamation, MsgRegisterTitle()
        Exit Function
    End If
    Dim newId As String
    newId = PersistFromActiveForm("register")
    m_renderer.SetReturnId newId
    If Len(newId) > 0 Then
        OnRegisterV2 = True
        If Not modCommon.IsHeadless() Then
            MsgBox MsgRegisterDone() & newId, vbInformation, MsgRegisterTitle()
        End If
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1582b] modUserFormCallback.OnRegisterV2 EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1583b] modUserFormCallback.OnRegisterV2 EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] OnRegisterV2: " & Err.Number & " " & Err.Description
End Function

' [BUG-B12] JP literal via ChrW per ADR-0006/0090/0094: "TOUROKU shimashita: "
Private Function MsgRegisterDone() As String
    MsgRegisterDone = ChrW(&H767B) & ChrW(&H9332) & ChrW(&H3057) & ChrW(&H307E) & _
                      ChrW(&H3057) & ChrW(&H305F) & ": "
End Function

' [BUG-B12] title "nareggi touroku"
Private Function MsgRegisterTitle() As String
    MsgRegisterTitle = ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H767B) & ChrW(&H9332)
End Function

' [BUG-B16 2026-06-11] Returns a non-empty warning listing empty required
' fields (or "" when all required fields are filled). Required flag comes
' from the renderer's per-control map (set at render time from the format).
Private Function MissingRequiredMessage() As String
    On Error GoTo ErrHandler
    MissingRequiredMessage = ""
    If m_renderer Is Nothing Then Exit Function
    Dim frm As Object
    Set frm = GetActiveDynForm()
    If frm Is Nothing Then Exit Function
    Dim missing As String
    missing = ""
    Dim ctl As Object
    For Each ctl In AllFormControls(frm)
        Dim nm As String
        nm = ctl.Name
        If Left(nm, 4) = "ctl_" And nm <> "txtFormatId" Then
            If m_renderer.IsRequiredCtl(nm) Then
                Dim v As String
                v = ""
                If TypeName(ctl) = "TextBox" Then
                    If CStr(ctl.Tag) <> "PLACEHOLDER" Then v = Trim(CStr(ctl.Text))
                ElseIf TypeName(ctl) = "ComboBox" Then
                    v = Trim(CStr(ctl.Value))
                End If
                If Len(v) = 0 Then
                    Dim fn As String
                    fn = m_renderer.GetFieldNameForCtl(nm)
                    If Len(fn) = 0 Then fn = Mid(nm, 5)
                    If Len(missing) > 0 Then missing = missing & vbCrLf
                    missing = missing & "・" & fn
                End If
            End If
        End If
    Next ctl
    If Len(missing) > 0 Then
        MissingRequiredMessage = MsgRequiredHead() & vbCrLf & missing
    End If
    Exit Function
ErrHandler:
    MissingRequiredMessage = ""
End Function

' [BUG-B16] head line: "hissu koumoku wo nyuuryoku shite kudasai:"
Private Function MsgRequiredHead() As String
    MsgRequiredHead = ChrW(&H5FC5) & ChrW(&H9808) & ChrW(&H9805) & ChrW(&H76EE) & _
                      ChrW(&H3092) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H3057) & _
                      ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ":"
End Function

' Called from dynamic form's btnUpdate_Click
Public Sub OnUpdate()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1584] modUserFormCallback.OnUpdate ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Sub
    ' [BUG-B16 2026-06-11] enforce required fields before persisting.
    Dim missMsg2 As String
    missMsg2 = MissingRequiredMessage()
    If Len(missMsg2) > 0 Then
        If Not modCommon.IsHeadless() Then MsgBox missMsg2, vbExclamation, MsgRegisterTitle()
        Exit Sub
    End If
    Dim badDateMsg2 As String
    badDateMsg2 = InvalidDateMessage()
    If Len(badDateMsg2) > 0 Then
        If Not modCommon.IsHeadless() Then MsgBox badDateMsg2, vbExclamation, MsgRegisterTitle()
        Exit Sub
    End If
    Dim id As String
    id = PersistFromActiveForm("edit")
    m_renderer.SetReturnId id
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1585] modUserFormCallback.OnUpdate EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1586] modUserFormCallback.OnUpdate EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] OnUpdate: " & Err.Number & " " & Err.Description
End Sub

' Called from dynamic form's btnDelete_Click
Public Sub OnDelete()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1587] modUserFormCallback.OnDelete ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Sub
    Dim kid As String
    kid = m_renderer.GetKnowledgeId()
    If Len(kid) = 0 Then Exit Sub
    ' [B31 2026-06-12] spec (ADR-0125/v2.3.1 banner): delete must CONFIRM first.
    ' Was: one click = silent physical delete.
    If Not modCommon.IsHeadless() Then
        Dim ansDel As VbMsgBoxResult
        ansDel = MsgBox(ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H300C) & kid & ChrW(&H300D) & ChrW(&H3092) & ChrW(&H524A) & ChrW(&H9664) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H304B) & ChrW(&HFF1F), _
            vbYesNo + vbExclamation + vbDefaultButton2, ChrW(&H30CA) & ChrW(&H30EC) & ChrW(&H30C3) & ChrW(&H30B8) & ChrW(&H524A) & ChrW(&H9664))
        If ansDel <> vbYes Then Exit Sub
    End If
    Dim ok As Boolean
    ok = modKnowledgeFileIO.DeleteKnowledge(kid)
    If ok Then m_renderer.SetReturnId "DELETED"
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1588] modUserFormCallback.OnDelete EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1589] modUserFormCallback.OnDelete EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] OnDelete: " & Err.Number & " " & Err.Description
End Sub

' Called from dynamic form's btnEdit_Click (M-09 view -> M-06 edit)
Public Sub OnEdit()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1590] modUserFormCallback.OnEdit ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Sub
    Dim kid As String
    kid = m_renderer.GetKnowledgeId()
    m_renderer.SetReturnId "EDIT:" & kid
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1591] modUserFormCallback.OnEdit EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1592] modUserFormCallback.OnEdit EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] OnEdit: " & Err.Number & " " & Err.Description
End Sub

' Called from dynamic form's btnClear_Click (M-05 register)
Public Sub OnClear()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1593] modUserFormCallback.OnClear ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim frm As Object
    Set frm = GetActiveDynForm()
    If frm Is Nothing Then Exit Sub
    Dim ctl As Object
    For Each ctl In frm.Controls
        Dim nm As String
        nm = ctl.Name
        ' Phase R-2: never clear the format selector or readonly header fields.
        If nm = "txtFormatId" Or nm = "cboFormatId" Or Left(nm, 7) = "hdrctl_" Then
            ' skip
        ElseIf TypeName(ctl) = "TextBox" Then
            ctl.Text = ""
        ElseIf TypeName(ctl) = "ComboBox" Then
            ctl.Value = ""
        End If
    Next ctl
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1594] modUserFormCallback.OnClear EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1595] modUserFormCallback.OnClear EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] OnClear: " & Err.Number & " " & Err.Description
End Sub

' Phase R-2 F-3: called from dynamic form's cboFormatId_Change. Reads the newly
' selected format and, if it differs from the current one, unloads the form so
' ShowForm's loop rebuilds it with the new format's field set.
Public Sub OnFormatChange()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1596] modUserFormCallback.OnFormatChange ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Sub
    Dim frm As Object
    Set frm = GetActiveDynForm()
    If frm Is Nothing Then Exit Sub
    Dim newFmt As String
    On Error Resume Next
    newFmt = CStr(frm.Controls("cboFormatId").Value)
    On Error GoTo ErrHandler
    ' 2026-06-11 fix: dropdown carries the display name; map back to FormatID.
    newFmt = m_renderer.ResolveDisplayToFormatId(newFmt)
    Dim before As String
    before = m_renderer.GetFormatId()
    m_renderer.RequestReformat newFmt
    If m_renderer.GetFormatId() <> before Then Unload frm
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1597] modUserFormCallback.OnFormatChange EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1598] modUserFormCallback.OnFormatChange EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] OnFormatChange: " & Err.Number & " " & Err.Description
End Sub

' R-3-a (2026-05-28): called from dynamic form's btnLoad_Click (M-06 edit).
' Reads the entered knowledge number and asks the renderer to reload + re-render.
Public Sub OnLoad()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1599] modUserFormCallback.OnLoad ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Sub
    Dim frm As Object
    Set frm = GetActiveDynForm()
    If frm Is Nothing Then Exit Sub
    Dim no As String
    On Error Resume Next
    no = CStr(FindCtl(frm, "txtKnowledgeNo").Text)   ' R-3-χ-4: frScroll 内
    On Error GoTo ErrHandler
    If Len(Trim(no)) = 0 Then Exit Sub
    m_renderer.RequestLoad no
    Unload frm
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1600] modUserFormCallback.OnLoad EXIT-OK"  ' [ADR-0100]
    Exit Sub
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1601] modUserFormCallback.OnLoad EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] OnLoad: " & Err.Number & " " & Err.Description
End Sub

' Phase R-3-χ-4: field 系 controls は frScroll Frame 内にある。frm.Controls は直下のみ
' なので、frame 子も含めて探索/列挙するヘルパを使う(1 段ネスト = frScroll)。
Private Function FindCtl(ByVal frm As Object, ByVal nm As String) As Object
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1602] modUserFormCallback.FindCtl ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim c As Object
    Set c = frm.Controls(nm)
    If Not c Is Nothing Then
        Set FindCtl = c
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1603] modUserFormCallback.FindCtl EXIT-OK"  ' [ADR-0100]
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
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1604] modUserFormCallback.FindCtl EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
        End If
    Next f
End Function

Private Function AllFormControls(ByVal frm As Object) As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1605] modUserFormCallback.AllFormControls ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1606] modUserFormCallback.AllFormControls EXIT-OK"  ' [ADR-0100]
End Function

' Read field values from the active dynamic form and persist as a
' knowledge file. Returns the knowledgeId on success, "" on failure.
' mode = "register" -> generate next knowledgeId
' mode = "edit"     -> reuse current knowledgeId from renderer
Private Function PersistFromActiveForm(ByVal mode As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1607] modUserFormCallback.PersistFromActiveForm ENTER"  ' [ADR-0100]
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
        ' 2026-06-11 fix: dropdown value is the display name; resolve to FormatID.
        If Len(CStr(ctlCbo.Value)) > 0 Then fmtId = m_renderer.ResolveDisplayToFormatId(CStr(ctlCbo.Value))
    End If
    On Error GoTo ErrHandler

    ' [B23 2026-06-11] guard: a format must be selected before persisting
    ' (was: empty FormatID produced orphan files like "-2026-0001.txt").
    If mode <> "edit" And Len(Trim$(fmtId)) = 0 Then
        If Not modCommon.IsHeadless() Then
            MsgBox ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H3092) & ChrW(&H9078) & ChrW(&H629E) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002), vbExclamation, MsgRegisterTitle()
        End If
        Exit Function
    End If

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
    ' [B30 2026-06-12] write UpdatedAt (and CreatedAt on register) so the
    ' M-08 result grid's update-date column and its descending sort work
    ' (search reads the in-file UpdatedAt key, not the file mtime).
    If mode <> "edit" Then
        sb = sb & "###CreatedAt###" & vbCrLf & Format$(Now, "yyyy-mm-dd hh:nn") & vbCrLf
    End If
    sb = sb & "###UpdatedAt###" & vbCrLf & Format$(Now, "yyyy-mm-dd hh:nn") & vbCrLf

    ' Phase O-2: look up fieldName via renderer's ctlName -> fieldName map
    ' so multi-byte JP field names (???? / ???????? / etc) persist correctly.
    ' txtFormatId is the format selector and is NOT a knowledge field, skip.
    Dim ctl As Object
    For Each ctl In AllFormControls(frm)   ' R-3-χ-4: frScroll frame 内の field も列挙
        Dim nm As String
        nm = ctl.Name
        If Left(nm, 4) = "ctl_" And nm <> "txtFormatId" Then
            Dim fldName As String
            fldName = m_renderer.GetFieldNameForCtl(nm)
            If Len(fldName) = 0 Then fldName = Mid(nm, 5)
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1608] modUserFormCallback.PersistFromActiveForm EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1609] modUserFormCallback.PersistFromActiveForm EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Debug.Print "[ERR] PersistFromActiveForm: " & Err.Number & " " & Err.Description
    PersistFromActiveForm = ""
End Function

' Test-mode public wrapper around GenerateKnowledgeId (Phase N-5 E2E).
Public Function TestGenerateKnowledgeId(ByVal formatId As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1610] modUserFormCallback.TestGenerateKnowledgeId ENTER"  ' [ADR-0100]
    TestGenerateKnowledgeId = GenerateKnowledgeId(formatId)
End Function

' Generate a knowledge id: <formatId>-<YYYY>-<NNNN>. NNNN scans data_dir
' for the largest existing suffix and uses +1.
Private Function GenerateKnowledgeId(ByVal formatId As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1611] modUserFormCallback.GenerateKnowledgeId ENTER"  ' [ADR-0100]
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
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1612] modUserFormCallback.GenerateKnowledgeId EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim f As Object
    For Each f In fso.GetFolder(dir).Files
        If LCase(fso.GetExtensionName(f.Name)) = "txt" Then
            Dim base As String
            base = fso.GetBaseName(f.Name)
            If InStr(1, base, prefix, vbTextCompare) = 1 Then
                Dim tail As String
                tail = Mid(base, Len(prefix) + 1)
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1613] modUserFormCallback.GetActiveDynForm ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    If m_renderer Is Nothing Then Exit Function
    Dim nm As String
    nm = m_renderer.GetDynFormName()
    Dim uf As Object
    For Each uf In VBA.UserForms
        If uf.Name = nm Then
            Set GetActiveDynForm = uf
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1614] modUserFormCallback.GetActiveDynForm EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next uf
    Set GetActiveDynForm = Nothing
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1615] modUserFormCallback.GetActiveDynForm EXIT-OK"  ' [ADR-0100]
    Exit Function
ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1616] modUserFormCallback.GetActiveDynForm EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Set GetActiveDynForm = Nothing
End Function

' [B32 2026-06-12] user decision Q4: invalid date values must BLOCK the
' persist (the red background was display-only). Mirrors MissingRequiredMessage.
Private Function InvalidDateMessage() As String
    On Error GoTo ErrHandler
    InvalidDateMessage = ""
    If m_renderer Is Nothing Then Exit Function
    Dim frm As Object
    Set frm = GetActiveDynForm()
    If frm Is Nothing Then Exit Function
    Dim bad As String
    bad = ""
    Dim ctl As Object
    For Each ctl In AllFormControls(frm)
        Dim nmD As String
        nmD = ctl.Name
        If Left(nmD, 4) = "ctl_" Then
            If m_renderer.IsDateCtl(nmD) Then
                Dim vD As String
                vD = ""
                If TypeName(ctl) = "TextBox" Then
                    If CStr(ctl.Tag) <> "PLACEHOLDER" Then vD = Trim(CStr(ctl.Text))
                End If
                If Len(vD) > 0 And Not IsDate(vD) Then
                    Dim fnD As String
                    fnD = m_renderer.GetFieldNameForCtl(nmD)
                    If Len(fnD) = 0 Then fnD = Mid(nmD, 5)
                    If Len(bad) > 0 Then bad = bad & vbCrLf
                    bad = bad & "・" & fnD
                End If
            End If
        End If
    Next ctl
    If Len(bad) > 0 Then
        InvalidDateMessage = ChrW(&H65E5) & ChrW(&H4ED8) & ChrW(&H306E) & ChrW(&H5F62) & ChrW(&H5F0F) & ChrW(&H304C) & ChrW(&H6B63) & ChrW(&H3057) & ChrW(&H304F) & ChrW(&H3042) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H003A) & vbCrLf & bad
    End If
    Exit Function
ErrHandler:
    InvalidDateMessage = ""
End Function
```
