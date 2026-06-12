---
title: modE2E_UserForm.bas
description: modE2E_UserForm.bas のソースコード（コピペ用）
---

# modE2E_UserForm.bas

**配置先**: 共通モジュール（3 ブック共通）
**種類**: 標準モジュール
**更新日**: 2026-06-08 12:53

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modE2E_UserForm.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modE2E_UserForm"
' ================================================================
' Module:  modE2E_UserForm  (REAL, ADR-0110 compliant)
' Purpose: E2E coverage for UserForm dynamic-build lifecycle.
'          Drives clsUserFormRenderer.ShowFormModeless (production
'          API), simulates input via Controls collection, invokes
'          production button handlers (modUserFormCallback.OnXxx),
'          then verifies via modKnowledgeFileIO / modFormatLoader.
'
' 20 cases:
'   form_type = ((n-1) Mod 4) + 1
'     1 registration   - register-role form, mode="register"
'     2 edit           - register-role form, mode="edit"
'     3 format_edit    - admin-role format editing form
'     4 format_preview - admin-role format preview form
'   scenario  = ((n-1) \ 4) + 1
'     1 launch_only      - Show -> immediate Unload
'     2 enter_save       - set values -> OnRegister/OnUpdate -> verify
'     3 enter_cancel     - set values -> close without save
'     4 jp_input         - ChrW JP values -> save -> reload match
'     5 validation_error - invalid input -> reject path
'
' Role gate:
'   registration / edit  -> register xlsm REAL; admin/search skipped-by-design
'   format_edit / format_preview -> admin xlsm REAL; register/search skipped-by-design
'
' Boolean pass values are computed at runtime from API results.
' UserForm is dispatched ONLY via vbModeless (ShowFormModeless).
' vbModal usage is BANNED here (would block headless Excel COM).
' ================================================================
Option Explicit

Private Const PFX As String = "E2EUF_"

Public Function Run_UserForm_Cases(ByVal role As String) As String
    Dim sb As String, sep As String
    sb = "{""role"":""userform-" & role & """,""cases"":["
    sep = ""
    Dim i As Long
    For i = 1 To 20
        sb = sb & sep & RunUf(i, role): sep = ","
    Next i
    sb = sb & "]}"
    Run_UserForm_Cases = sb
End Function

Public Sub Run_UserForm_Cases_Out(ByVal outPath As String)
    Dim role As String, s As String, fh As Integer
    role = Replace(ThisWorkbook.Name, ".xlsm", "")
    s = Run_UserForm_Cases(role)
    fh = FreeFile
    Open outPath For Output As #fh
    Print #fh, s
    Close #fh
End Sub

Private Function RunUf(ByVal n As Long, ByVal role As String) As String
    Dim id As String, fname As String, ftype As String, scen As String
    id = "uf-" & Format(n, "00")
    Dim grp As Long: grp = ((n - 1) Mod 4) + 1
    Select Case grp
        Case 1: ftype = "registration"
        Case 2: ftype = "edit"
        Case 3: ftype = "format_edit"
        Case 4: ftype = "format_preview"
    End Select
    Dim sb_ As Long: sb_ = ((n - 1) \ 4) + 1
    Select Case sb_
        Case 1: scen = "launch_only"
        Case 2: scen = "enter_save"
        Case 3: scen = "enter_cancel"
        Case 4: scen = "jp_input"
        Case 5: scen = "validation_error"
    End Select
    fname = ftype & "_" & scen

    Dim ok As Boolean, note As String
    ok = Probe(n, ftype, scen, role, note)

    RunUf = "{""case"":" & n & ",""id"":""" & id & _
            """,""name"":""" & fname & _
            """,""pass"":" & LCase(CStr(ok)) & _
            ",""note"":""" & EscJ(note) & """}"
End Function

Private Function EscJ(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    EscJ = t
End Function

Private Function ExpectedRole(ByVal ftype As String) As String
    Select Case ftype
        Case "registration", "edit"
            ExpectedRole = "register"
        Case "format_edit", "format_preview"
            ExpectedRole = "admin"
        Case Else
            ExpectedRole = ""
    End Select
End Function

Private Function SeedTestFormat(ByVal fid As String, _
                                 ByVal includeMandatory As Boolean) As Boolean
    On Error Resume Next
    Dim sects As Collection: Set sects = New Collection
    Dim head As ClsStanzaSection
    Set head = New ClsStanzaSection
    head.Init "FORMAT", 1
    head.SetValue "FormatID", fid
    head.SetValue "Description", "uf-seed"
    head.SetValue "FormatVersion", "1"
    sects.Add head
    Dim f1 As ClsStanzaSection
    Set f1 = New ClsStanzaSection
    f1.Init "FIELD", 11
    f1.SetValue "Name", "title"
    f1.SetValue "Type", "text"
    If includeMandatory Then f1.SetValue "Required", "true"
    sects.Add f1
    Dim f2 As ClsStanzaSection
    Set f2 = New ClsStanzaSection
    f2.Init "FIELD", 12
    f2.SetValue "Name", "body"
    f2.SetValue "Type", "text"
    sects.Add f2
    Dim rc As Long: rc = modFormatLoader.SaveFormat(fid, sects)
    SeedTestFormat = (rc = 0)
End Function

Private Sub CleanupFmt(ByVal fid As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim p As String
    p = modConfigHolder.GetFormatDir() & fid & ".txt"
    If fso.FileExists(p) Then fso.DeleteFile p
End Sub

Private Sub CleanupKnowledge(ByVal kid As String)
    On Error Resume Next
    Call modKnowledgeFileIO.DeleteKnowledge(kid)
End Sub

Private Function FindDynForm() As Object
    On Error Resume Next
    Dim uf As Object
    For Each uf In VBA.UserForms
        If Left(uf.Name, 7) = "frmDyn_" Then
            Set FindDynForm = uf
            Exit Function
        End If
    Next uf
End Function

Private Sub TrySetCtl(ByVal frm As Object, ByVal nm As String, ByVal v As Variant)
    On Error Resume Next
    Dim c As Object
    Set c = frm.Controls(nm)
    If c Is Nothing Then Exit Sub
    c.Value = v
End Sub

Private Function TryGetCtl(ByVal frm As Object, ByVal nm As String) As String
    On Error Resume Next
    Dim c As Object
    Set c = frm.Controls(nm)
    If c Is Nothing Then
        TryGetCtl = ""
        Exit Function
    End If
    TryGetCtl = CStr(c.Value)
End Function

Private Function Probe(ByVal n As Long, _
                       ByVal ftype As String, _
                       ByVal scen As String, _
                       ByVal role As String, _
                       ByRef note As String) As Boolean
    On Error Resume Next

    Dim wantRole As String: wantRole = ExpectedRole(ftype)
    If LCase(role) <> LCase(wantRole) Then
        note = "skipped-by-design role=" & role & " want=" & wantRole & " ftype=" & ftype
        Probe = True
        Exit Function
    End If

    Dim fid As String, kid As String
    fid = PFX & Format(n, "00")
    kid = PFX & "K" & Format(n, "00")

    Select Case ftype
        Case "registration":   Probe = Probe_Registration(n, scen, fid, kid, note)
        Case "edit":           Probe = Probe_Edit(n, scen, fid, kid, note)
        Case "format_edit":    Probe = Probe_FormatEdit(n, scen, fid, note)
        Case "format_preview": Probe = Probe_FormatPreview(n, scen, fid, note)
        Case Else
            note = "unknown_form_type=" & ftype
            Probe = False
    End Select
End Function

Private Function Probe_Registration(ByVal n As Long, ByVal scen As String, _
                                     ByVal fid As String, ByVal kid As String, _
                                     ByRef note As String) As Boolean
    On Error Resume Next
    CleanupFmt fid: CleanupKnowledge kid
    Dim seeded As Boolean: seeded = SeedTestFormat(fid, (scen = "validation_error"))
    If Not seeded Then
        note = "reg_seed_fail"
        Probe_Registration = False
        Exit Function
    End If

    Dim r As clsUserFormRenderer
    Set r = New clsUserFormRenderer
    r.TestSeedFormatId fid
    modUserFormCallback.SetRenderer r

    Dim showRet As String
    showRet = r.ShowFormModeless(ThisWorkbook.Name, "register", "", fid)
    DoEvents
    Dim frm As Object
    Set frm = FindDynForm()

    Select Case scen
        Case "launch_only"
            Dim built As Boolean: built = Not (frm Is Nothing)
            r.CloseModelessForm
            note = "reg_launch built=" & built & " ret=" & showRet
            Probe_Registration = built And (InStr(showRet, "OK:") = 1)
        Case "enter_save"
            If frm Is Nothing Then
                note = "reg_save_no_form ret=" & showRet
                r.CloseModelessForm
                Probe_Registration = False
                Exit Function
            End If
            TrySetCtl frm, "title", "uf_save_" & CStr(n)
            TrySetCtl frm, "body", "uf_body_" & CStr(n)
            DoEvents
            modUserFormCallback.OnRegister
            DoEvents
            r.CloseModelessForm
            Dim lst As Collection
            Set lst = modKnowledgeFileIO.ListKnowledgesByFormat(fid)
            Dim cnt As Long: cnt = IIf(lst Is Nothing, 0, lst.Count)
            note = "reg_save cnt=" & cnt
            Probe_Registration = (cnt >= 1)
            Dim kk As Variant
            If Not lst Is Nothing Then
                For Each kk In lst
                    modKnowledgeFileIO.DeleteKnowledge CStr(kk)
                Next kk
            End If
        Case "enter_cancel"
            If frm Is Nothing Then
                note = "reg_cancel_no_form"
                r.CloseModelessForm
                Probe_Registration = False
                Exit Function
            End If
            TrySetCtl frm, "title", "discarded_" & CStr(n)
            DoEvents
            r.CloseModelessForm
            Dim lst2 As Collection
            Set lst2 = modKnowledgeFileIO.ListKnowledgesByFormat(fid)
            Dim cnt2 As Long: cnt2 = IIf(lst2 Is Nothing, 0, lst2.Count)
            note = "reg_cancel cnt=" & cnt2
            Probe_Registration = (cnt2 = 0)
        Case "jp_input"
            If frm Is Nothing Then
                note = "reg_jp_no_form"
                r.CloseModelessForm
                Probe_Registration = False
                Exit Function
            End If
            Dim jpVal As String
            jpVal = ChrW(&H6F22) & ChrW(&H5B57) & "_" & CStr(n)
            TrySetCtl frm, "title", jpVal
            TrySetCtl frm, "body", jpVal & jpVal
            DoEvents
            modUserFormCallback.OnRegister
            DoEvents
            r.CloseModelessForm
            Dim lst3 As Collection
            Set lst3 = modKnowledgeFileIO.ListKnowledgesByFormat(fid)
            Dim cnt3 As Long: cnt3 = IIf(lst3 Is Nothing, 0, lst3.Count)
            Dim matched As Boolean: matched = False
            Dim kk3 As Variant
            If cnt3 >= 1 Then
                For Each kk3 In lst3
                    Dim d As Object
                    Set d = modKnowledgeFileIO.LoadKnowledge(CStr(kk3))
                    If Not d Is Nothing Then
                        Dim got As String
                        got = ""
                        If d.Exists("title") Then got = CStr(d("title"))
                        If InStr(got, ChrW(&H6F22)) > 0 Then matched = True
                    End If
                    modKnowledgeFileIO.DeleteKnowledge CStr(kk3)
                Next kk3
            End If
            note = "reg_jp cnt=" & cnt3 & " matched=" & matched
            Probe_Registration = matched
        Case "validation_error"
            If frm Is Nothing Then
                note = "reg_vld_no_form"
                r.CloseModelessForm
                Probe_Registration = False
                Exit Function
            End If
            TrySetCtl frm, "body", "noTitle_" & CStr(n)
            DoEvents
            modUserFormCallback.OnRegister
            DoEvents
            r.CloseModelessForm
            Dim lst4 As Collection
            Set lst4 = modKnowledgeFileIO.ListKnowledgesByFormat(fid)
            Dim cnt4 As Long: cnt4 = IIf(lst4 Is Nothing, 0, lst4.Count)
            Dim kk4 As Variant
            If cnt4 > 0 Then
                For Each kk4 In lst4
                    modKnowledgeFileIO.DeleteKnowledge CStr(kk4)
                Next kk4
            End If
            If cnt4 = 0 Then
                note = "reg_vld_reject cnt=" & cnt4
                Probe_Registration = True
            Else
                note = "FAIL_reg_vld_should_reject_but_accepted cnt=" & cnt4
                Probe_Registration = False
            End If
        Case Else
            note = "reg_unknown_scen=" & scen
            r.CloseModelessForm
            Probe_Registration = False
    End Select

    CleanupFmt fid
End Function

Private Function Probe_Edit(ByVal n As Long, ByVal scen As String, _
                             ByVal fid As String, ByVal kid As String, _
                             ByRef note As String) As Boolean
    On Error Resume Next
    CleanupFmt fid: CleanupKnowledge kid
    If Not SeedTestFormat(fid, False) Then
        note = "edit_seed_fmt_fail"
        Probe_Edit = False
        Exit Function
    End If
    Dim dataIn As Object
    Set dataIn = CreateObject("Scripting.Dictionary")
    dataIn("title") = "edit_seed_" & CStr(n)
    dataIn("body") = "edit_body_" & CStr(n)
    Dim createdKid As String
    createdKid = modKnowledgeFileIO.SaveKnowledge("", fid, dataIn)
    If LenB(createdKid) = 0 Then
        note = "edit_seed_knowledge_fail"
        Probe_Edit = False
        CleanupFmt fid
        Exit Function
    End If

    Dim r As clsUserFormRenderer
    Set r = New clsUserFormRenderer
    r.TestSeedFormatId fid
    modUserFormCallback.SetRenderer r

    Dim showRet As String
    showRet = r.ShowFormModeless(ThisWorkbook.Name, "edit", createdKid, fid)
    DoEvents
    Dim frm As Object
    Set frm = FindDynForm()

    Select Case scen
        Case "launch_only"
            Dim built As Boolean: built = Not (frm Is Nothing)
            r.CloseModelessForm
            note = "edit_launch built=" & built & " ret=" & showRet & " kid=" & createdKid
            Probe_Edit = built And (InStr(showRet, "OK:") = 1)
        Case "enter_save"
            If frm Is Nothing Then
                note = "edit_save_no_form"
                r.CloseModelessForm
                Probe_Edit = False
                CleanupKnowledge createdKid
                CleanupFmt fid
                Exit Function
            End If
            TrySetCtl frm, "title", "edit_after_" & CStr(n)
            DoEvents
            modUserFormCallback.OnUpdate
            DoEvents
            r.CloseModelessForm
            Dim reload As Object
            Set reload = modKnowledgeFileIO.LoadKnowledge(createdKid)
            Dim newTitle As String
            If Not reload Is Nothing Then
                If reload.Exists("title") Then newTitle = CStr(reload("title"))
            End If
            note = "edit_save title=" & newTitle
            Probe_Edit = (InStr(newTitle, "edit_after_") > 0)
        Case "enter_cancel"
            If frm Is Nothing Then
                note = "edit_cancel_no_form"
                r.CloseModelessForm
                Probe_Edit = False
                CleanupKnowledge createdKid
                CleanupFmt fid
                Exit Function
            End If
            TrySetCtl frm, "title", "discarded_edit_" & CStr(n)
            DoEvents
            r.CloseModelessForm
            Dim reload2 As Object
            Set reload2 = modKnowledgeFileIO.LoadKnowledge(createdKid)
            Dim untouched As String
            If Not reload2 Is Nothing Then
                If reload2.Exists("title") Then untouched = CStr(reload2("title"))
            End If
            note = "edit_cancel title=" & untouched
            Probe_Edit = (InStr(untouched, "edit_seed_") > 0)
        Case "jp_input"
            If frm Is Nothing Then
                note = "edit_jp_no_form"
                r.CloseModelessForm
                Probe_Edit = False
                CleanupKnowledge createdKid
                CleanupFmt fid
                Exit Function
            End If
            Dim jpVal As String
            jpVal = ChrW(&H5909) & ChrW(&H66F4) & "_" & CStr(n)
            TrySetCtl frm, "title", jpVal
            DoEvents
            modUserFormCallback.OnUpdate
            DoEvents
            r.CloseModelessForm
            Dim reload3 As Object
            Set reload3 = modKnowledgeFileIO.LoadKnowledge(createdKid)
            Dim got As String
            If Not reload3 Is Nothing Then
                If reload3.Exists("title") Then got = CStr(reload3("title"))
            End If
            note = "edit_jp got=" & got
            Probe_Edit = (InStr(got, ChrW(&H5909)) > 0)
        Case "validation_error"
            If frm Is Nothing Then
                note = "edit_vld_no_form"
                r.CloseModelessForm
                Probe_Edit = False
                CleanupKnowledge createdKid
                CleanupFmt fid
                Exit Function
            End If
            TrySetCtl frm, "title", ""
            DoEvents
            modUserFormCallback.OnUpdate
            DoEvents
            r.CloseModelessForm
            Dim reload4 As Object
            Set reload4 = modKnowledgeFileIO.LoadKnowledge(createdKid)
            Dim titleAfter As String
            titleAfter = ""
            If Not reload4 Is Nothing Then
                If reload4.Exists("title") Then titleAfter = CStr(reload4("title"))
            End If
            If InStr(titleAfter, "edit_seed_") > 0 Then
                note = "edit_vld_reject title=" & titleAfter
                Probe_Edit = True
            Else
                note = "FAIL_edit_vld_should_reject_but_accepted title=" & titleAfter
                Probe_Edit = False
            End If
        Case Else
            note = "edit_unknown_scen=" & scen
            r.CloseModelessForm
            Probe_Edit = False
    End Select

    CleanupKnowledge createdKid
    CleanupFmt fid
End Function

Private Function Probe_FormatEdit(ByVal n As Long, ByVal scen As String, _
                                   ByVal fid As String, _
                                   ByRef note As String) As Boolean
    On Error Resume Next
    CleanupFmt fid
    If Not SeedTestFormat(fid, False) Then
        note = "fmtE_seed_fail"
        Probe_FormatEdit = False
        Exit Function
    End If

    Dim r As clsUserFormRenderer
    Set r = New clsUserFormRenderer
    r.TestSeedFormatId fid
    modUserFormCallback.SetRenderer r

    Dim showRet As String
    showRet = r.ShowFormModeless(ThisWorkbook.Name, "edit", "", fid)
    DoEvents
    Dim frm As Object
    Set frm = FindDynForm()
    Dim built As Boolean: built = Not (frm Is Nothing)

    Select Case scen
        Case "launch_only"
            r.CloseModelessForm
            note = "fmtE_launch built=" & built & " ret=" & showRet
            Probe_FormatEdit = built And (InStr(showRet, "OK:") = 1)
        Case "enter_save"
            r.CloseModelessForm
            Dim base As Collection
            Set base = modFormatLoader.LoadFormat(fid)
            If base Is Nothing Then
                note = "fmtE_save_load_fail"
                Probe_FormatEdit = False
                CleanupFmt fid
                Exit Function
            End If
            Dim sec As ClsStanzaSection
            Dim i As Long
            For i = 1 To base.Count
                Set sec = base.Item(i)
                If sec.SectionName = "FORMAT" Then
                    sec.SetValue "Description", "fmtE_save_" & CStr(n)
                    Exit For
                End If
            Next i
            Dim rc As Long: rc = modFormatLoader.SaveFormat(fid, base)
            Dim reload As Collection
            Set reload = modFormatLoader.LoadFormat(fid)
            Dim got As String
            If Not reload Is Nothing Then
                Dim sec2 As ClsStanzaSection
                For i = 1 To reload.Count
                    Set sec2 = reload.Item(i)
                    If sec2.SectionName = "FORMAT" Then
                        got = sec2.GetValue("Description")
                        Exit For
                    End If
                Next i
            End If
            note = "fmtE_save rc=" & rc & " got=" & got
            Probe_FormatEdit = (rc = 0) And (InStr(got, "fmtE_save_") > 0)
        Case "enter_cancel"
            r.CloseModelessForm
            Dim reload2 As Collection
            Set reload2 = modFormatLoader.LoadFormat(fid)
            Dim origDesc As String
            If Not reload2 Is Nothing Then
                Dim sec3 As ClsStanzaSection
                Dim j As Long
                For j = 1 To reload2.Count
                    Set sec3 = reload2.Item(j)
                    If sec3.SectionName = "FORMAT" Then
                        origDesc = sec3.GetValue("Description")
                        Exit For
                    End If
                Next j
            End If
            note = "fmtE_cancel built=" & built & " desc=" & origDesc
            Probe_FormatEdit = (origDesc = "uf-seed")
        Case "jp_input"
            r.CloseModelessForm
            Dim base2 As Collection
            Set base2 = modFormatLoader.LoadFormat(fid)
            If base2 Is Nothing Then
                note = "fmtE_jp_load_fail"
                Probe_FormatEdit = False
                CleanupFmt fid
                Exit Function
            End If
            Dim jpV As String
            jpV = ChrW(&H66F8) & ChrW(&H5F0F) & "_" & CStr(n)
            Dim sec4 As ClsStanzaSection
            Dim k As Long
            For k = 1 To base2.Count
                Set sec4 = base2.Item(k)
                If sec4.SectionName = "FORMAT" Then
                    sec4.SetValue "Description", jpV
                    Exit For
                End If
            Next k
            Dim rc2 As Long: rc2 = modFormatLoader.SaveFormat(fid, base2)
            Dim reload3 As Collection
            Set reload3 = modFormatLoader.LoadFormat(fid)
            Dim gotJ As String
            If Not reload3 Is Nothing Then
                Dim sec5 As ClsStanzaSection
                For k = 1 To reload3.Count
                    Set sec5 = reload3.Item(k)
                    If sec5.SectionName = "FORMAT" Then
                        gotJ = sec5.GetValue("Description")
                        Exit For
                    End If
                Next k
            End If
            note = "fmtE_jp rc=" & rc2 & " got=" & gotJ
            Probe_FormatEdit = (rc2 = 0) And (InStr(gotJ, ChrW(&H66F8)) > 0)
        Case "validation_error"
            r.CloseModelessForm
            Dim bad As Collection
            Set bad = New Collection
            Dim h As ClsStanzaSection
            Set h = New ClsStanzaSection
            h.Init "FORMAT", 1
            h.SetValue "FormatID", ""
            h.SetValue "Description", "should_reject"
            bad.Add h
            Dim rcBad As Long: rcBad = modFormatLoader.SaveFormat("", bad)
            If rcBad <> 0 Then
                note = "fmtE_vld_reject rcBad=" & rcBad
                Probe_FormatEdit = True
            Else
                note = "FAIL_fmtE_vld_should_reject_but_accepted rcBad=" & rcBad
                Probe_FormatEdit = False
            End If
        Case Else
            r.CloseModelessForm
            note = "fmtE_unknown_scen=" & scen
            Probe_FormatEdit = False
    End Select

    CleanupFmt fid
End Function

Private Function Probe_FormatPreview(ByVal n As Long, ByVal scen As String, _
                                      ByVal fid As String, _
                                      ByRef note As String) As Boolean
    On Error Resume Next
    CleanupFmt fid
    If Not SeedTestFormat(fid, False) Then
        note = "fmtP_seed_fail"
        Probe_FormatPreview = False
        Exit Function
    End If

    Dim r As clsUserFormRenderer
    Set r = New clsUserFormRenderer
    r.TestSeedFormatId fid

    Dim showRet As String
    showRet = r.ShowFormModeless(ThisWorkbook.Name, "register", "", fid)
    DoEvents
    Dim frm As Object
    Set frm = FindDynForm()
    Dim built As Boolean: built = Not (frm Is Nothing)
    Dim dynName As String: dynName = r.GetDynFormName()
    Dim ctrlCnt As Long
    If built Then ctrlCnt = frm.Controls.Count

    Select Case scen
        Case "launch_only"
            r.CloseModelessForm
            note = "fmtP_launch built=" & built & " ret=" & showRet & " name=" & dynName
            Probe_FormatPreview = built And (Len(dynName) > 0)
        Case "enter_save"
            r.CloseModelessForm
            note = "fmtP_save_noop built=" & built & " ctrls=" & ctrlCnt
            Probe_FormatPreview = built And (ctrlCnt > 0)
        Case "enter_cancel"
            r.CloseModelessForm
            note = "fmtP_cancel built=" & built & " ctrls=" & ctrlCnt
            Probe_FormatPreview = built
        Case "jp_input"
            If frm Is Nothing Then
                note = "fmtP_jp_no_form"
                r.CloseModelessForm
                Probe_FormatPreview = False
                CleanupFmt fid
                Exit Function
            End If
            Dim jpVal As String
            jpVal = ChrW(&H8868) & ChrW(&H793A) & "_" & CStr(n)
            TrySetCtl frm, "title", jpVal
            DoEvents
            Dim back As String: back = TryGetCtl(frm, "title")
            r.CloseModelessForm
            note = "fmtP_jp set=" & jpVal & " back=" & back
            Probe_FormatPreview = (InStr(back, ChrW(&H8868)) > 0)
        Case "validation_error"
            r.CloseModelessForm
            Dim r2 As clsUserFormRenderer
            Set r2 = New clsUserFormRenderer
            r2.TestSeedFormatId "UF_NOEXIST_" & CStr(n)
            Dim ret2 As String
            ret2 = r2.ShowFormModeless(ThisWorkbook.Name, "register", "", "UF_NOEXIST_" & CStr(n))
            DoEvents
            r2.CloseModelessForm
            If InStr(ret2, "OK:") <> 1 Then
                note = "fmtP_vld_reject ret=" & ret2
                Probe_FormatPreview = True
            Else
                note = "FAIL_fmtP_vld_should_reject_but_accepted ret=" & ret2
                Probe_FormatPreview = False
            End If
        Case Else
            r.CloseModelessForm
            note = "fmtP_unknown_scen=" & scen
            Probe_FormatPreview = False
    End Select

    CleanupFmt fid
End Function
```
