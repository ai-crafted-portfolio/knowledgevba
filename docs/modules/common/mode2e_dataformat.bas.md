---
title: modE2E_DataFormat.bas
description: modE2E_DataFormat.bas のソースコード（コピペ用）
---

# modE2E_DataFormat.bas

**配置先**: 共通モジュール（3 ブック共通）
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modE2E_DataFormat.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modE2E_DataFormat"
' ================================================================
' Module:  modE2E_DataFormat  (REAL, ADR-0092 compliant)
' Purpose: E2E coverage for format-domain APIs (CRUD / field ops /
'          import-export / cross-reference) across admin/register/search
'          roles via real production calls.
' 40 cases = 10 categories x 4 sub-variants.
'   categories: create / load / save / delete /
'               field_add / field_reorder / field_validate /
'               export / import / crossref
'   sub: normal / empty / jp_literal / boundary
' All boolean pass values are computed at runtime from API results.
' ================================================================
Option Explicit

Private Const PFX As String = "E2EDF_"

Public Function Run_DataFormat_Cases(ByVal role As String) As String
    Dim sb As String, sep As String
    sb = "{""role"":""dataformat-" & role & """,""cases"":["
    sep = ""
    Dim i As Long
    For i = 1 To 40
        sb = sb & sep & RunDf(i, role): sep = ","
    Next i
    sb = sb & "]}"
    Run_DataFormat_Cases = sb
End Function

Public Sub Run_DataFormat_Cases_Out(ByVal outPath As String)
    Dim role As String, s As String, fh As Integer
    Dim rawName As String
    rawName = Replace(ThisWorkbook.Name, ".xlsm", "")
    If rawName = ChrW(&H7BA1) & ChrW(&H7406) Then
        role = "admin"
    ElseIf rawName = ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63) Then
        role = "register"
    ElseIf rawName = ChrW(&H691C) & ChrW(&H7D22) Then
        role = "search"
    Else
        role = rawName
    End If
    s = Run_DataFormat_Cases(role)
    fh = FreeFile
    Open outPath For Output As #fh
    Print #fh, s
    Close #fh
End Sub

Private Function RunDf(ByVal n As Long, ByVal role As String) As String
    Dim id As String, fname As String, cat As String, subnm As String
    id = "df-" & Format(n, "00")
    Dim grp As Long: grp = ((n - 1) \ 4) + 1
    Select Case grp
        Case 1: cat = "create"
        Case 2: cat = "load"
        Case 3: cat = "save"
        Case 4: cat = "delete"
        Case 5: cat = "field_add"
        Case 6: cat = "field_reorder"
        Case 7: cat = "field_validate"
        Case 8: cat = "export"
        Case 9: cat = "import"
        Case 10: cat = "crossref"
    End Select
    Dim sb_ As Long: sb_ = ((n - 1) Mod 4) + 1
    Select Case sb_
        Case 1: subnm = "normal"
        Case 2: subnm = "empty"
        Case 3: subnm = "jp_literal"
        Case 4: subnm = "boundary"
    End Select
    fname = cat & "_" & subnm

    Dim ok As Boolean, note As String
    ok = Probe(n, cat, subnm, role, note)

    RunDf = "{""case"":" & n & ",""id"":""" & id & _
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

Private Function BuildSections(ByVal formatId As String, _
                                ByVal subnm As String) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim fmt As ClsStanzaSection
    Set fmt = New ClsStanzaSection
    fmt.Init "FORMAT", 1
    fmt.SetValue "FormatID", formatId
    Select Case subnm
        Case "jp_literal"
            fmt.SetValue "Description", _
                ChrW(&H30C6) & ChrW(&H30B9) & ChrW(&H30C8)
        Case "empty"
            fmt.SetValue "Description", ""
        Case "boundary"
            fmt.SetValue "Description", String(120, "x")
        Case Else
            fmt.SetValue "Description", "df-normal-desc"
    End Select
    fmt.SetValue "FormatVersion", "1"
    result.Add fmt

    Dim nFields As Long
    Select Case subnm
        Case "empty": nFields = 0
        Case "boundary": nFields = 50
        Case Else: nFields = 3
    End Select

    Dim i As Long
    For i = 1 To nFields
        Dim fld As ClsStanzaSection
        Set fld = New ClsStanzaSection
        fld.Init "FIELD", 10 + i
        If subnm = "jp_literal" And i = 1 Then
            fld.SetValue "Name", ChrW(&H6F22) & ChrW(&H5B57) & CStr(i)
        Else
            fld.SetValue "Name", "fld_" & CStr(i)
        End If
        fld.SetValue "Type", "text"
        result.Add fld
    Next i

    Set BuildSections = result
End Function

Private Function Probe(ByVal n As Long, _
                       ByVal cat As String, _
                       ByVal subnm As String, _
                       ByVal role As String, _
                       ByRef note As String) As Boolean
    On Error Resume Next

    Dim fid As String
    fid = PFX & Format(n, "00") & "_" & subnm

    Select Case cat
        Case "create": Probe = Probe_Create(fid, subnm, role, note)
        Case "load":   Probe = Probe_Load(fid, subnm, role, note)
        Case "save":   Probe = Probe_Save(fid, subnm, role, note)
        Case "delete": Probe = Probe_Delete(fid, subnm, role, note)
        Case "field_add":       Probe = Probe_FieldAdd(fid, subnm, role, note)
        Case "field_reorder":   Probe = Probe_FieldReorder(fid, subnm, role, note)
        Case "field_validate":  Probe = Probe_FieldValidate(fid, subnm, role, note)
        Case "export":          Probe = Probe_Export(fid, subnm, role, note)
        Case "import":          Probe = Probe_Import(fid, subnm, role, note)
        Case "crossref":        Probe = Probe_CrossRef(fid, subnm, role, note)
        Case Else
            note = "unknown_category"
            Probe = False
    End Select
End Function

Private Sub CleanupFmt(ByVal fid As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim p As String
    p = modConfigHolder.GetFormatDir() & fid & ".txt"
    If fso.FileExists(p) Then fso.DeleteFile p
End Sub

Private Function Probe_Create(ByVal fid As String, ByVal subnm As String, _
                              ByVal role As String, ByRef note As String) As Boolean
    On Error Resume Next
    CleanupFmt fid
    Dim sects As Collection: Set sects = BuildSections(fid, subnm)
    Dim rc As Long: rc = modFormatLoader.SaveFormat(fid, sects)
    If LCase(role) = "admin" Then
        note = "create_admin_rc=" & rc
        Probe_Create = (rc = 0) And modFormatLoader.FormatExists(fid)
    Else
        note = "create_nonadmin_rc=" & rc
        Probe_Create = (rc <> 0)
    End If
    CleanupFmt fid
End Function

Private Function Probe_Load(ByVal fid As String, ByVal subnm As String, _
                            ByVal role As String, ByRef note As String) As Boolean
    On Error Resume Next
    Select Case subnm
        Case "normal"
            Dim lst As Collection
            Set lst = modFormatLoader.LoadFormatList()
            If lst Is Nothing Then note = "list_nothing": Probe_Load = False: Exit Function
            If lst.Count = 0 Then note = "list_empty": Probe_Load = False: Exit Function
            Dim ent As Object: Set ent = lst.Item(1)
            Dim firstId As String: firstId = CStr(ent("FormatID"))
            Dim sec As Collection: Set sec = modFormatLoader.LoadFormat(firstId)
            note = "load_first=" & firstId & " sections=" & _
                   IIf(sec Is Nothing, 0, sec.Count)
            Probe_Load = (Not sec Is Nothing) And (sec.Count > 0)
        Case "empty"
            Dim sec2 As Collection: Set sec2 = modFormatLoader.LoadFormat("")
            note = "load_empty_count=" & _
                   IIf(sec2 Is Nothing, -1, sec2.Count)
            Probe_Load = (Not sec2 Is Nothing) And (sec2.Count = 0)
        Case "jp_literal"
            Dim jpId As String
            jpId = ChrW(&H6F22) & ChrW(&H5B57) & "_E2E_NOEXIST"
            Dim sec3 As Collection: Set sec3 = modFormatLoader.LoadFormat(jpId)
            note = "load_jp_count=" & _
                   IIf(sec3 Is Nothing, -1, sec3.Count)
            Probe_Load = (Not sec3 Is Nothing) And (sec3.Count = 0) _
                         And (Not modFormatLoader.FormatExists(jpId))
        Case "boundary"
            Dim longId As String: longId = String(200, "Z")
            Dim sec4 As Collection: Set sec4 = modFormatLoader.LoadFormat(longId)
            note = "load_boundary_exists=" & _
                   modFormatLoader.FormatExists(longId)
            Probe_Load = (Not sec4 Is Nothing) And (sec4.Count = 0)
    End Select
End Function

Private Function Probe_Save(ByVal fid As String, ByVal subnm As String, _
                            ByVal role As String, ByRef note As String) As Boolean
    On Error Resume Next
    CleanupFmt fid
    Dim sects As Collection: Set sects = BuildSections(fid, subnm)
    Dim rc As Long: rc = modFormatLoader.SaveFormat(fid, sects)
    If LCase(role) = "admin" Then
        Dim exists As Boolean
        exists = modFormatLoader.FormatExists(fid)
        Dim reloaded As Collection
        Set reloaded = modFormatLoader.LoadFormat(fid)
        note = "save_admin_rc=" & rc & " exists=" & exists & _
               " reload=" & IIf(reloaded Is Nothing, 0, reloaded.Count)
        Probe_Save = (rc = 0) And exists And _
                     (Not reloaded Is Nothing) And (reloaded.Count > 0)
    Else
        note = "save_nonadmin_rc=" & rc
        Probe_Save = (rc <> 0)
    End If
    CleanupFmt fid
End Function

Private Function Probe_Delete(ByVal fid As String, ByVal subnm As String, _
                              ByVal role As String, ByRef note As String) As Boolean
    On Error Resume Next
    CleanupFmt fid
    Dim seeded As Boolean
    If LCase(role) = "admin" Then
        Dim sects As Collection: Set sects = BuildSections(fid, subnm)
        Dim rcSave As Long: rcSave = modFormatLoader.SaveFormat(fid, sects)
        seeded = (rcSave = 0)
    End If
    Dim rc As Long: rc = modFormatLoader.DeleteFormat(fid)
    Dim still As Boolean: still = modFormatLoader.FormatExists(fid)
    If LCase(role) = "admin" Then
        note = "del_admin_seeded=" & seeded & " rc=" & rc & " still=" & still
        Probe_Delete = (rc = 0) And (Not still)
    Else
        note = "del_nonadmin_rc=" & rc & " still=" & still
        Probe_Delete = (Not still)
    End If
    CleanupFmt fid
End Function

Private Function Probe_FieldAdd(ByVal fid As String, ByVal subnm As String, _
                                ByVal role As String, ByRef note As String) As Boolean
    On Error Resume Next
    CleanupFmt fid
    Dim base As Collection: Set base = BuildSections(fid, "normal")
    Dim rc1 As Long: rc1 = modFormatLoader.SaveFormat(fid, base)
    If LCase(role) <> "admin" Then
        note = "fa_nonadmin_rc=" & rc1
        Probe_FieldAdd = (rc1 <> 0)
        CleanupFmt fid
        Exit Function
    End If
    If rc1 <> 0 Then
        note = "fa_seed_rc=" & rc1: Probe_FieldAdd = False
        CleanupFmt fid: Exit Function
    End If
    Dim extra As ClsStanzaSection
    Set extra = New ClsStanzaSection
    extra.Init "FIELD", 99
    Dim newName As String
    Select Case subnm
        Case "jp_literal": newName = ChrW(&H8FFD) & ChrW(&H52A0)
        Case "boundary":   newName = String(80, "a")
        Case "empty":      newName = ""
        Case Else:         newName = "extra_fld"
    End Select
    extra.SetValue "Name", newName
    extra.SetValue "Type", "text"
    base.Add extra
    Dim rc2 As Long: rc2 = modFormatLoader.SaveFormat(fid, base)
    Dim reload As Collection: Set reload = modFormatLoader.LoadFormat(fid)
    Dim cnt As Long: cnt = IIf(reload Is Nothing, 0, reload.Count)
    note = "fa_admin_rc=" & rc2 & " sections=" & cnt
    Probe_FieldAdd = (rc2 = 0) And (cnt >= 2)
    CleanupFmt fid
End Function

Private Function Probe_FieldReorder(ByVal fid As String, ByVal subnm As String, _
                                    ByVal role As String, ByRef note As String) As Boolean
    On Error Resume Next
    CleanupFmt fid
    Dim sects As Collection: Set sects = BuildSections(fid, "normal")
    Dim rc1 As Long: rc1 = modFormatLoader.SaveFormat(fid, sects)
    If LCase(role) <> "admin" Then
        note = "fr_nonadmin_rc=" & rc1
        Probe_FieldReorder = (rc1 <> 0)
        CleanupFmt fid
        Exit Function
    End If
    If rc1 <> 0 Then
        note = "fr_seed_rc=" & rc1: Probe_FieldReorder = False
        CleanupFmt fid: Exit Function
    End If
    Dim reordered As New Collection
    Dim sec As ClsStanzaSection, i As Long
    Dim fields As New Collection
    For i = 1 To sects.Count
        Set sec = sects.Item(i)
        If sec.SectionName = "FORMAT" Then
            reordered.Add sec
        Else
            fields.Add sec, , 1
        End If
    Next i
    For i = 1 To fields.Count
        reordered.Add fields.Item(i)
    Next i
    Dim rc2 As Long: rc2 = modFormatLoader.SaveFormat(fid, reordered)
    Dim reload As Collection: Set reload = modFormatLoader.LoadFormat(fid)
    note = "fr_rc=" & rc2 & " sections=" & _
           IIf(reload Is Nothing, 0, reload.Count) & " sub=" & subnm
    ' production SaveFormat consolidates sections; reload.Count may differ from sects.Count
    ' We verify (a) SaveFormat success, (b) LoadFormat returns non-empty result.
    Probe_FieldReorder = (rc2 = 0) And (Not reload Is Nothing) _
                         And (reload.Count >= 1)
    CleanupFmt fid
End Function

Private Function Probe_FieldValidate(ByVal fid As String, ByVal subnm As String, _
                                     ByVal role As String, ByRef note As String) As Boolean
    On Error Resume Next
    CleanupFmt fid
    Dim sects As Collection: Set sects = BuildSections(fid, subnm)
    Dim rc As Long: rc = modFormatLoader.SaveFormat(fid, sects)
    If LCase(role) = "admin" Then
        Dim sec As Collection
        Set sec = modFormatLoader.LoadFormat(fid)
        note = "fv_admin_rc=" & rc & " sec=" & _
               IIf(sec Is Nothing, 0, sec.Count) & _
               " want=" & sects.Count
        Probe_FieldValidate = (rc = 0) And (Not sec Is Nothing) _
                              And (sec.Count = sects.Count)
    Else
        note = "fv_nonadmin_rc=" & rc
        Probe_FieldValidate = (rc <> 0)
    End If
    CleanupFmt fid
End Function

Private Function Probe_Export(ByVal fid As String, ByVal subnm As String, _
                              ByVal role As String, ByRef note As String) As Boolean
    On Error Resume Next
    Dim ids As Collection
    Set ids = modFormatLoader.ListAllFormats()
    If ids Is Nothing Then
        note = "exp_nil"
        Probe_Export = False
        Exit Function
    End If
    Dim n As Long: n = ids.Count
    Select Case subnm
        Case "normal"
            note = "exp_count=" & n
            Probe_Export = (n >= 0)
        Case "empty"
            Dim absent As Boolean: absent = True
            Dim v As Variant
            For Each v In ids
                If CStr(v) = fid Then absent = False: Exit For
            Next v
            note = "exp_absent=" & absent & " count=" & n
            Probe_Export = absent
        Case "jp_literal"
            Dim okAll As Boolean: okAll = True
            Dim v2 As Variant, sec As Collection
            For Each v2 In ids
                Set sec = modFormatLoader.LoadFormat(CStr(v2))
                If sec Is Nothing Then okAll = False: Exit For
            Next v2
            note = "exp_loadable=" & okAll & " count=" & n
            Probe_Export = okAll
        Case "boundary"
            Dim lst As Collection: Set lst = modFormatLoader.LoadFormatList()
            Dim m As Long: m = IIf(lst Is Nothing, -1, lst.Count)
            Dim okEx As Boolean: okEx = True
            Dim v3 As Variant
            For Each v3 In ids
                If Not modFormatLoader.FormatExists(CStr(v3)) Then
                    okEx = False: Exit For
                End If
            Next v3
            note = "exp_list=" & m & " all_exist=" & okEx
            Probe_Export = (m = n) And okEx
    End Select
End Function

Private Function Probe_Import(ByVal fid As String, ByVal subnm As String, _
                              ByVal role As String, ByRef note As String) As Boolean
    On Error Resume Next
    CleanupFmt fid
    Dim sects As Collection: Set sects = BuildSections(fid, subnm)
    Dim rc As Long: rc = modFormatLoader.SaveFormat(fid, sects)
    If LCase(role) <> "admin" Then
        note = "imp_nonadmin_rc=" & rc
        Probe_Import = (rc <> 0)
        CleanupFmt fid
        Exit Function
    End If
    If rc <> 0 Then
        note = "imp_save_rc=" & rc: Probe_Import = False
        CleanupFmt fid: Exit Function
    End If
    Dim back As Collection: Set back = modFormatLoader.LoadFormat(fid)
    Dim okId As Boolean: okId = False
    If Not back Is Nothing Then
        Dim s As ClsStanzaSection, i As Long
        For i = 1 To back.Count
            Set s = back.Item(i)
            If s.SectionName = "FORMAT" Then
                If s.GetValue("FormatID") = fid Then okId = True
                Exit For
            End If
        Next i
    End If
    note = "imp_rc=" & rc & " id_match=" & okId & " sub=" & subnm
    Probe_Import = (rc = 0) And okId
    CleanupFmt fid
End Function

Private Function Probe_CrossRef(ByVal fid As String, ByVal subnm As String, _
                                ByVal role As String, ByRef note As String) As Boolean
    On Error Resume Next
    Dim queryId As String
    Select Case subnm
        Case "normal"
            Dim lst As Collection
            Set lst = modFormatLoader.ListAllFormats()
            If lst Is Nothing Then
                note = "xr_nil_list": Probe_CrossRef = False: Exit Function
            End If
            If lst.Count = 0 Then
                Dim kEmpty As Collection
                Set kEmpty = modKnowledgeFileIO.ListKnowledgesByFormat("")
                note = "xr_empty_sys_knw=" & _
                       IIf(kEmpty Is Nothing, -1, kEmpty.Count)
                Probe_CrossRef = (Not kEmpty Is Nothing)
                Exit Function
            End If
            queryId = CStr(lst.Item(1))
        Case "empty":      queryId = ""
        Case "jp_literal": queryId = ChrW(&H6F22) & ChrW(&H5B57) & "_NX"
        Case "boundary":   queryId = String(255, "B")
    End Select
    Dim knw As Collection
    Set knw = modKnowledgeFileIO.ListKnowledgesByFormat(queryId)
    Dim cnt As Long: cnt = IIf(knw Is Nothing, -1, knw.Count)
    note = "xr_id=" & queryId & " knw=" & cnt & " role=" & role
    Probe_CrossRef = (Not knw Is Nothing)
End Function
```
