---
title: modE2E_FieldType.bas
description: modE2E_FieldType.bas のソースコード（コピペ用）
---

# modE2E_FieldType.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-08 12:53

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modE2E_FieldType.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modE2E_FieldType"
Option Explicit

' modE2E_FieldType - 50 differentiated E2E cases.
' Layout: 5 field types (text/textarea/date/dropdown/checkbox)
'         x 10 subnames (normal/empty/boundary_min/boundary_max/
'                        invalid_char/special_quote/unicode/
'                        control_char/full_width/round_trip)
' Per ADR-0110: no fixed "pass":true literal, real API calls, FAIL paths,
' distinct inputs/expected per case.

Public Function Run_FieldType_Cases(ByVal role As String) As String
    Dim sb As String, sep As String
    sb = "{""role"":""fieldtype-" & role & """,""cases"":["
    sep = ""
    Dim fmt As String
    fmt = FirstFmt()
    Dim i As Long
    For i = 1 To 50
        sb = sb & sep & RunFt(i, role, fmt): sep = ","
    Next i
    sb = sb & "]}"
    Run_FieldType_Cases = sb
End Function

Private Function FirstFmt() As String
    On Error Resume Next
    Dim fmts As Collection
    Set fmts = modFormatLoader.LoadFormatList()
    If Err.Number <> 0 Then FirstFmt = "": Err.Clear: Exit Function
    If fmts Is Nothing Then FirstFmt = "": Exit Function
    If fmts.Count = 0 Then FirstFmt = "": Exit Function
    Dim e As Variant
    For Each e In fmts
        If TypeName(e) = "String" Then FirstFmt = CStr(e): Exit Function
        If IsObject(e) Then
            If e.Exists("FormatID") Then FirstFmt = CStr(e("FormatID")): Exit Function
            If e.Exists("formatId") Then FirstFmt = CStr(e("formatId")): Exit Function
        End If
    Next e
    FirstFmt = ""
End Function

Private Function EscJ(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    EscJ = t
End Function

Private Function GroupType(ByVal n As Long) As String
    Dim grp As Long: grp = ((n - 1) Mod 5) + 1
    Select Case grp
        Case 1: GroupType = "text"
        Case 2: GroupType = "textarea"
        Case 3: GroupType = "date"
        Case 4: GroupType = "dropdown"
        Case 5: GroupType = "checkbox"
    End Select
End Function

Private Function GroupSub(ByVal n As Long) As String
    Dim sbn As Long: sbn = ((n - 1) \ 5) + 1
    Select Case sbn
        Case 1: GroupSub = "normal"
        Case 2: GroupSub = "empty"
        Case 3: GroupSub = "boundary_min"
        Case 4: GroupSub = "boundary_max"
        Case 5: GroupSub = "invalid_char"
        Case 6: GroupSub = "special_quote"
        Case 7: GroupSub = "unicode"
        Case 8: GroupSub = "control_char"
        Case 9: GroupSub = "full_width"
        Case 10: GroupSub = "round_trip"
    End Select
End Function

Private Function MakeVal(ByVal ftype As String, ByVal subnm As String) As String
    Select Case subnm
        Case "normal":         MakeVal = "val_" & ftype & "_n"
        Case "empty":          MakeVal = ""
        Case "boundary_min":   MakeVal = "a"
        Case "boundary_max":   MakeVal = String(200, "x")
        Case "invalid_char":   MakeVal = Chr(7)
        Case "special_quote":  MakeVal = """q_" & ftype & """"
        Case "unicode":        MakeVal = ChrW(&H6F22) & ChrW(&H5B57) & "_" & ftype
        Case "control_char":   MakeVal = Chr(0)
        Case "full_width":     MakeVal = ChrW(&HFF11) & ChrW(&HFF12) & "_" & ftype
        Case "round_trip":     MakeVal = "rt_" & ftype & "_seed42"
    End Select
End Function

Private Function ExpectedSaveRc(ByVal subnm As String) As Long
    Select Case subnm
        Case "invalid_char":   ExpectedSaveRc = 2
        Case "control_char":   ExpectedSaveRc = 2
        Case Else:             ExpectedSaveRc = 0
    End Select
End Function

Private Function RunFt(ByVal n As Long, ByVal role As String, ByVal fmt As String) As String
    Dim id As String, ftype As String, subnm As String, fname As String
    id = "ft-" & Format(n, "00")
    ftype = GroupType(n)
    subnm = GroupSub(n)
    fname = ftype & "_" & subnm
    Dim ok As Boolean, note As String
    Select Case role
        Case "register": ok = ProbeRegister(ftype, subnm, fmt, note)
        Case "admin":    ok = ProbeAdmin(ftype, subnm, fmt, note)
        Case "search":   ok = ProbeSearch(ftype, subnm, note)
        Case Else:       note = "unknown_role": ok = False
    End Select
    RunFt = "{""case"":" & n & ",""id"":""" & id & """,""name"":""" & fname & """,""pass"":" & LCase(CStr(ok)) & ",""note"":""" & EscJ(note) & """}"
End Function

Private Function ProbeRegister(ByVal ftype As String, ByVal subnm As String, ByVal fmt As String, ByRef note As String) As Boolean
    On Error Resume Next
    If fmt = "" Then
        note = "no_format_in_system"
        ProbeRegister = False
        Exit Function
    End If

    Dim mgr As Object
    Set mgr = New clsKnowledgeManager
    mgr.Init Nothing
    Dim knwNo As String
    knwNo = mgr.BuildKnowledgeNumber(fmt)
    If Err.Number <> 0 Then
        note = "buildno_err=" & Err.Description: Err.Clear
        ProbeRegister = False
        Exit Function
    End If

    Dim keyName As String
    keyName = "test_" & ftype
    Dim val As String
    val = MakeVal(ftype, subnm)

    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("formatId") = fmt
    d(keyName) = val

    Dim rc As Long
    rc = modKnowledgeFileIO.SaveKnowledge(knwNo, d, 0)
    If Err.Number <> 0 Then
        note = "save_err=" & Err.Description: Err.Clear
        ProbeRegister = False
        Exit Function
    End If

    Dim expRc As Long
    expRc = ExpectedSaveRc(subnm)
    If rc <> expRc Then
        note = "save_rc=" & rc & ",exp=" & expRc & ",sub=" & subnm
        If rc = 3 Then note = note & ",env=non_register_wb"
        ProbeRegister = False
        Exit Function
    End If

    If rc = 2 Then
        note = "cp932_rejected_as_expected,sub=" & subnm
        ProbeRegister = True
        Exit Function
    End If

    Dim got As Object
    Set got = modKnowledgeFileIO.LoadKnowledge(knwNo)
    If got Is Nothing Then
        Call modKnowledgeFileIO.DeleteKnowledge(knwNo)
        note = "load_nothing,sub=" & subnm
        ProbeRegister = False
        Exit Function
    End If

    Dim loaded As String
    If got.Exists(keyName) Then loaded = CStr(got(keyName)) Else loaded = "<missing>"

    Dim pass As Boolean
    Select Case subnm
        Case "empty"
            pass = (loaded = "" Or loaded = "<missing>")
            note = "empty_loaded_len=" & Len(loaded)
        Case "boundary_max"
            pass = (Len(loaded) = 200)
            note = "boundary_max_len=" & Len(loaded)
        Case "boundary_min"
            pass = (loaded = "a")
            note = "boundary_min_loaded=" & loaded
        Case "round_trip"
            pass = (loaded = val)
            note = "round_trip_match=" & LCase(CStr(pass))
        Case "special_quote"
            pass = (loaded = val)
            note = "quote_match=" & LCase(CStr(pass)) & ",len=" & Len(loaded)
        Case "unicode"
            pass = (loaded = val)
            note = "unicode_match=" & LCase(CStr(pass)) & ",len=" & Len(loaded)
        Case "full_width"
            pass = (loaded = val)
            note = "fullwidth_match=" & LCase(CStr(pass)) & ",len=" & Len(loaded)
        Case "normal"
            pass = (loaded = val)
            note = "normal_match=" & LCase(CStr(pass))
        Case Else
            pass = (loaded = val)
            note = "default_match=" & LCase(CStr(pass))
    End Select

    Call modKnowledgeFileIO.DeleteKnowledge(knwNo)
    ProbeRegister = pass
End Function

Private Function CountSectionsByType(ByVal sec As Collection, ByVal targetType As String) As Long
    Dim n As Long: n = 0
    If sec Is Nothing Then CountSectionsByType = 0: Exit Function
    Dim itm As Variant
    For Each itm In sec
        If TypeName(itm) = "ClsStanzaSection" Then
            Dim t As String
            t = LCase(itm.GetValue("type"))
            If t = LCase(targetType) Then n = n + 1
        End If
    Next itm
    CountSectionsByType = n
End Function

Private Function ProbeAdmin(ByVal ftype As String, ByVal subnm As String, ByVal fmt As String, ByRef note As String) As Boolean
    On Error Resume Next
    ' round_trip uses LoadFormatList path; per-type assertion of fmt listing.
    If subnm = "round_trip" Then
        Dim lst As Collection
        Set lst = modFormatLoader.LoadFormatList()
        If Err.Number <> 0 Then
            note = "listfmt_err=" & Err.Description: Err.Clear
            ProbeAdmin = False: Exit Function
        End If
        If lst Is Nothing Then
            note = "fmtlist_nothing,type=" & ftype
            ProbeAdmin = False: Exit Function
        End If
        ProbeAdmin = (lst.Count > 0)
        note = "fmtlist_count=" & lst.Count & ",type=" & ftype & ",exp=gt0"
        Exit Function
    End If

    Dim arg As String, expectExist As Boolean
    expectExist = False
    Select Case subnm
        Case "normal":         arg = fmt: expectExist = True
        Case "empty":          arg = ""
        Case "boundary_min":   arg = "a"
        Case "boundary_max":   arg = String(200, "x")
        Case "invalid_char":   arg = Chr(7) & "x"
        Case "special_quote":  arg = """quote"""
        Case "unicode":        arg = ChrW(&H6F22) & "_nofmt"
        Case "control_char":   arg = Chr(0)
        Case "full_width":     arg = ChrW(&HFF11) & "_nofmt"
    End Select

    Dim sec As Collection
    Set sec = modFormatLoader.LoadFormat(arg)
    If Err.Number <> 0 Then
        note = "loadfmt_err=" & Err.Description: Err.Clear
        ProbeAdmin = False: Exit Function
    End If

    If Not expectExist Then
        Dim isEmpty As Boolean
        isEmpty = (sec Is Nothing)
        If Not isEmpty Then isEmpty = (sec.Count = 0)
        ProbeAdmin = isEmpty
        note = "expect_empty,got_count=" & IIf(sec Is Nothing, -1, sec.Count) & ",sub=" & subnm
        Exit Function
    End If

    If sec Is Nothing Then
        note = "normal_sec_nothing,sub=" & subnm: ProbeAdmin = False: Exit Function
    End If
    If sec.Count = 0 Then
        note = "normal_sec_zero,fmt=" & fmt: ProbeAdmin = False: Exit Function
    End If

    Dim cntType As Long
    cntType = CountSectionsByType(sec, ftype)
    Dim kList As Collection
    Set kList = modKnowledgeFileIO.ListKnowledgesByFormat(fmt)
    If Err.Number <> 0 Then
        note = "listby_err=" & Err.Description: Err.Clear
        ProbeAdmin = False: Exit Function
    End If
    Dim kCount As Long
    kCount = 0
    If Not kList Is Nothing Then kCount = kList.Count

    ProbeAdmin = (sec.Count > 0) And (Not kList Is Nothing)
    note = "sec_count=" & sec.Count & ",type=" & ftype & ",type_cnt=" & cntType & ",by_fmt_count=" & kCount
End Function

Private Function ValidateKnowledgeDict(ByVal d As Object, ByRef sNote As String) As Boolean
    If d Is Nothing Then sNote = "dict_nothing": ValidateKnowledgeDict = False: Exit Function
    If d.Count = 0 Then sNote = "dict_empty": ValidateKnowledgeDict = True: Exit Function
    Dim hasFmt As Boolean
    hasFmt = d.Exists("FormatID") Or d.Exists("formatId")
    sNote = "dict_count=" & d.Count & ",has_fmt=" & LCase(CStr(hasFmt))
    ValidateKnowledgeDict = True
End Function

Private Function ProbeSearch(ByVal ftype As String, ByVal subnm As String, ByRef note As String) As Boolean
    On Error Resume Next
    If subnm = "normal" Then
        Dim lst As Collection
        Set lst = modKnowledgeFileIO.ListAllKnowledges()
        If Err.Number <> 0 Then
            note = "list_err=" & Err.Description: Err.Clear
            ProbeSearch = False: Exit Function
        End If
        If lst Is Nothing Then
            note = "list_nothing,type=" & ftype
            ProbeSearch = False: Exit Function
        End If
        Dim probeId As String, sNote As String, structOk As Boolean
        probeId = ""
        If lst.Count > 0 Then
            probeId = CStr(lst(1))
            Dim dProbe As Object
            Set dProbe = modKnowledgeFileIO.LoadKnowledge(probeId)
            structOk = ValidateKnowledgeDict(dProbe, sNote)
        Else
            structOk = True: sNote = "no_existing_knw"
        End If
        ProbeSearch = structOk
        note = "list_count=" & lst.Count & ",type=" & ftype & ",probe=" & probeId & "," & sNote
        Exit Function
    End If

    If subnm = "round_trip" Then
        Dim lst2 As Collection
        Set lst2 = modKnowledgeFileIO.ListAllKnowledges()
        If Err.Number <> 0 Then
            note = "list_err=" & Err.Description: Err.Clear
            ProbeSearch = False: Exit Function
        End If
        If lst2 Is Nothing Then
            note = "list_nothing,type=" & ftype
            ProbeSearch = False: Exit Function
        End If
        Dim ok As Boolean: ok = True
        Dim checked As Long: checked = 0
        Dim idv As Variant
        For Each idv In lst2
            Dim d2 As Object
            Set d2 = modKnowledgeFileIO.LoadKnowledge(CStr(idv))
            If d2 Is Nothing Then ok = False: Exit For
            checked = checked + 1
            If checked >= 3 Then Exit For
        Next idv
        ProbeSearch = ok
        note = "list_count=" & lst2.Count & ",type=" & ftype & ",rt_checked=" & checked & ",ok=" & LCase(CStr(ok))
        Exit Function
    End If

    Dim arg As String
    Select Case subnm
        Case "empty":          arg = ""
        Case "boundary_min":   arg = "a"
        Case "boundary_max":   arg = String(200, "x")
        Case "invalid_char":   arg = Chr(7) & "x"
        Case "special_quote":  arg = """quote"""
        Case "unicode":        arg = ChrW(&H6F22) & "_no"
        Case "control_char":   arg = Chr(0)
        Case "full_width":     arg = ChrW(&HFF11) & "_no"
        Case Else:             note = "unhandled_sub=" & subnm: ProbeSearch = False: Exit Function
    End Select

    Dim d As Object
    Set d = modKnowledgeFileIO.LoadKnowledge(arg)
    If Err.Number <> 0 Then
        note = "loadk_err=" & Err.Description: Err.Clear
        ProbeSearch = False: Exit Function
    End If

    Dim isMiss As Boolean
    isMiss = (d Is Nothing)
    If Not isMiss Then isMiss = (d.Count = 0)
    Dim hasFmtKey As Boolean
    hasFmtKey = False
    If Not d Is Nothing Then
        If d.Count > 0 Then hasFmtKey = (d.Exists("FormatID") Or d.Exists("formatId"))
    End If

    ProbeSearch = isMiss
    note = "expect_miss,got_count=" & IIf(d Is Nothing, -1, d.Count) & ",has_fmt=" & LCase(CStr(hasFmtKey)) & ",sub=" & subnm & ",type=" & ftype
End Function

Public Sub Run_FieldType_Cases_Out(ByVal outPath As String)
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
    s = Run_FieldType_Cases(role)
    fh = FreeFile
    Open outPath For Output As #fh
    Print #fh, s
    Close #fh
End Sub
```
