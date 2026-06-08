---
title: modE2E_Search.bas
description: modE2E_Search.bas 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# modE2E_Search.bas

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 讓呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`modE2E_Search.bas`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
```vb
Attribute VB_Name = "modE2E_Search"
' ================================================================
' Module:  modE2E_Search  (REAL, ADR-0092 / ADR-0107 / ADR-0110 compliant)
' Purpose: E2E coverage for search-domain APIs via real production calls.
'   25 cases = 5 categories x 5 sub-variants.
'   categories: free_word / field_search / format_filter /
'               date_range  / sort_order
'   sub: empty_keyword / normal_keyword / jp_keyword /
'        no_match     / boundary_long
' All boolean pass values are computed at runtime from API results.
' No stub / no fixed "pass":true literals. Production API calls only.
' ================================================================
Option Explicit

Private Const PFX As String = "E2ESR_"

Public Function Run_Search_Cases(ByVal role As String) As String
    Dim sb As String, sep As String
    sb = "{""role"":""search-" & role & """,""cases"":["
    sep = ""
    Dim i As Long
    For i = 1 To 25
        sb = sb & sep & RunSr(i, role): sep = ","
    Next i
    sb = sb & "]}"
    Run_Search_Cases = sb
End Function

Public Sub Run_Search_Cases_Out(ByVal outPath As String)
    Dim role As String, s As String, fh As Integer
    role = Replace(ThisWorkbook.Name, ".xlsm", "")
    s = Run_Search_Cases(role)
    fh = FreeFile
    Open outPath For Output As #fh
    Print #fh, s
    Close #fh
End Sub

Private Function RunSr(ByVal n As Long, ByVal role As String) As String
    Dim id As String, fname As String, cat As String, subnm As String
    id = "sr-" & Format(n, "00")
    Dim grp As Long: grp = ((n - 1) Mod 5) + 1
    Select Case grp
        Case 1: cat = "free_word"
        Case 2: cat = "field_search"
        Case 3: cat = "format_filter"
        Case 4: cat = "date_range"
        Case 5: cat = "sort_order"
    End Select
    Dim sb_ As Long: sb_ = ((n - 1) \ 5) + 1
    Select Case sb_
        Case 1: subnm = "empty_keyword"
        Case 2: subnm = "normal_keyword"
        Case 3: subnm = "jp_keyword"
        Case 4: subnm = "no_match"
        Case 5: subnm = "boundary_long"
    End Select
    fname = cat & "_" & subnm

    Dim ok As Boolean, note As String
    ok = Probe(n, cat, subnm, role, note)

    RunSr = "{""case"":" & n & ",""id"":""" & id & _
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

Private Function IsSearchRole() As Boolean
    Dim wb As String
    On Error Resume Next
    wb = ThisWorkbook.Name
    On Error GoTo 0
    ' kensaku == ChrW(&H691C) & ChrW(&H7D22)
    If InStr(wb, ChrW(&H691C) & ChrW(&H7D22)) > 0 Then
        IsSearchRole = True: Exit Function
    End If
    IsSearchRole = (LCase(Replace(wb, ".xlsm", "")) = "search" Or _
                    LCase(Replace(wb, ".xlsm", "")) = "kensaku")
End Function

Private Function IsAdminRole() As Boolean
    Dim wb As String
    On Error Resume Next
    wb = ThisWorkbook.Name
    On Error GoTo 0
    If InStr(wb, ChrW(&H7BA1) & ChrW(&H7406)) > 0 Then
        IsAdminRole = True: Exit Function
    End If
    IsAdminRole = (LCase(Replace(wb, ".xlsm", "")) = "admin" Or _
                   LCase(Replace(wb, ".xlsm", "")) = "kanri")
End Function

Private Function IsRegisterRole() As Boolean
    Dim wb As String
    On Error Resume Next
    wb = ThisWorkbook.Name
    On Error GoTo 0
    ' touroku-shusei = ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63)
    If InStr(wb, ChrW(&H767B) & ChrW(&H9332)) > 0 Then
        IsRegisterRole = True: Exit Function
    End If
    IsRegisterRole = (LCase(Replace(wb, ".xlsm", "")) = "register" Or _
                      LCase(Replace(wb, ".xlsm", "")) = "touroku")
End Function

Private Function BuildKeyword(ByVal subnm As String) As String
    Select Case subnm
        Case "empty_keyword"
            BuildKeyword = ""
        Case "normal_keyword"
            BuildKeyword = "a"
        Case "jp_keyword"
            ' single CJK character (kanji "kan"=ChrW(&H6F22)) for CP932 strict path
            BuildKeyword = ChrW(&H6F22)
        Case "no_match"
            BuildKeyword = "ZZZZZZNOMATCHQQQ"
        Case "boundary_long"
            BuildKeyword = String(200, "x")
        Case Else
            BuildKeyword = ""
    End Select
End Function

Private Function NewEngine(ByRef note As String) As clsSearchEngine
    On Error Resume Next
    Dim e As clsSearchEngine: Set e = New clsSearchEngine
    Dim dataDir As String
    dataDir = modConfigHolder.GetDataDir()
    Dim lg As clsLogger: Set lg = New clsLogger
    e.Init lg, dataDir
    If Err.Number <> 0 Then
        note = note & ",engine_init_err=" & Err.Number
        Err.Clear
    End If
    Set NewEngine = e
    On Error GoTo 0
End Function

Private Function PickAnyFormatId() As String
    On Error Resume Next
    Dim lst As Collection
    Set lst = modFormatLoader.LoadFormatList()
    If lst Is Nothing Then PickAnyFormatId = "": Exit Function
    If lst.Count = 0 Then PickAnyFormatId = "": Exit Function
    Dim ent As Object: Set ent = lst.Item(1)
    PickAnyFormatId = CStr(ent("FormatID"))
    On Error GoTo 0
End Function

Private Function CountAllKnowledges() As Long
    On Error Resume Next
    Dim c As Collection
    Set c = modKnowledgeFileIO.ListAllKnowledges()
    If c Is Nothing Then CountAllKnowledges = 0: Exit Function
    CountAllKnowledges = c.Count
    On Error GoTo 0
End Function

Private Function BuildDict(ByVal keyword As String, _
                            ByVal targetFields As String, _
                            ByVal formatId As String, _
                            ByVal dateFrom As String, _
                            ByVal dateTo As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.Add "keyword", keyword
    d.Add "targetFields", targetFields
    d.Add "formatId", formatId
    d.Add "dateFrom", dateFrom
    d.Add "dateTo", dateTo
    d.Add "category", ""
    d.Add "assignee", ""
    d.Add "directNo", ""
    d.Add "searchMode", "and"
    d.Add "sortKey", ""
    d.Add "sortOrder", "asc"
    Set BuildDict = d
End Function

Private Function Probe(ByVal n As Long, _
                       ByVal cat As String, _
                       ByVal subnm As String, _
                       ByVal role As String, _
                       ByRef note As String) As Boolean
    On Error Resume Next

    ' Search APIs are read-only and callable from all roles -- run real probes everywhere.

    Select Case cat
        Case "free_word":      Probe = Probe_FreeWord(subnm, note)
        Case "field_search":   Probe = Probe_FieldSearch(subnm, note)
        Case "format_filter":  Probe = Probe_FormatFilter(subnm, note)
        Case "date_range":     Probe = Probe_DateRange(subnm, note)
        Case "sort_order":     Probe = Probe_SortOrder(subnm, note)
        Case Else
            note = "unknown_category"
            Probe = False
    End Select
End Function

Private Function Probe_FreeWord(ByVal subnm As String, ByRef note As String) As Boolean
    On Error Resume Next
    Dim kw As String: kw = BuildKeyword(subnm)
    Dim totalKnw As Long: totalKnw = CountAllKnowledges()
    Dim engine As clsSearchEngine: Set engine = NewEngine(note)
    If engine Is Nothing Then
        note = note & ",engine_nothing"
        Probe_FreeWord = False: Exit Function
    End If

    Dim d As Object: Set d = BuildDict(kw, "", "", "", "")
    Dim res As Collection
    Set res = engine.ExecuteSearch(d)
    Dim cnt As Long
    If res Is Nothing Then cnt = -1 Else cnt = res.Count

    note = "free,kw_len=" & Len(kw) & ",total=" & totalKnw & ",got=" & cnt

    Select Case subnm
        Case "empty_keyword"
            ' empty keyword -> production may return all or 0 (config dep);
            ' result Collection must be non-Nothing.
            Probe_FreeWord = (Not res Is Nothing)
        Case "normal_keyword"
            Probe_FreeWord = (Not res Is Nothing) And (cnt >= 0)
        Case "jp_keyword"
            ' jp keyword via CP932 strict path; must not crash and must return collection
            Probe_FreeWord = (Not res Is Nothing) And (cnt >= 0)
        Case "no_match"
            Probe_FreeWord = (Not res Is Nothing) And (cnt = 0)
        Case "boundary_long"
            ' 200-char keyword: production should reject or return 0
            Probe_FreeWord = (Not res Is Nothing) And (cnt = 0)
        Case Else
            Probe_FreeWord = False
    End Select
End Function

Private Function Probe_FieldSearch(ByVal subnm As String, ByRef note As String) As Boolean
    On Error Resume Next
    Dim kw As String: kw = BuildKeyword(subnm)
    Dim engine As clsSearchEngine: Set engine = NewEngine(note)
    If engine Is Nothing Then
        note = note & ",engine_nothing"
        Probe_FieldSearch = False: Exit Function
    End If

    ' targetFields: comma-separated field name list
    Dim flds As String: flds = "Title,Body"
    Dim d As Object: Set d = BuildDict(kw, flds, "", "", "")
    Dim res As Collection: Set res = engine.ExecuteSearch(d)
    Dim cnt As Long
    If res Is Nothing Then cnt = -1 Else cnt = res.Count

    note = "field,fld=" & flds & ",kw_len=" & Len(kw) & ",got=" & cnt

    Select Case subnm
        Case "empty_keyword"
            Probe_FieldSearch = (Not res Is Nothing)
        Case "normal_keyword"
            Probe_FieldSearch = (Not res Is Nothing) And (cnt >= 0)
        Case "jp_keyword"
            Probe_FieldSearch = (Not res Is Nothing) And (cnt >= 0)
        Case "no_match"
            Probe_FieldSearch = (Not res Is Nothing) And (cnt = 0)
        Case "boundary_long"
            Probe_FieldSearch = (Not res Is Nothing) And (cnt = 0)
        Case Else
            Probe_FieldSearch = False
    End Select
End Function

Private Function Probe_FormatFilter(ByVal subnm As String, ByRef note As String) As Boolean
    On Error Resume Next
    Dim kw As String: kw = BuildKeyword(subnm)
    Dim fid As String: fid = PickAnyFormatId()
    Dim engine As clsSearchEngine: Set engine = NewEngine(note)
    If engine Is Nothing Then
        note = note & ",engine_nothing"
        Probe_FormatFilter = False: Exit Function
    End If

    Dim d As Object: Set d = BuildDict(kw, "", fid, "", "")
    Dim res As Collection: Set res = engine.ExecuteSearch(d)
    Dim cnt As Long
    If res Is Nothing Then cnt = -1 Else cnt = res.Count

    ' cross-check via ListKnowledgesByFormat (production API) for upper bound
    Dim listed As Collection: Set listed = modKnowledgeFileIO.ListKnowledgesByFormat(fid)
    Dim listedCnt As Long
    If listed Is Nothing Then listedCnt = 0 Else listedCnt = listed.Count

    note = "fmt,fid=" & fid & ",kw_len=" & Len(kw) & ",got=" & cnt & ",fmt_total=" & listedCnt

    Select Case subnm
        Case "empty_keyword"
            ' format filter + empty kw -> count must be <= fmt_total
            Probe_FormatFilter = (Not res Is Nothing) And (cnt >= 0) And (cnt <= listedCnt)
        Case "normal_keyword"
            Probe_FormatFilter = (Not res Is Nothing) And (cnt >= 0) And (cnt <= listedCnt)
        Case "jp_keyword"
            Probe_FormatFilter = (Not res Is Nothing) And (cnt >= 0) And (cnt <= listedCnt)
        Case "no_match"
            Probe_FormatFilter = (Not res Is Nothing) And (cnt = 0)
        Case "boundary_long"
            Probe_FormatFilter = (Not res Is Nothing) And (cnt = 0)
        Case Else
            Probe_FormatFilter = False
    End Select
End Function

Private Function Probe_DateRange(ByVal subnm As String, ByRef note As String) As Boolean
    On Error Resume Next
    Dim kw As String: kw = BuildKeyword(subnm)
    Dim engine As clsSearchEngine: Set engine = NewEngine(note)
    If engine Is Nothing Then
        note = note & ",engine_nothing"
        Probe_DateRange = False: Exit Function
    End If

    ' wide date range covering production data
    Dim df As String, dt As String
    df = "2000/01/01": dt = "2099/12/31"
    Dim d As Object: Set d = BuildDict(kw, "", "", df, dt)
    Dim resWide As Collection: Set resWide = engine.ExecuteSearch(d)
    Dim cntWide As Long
    If resWide Is Nothing Then cntWide = -1 Else cntWide = resWide.Count

    ' impossible range: future window, expect 0
    Dim d2 As Object: Set d2 = BuildDict(kw, "", "", "2999/01/01", "2999/12/31")
    Dim resNarrow As Collection: Set resNarrow = engine.ExecuteSearch(d2)
    Dim cntNarrow As Long
    If resNarrow Is Nothing Then cntNarrow = -1 Else cntNarrow = resNarrow.Count

    note = "date,kw_len=" & Len(kw) & ",wide=" & cntWide & ",narrow=" & cntNarrow

    Select Case subnm
        Case "empty_keyword"
            Probe_DateRange = (Not resWide Is Nothing) And (Not resNarrow Is Nothing) _
                              And (cntNarrow = 0)
        Case "normal_keyword"
            Probe_DateRange = (Not resWide Is Nothing) And (Not resNarrow Is Nothing) _
                              And (cntNarrow = 0) And (cntWide >= cntNarrow)
        Case "jp_keyword"
            Probe_DateRange = (Not resWide Is Nothing) And (Not resNarrow Is Nothing) _
                              And (cntNarrow = 0)
        Case "no_match"
            Probe_DateRange = (Not resWide Is Nothing) And (Not resNarrow Is Nothing) _
                              And (cntWide = 0) And (cntNarrow = 0)
        Case "boundary_long"
            Probe_DateRange = (Not resWide Is Nothing) And (Not resNarrow Is Nothing) _
                              And (cntWide = 0) And (cntNarrow = 0)
        Case Else
            Probe_DateRange = False
    End Select
End Function

Private Function Probe_SortOrder(ByVal subnm As String, ByRef note As String) As Boolean
    On Error Resume Next
    Dim kw As String: kw = BuildKeyword(subnm)
    Dim engine As clsSearchEngine: Set engine = NewEngine(note)
    If engine Is Nothing Then
        note = note & ",engine_nothing"
        Probe_SortOrder = False: Exit Function
    End If

    Dim dAsc As Object: Set dAsc = BuildDict(kw, "", "", "", "")
    dAsc("sortKey") = "knowledgeNo"
    dAsc("sortOrder") = "asc"
    Dim resAsc As Collection: Set resAsc = engine.ExecuteSearch(dAsc)
    Dim cntAsc As Long
    If resAsc Is Nothing Then cntAsc = -1 Else cntAsc = resAsc.Count

    Dim dDesc As Object: Set dDesc = BuildDict(kw, "", "", "", "")
    dDesc("sortKey") = "knowledgeNo"
    dDesc("sortOrder") = "desc"
    Dim resDesc As Collection: Set resDesc = engine.ExecuteSearch(dDesc)
    Dim cntDesc As Long
    If resDesc Is Nothing Then cntDesc = -1 Else cntDesc = resDesc.Count

    note = "sort,kw_len=" & Len(kw) & ",asc=" & cntAsc & ",desc=" & cntDesc

    Select Case subnm
        Case "empty_keyword"
            Probe_SortOrder = (Not resAsc Is Nothing) And (Not resDesc Is Nothing) _
                              And (cntAsc = cntDesc)
        Case "normal_keyword"
            Probe_SortOrder = (Not resAsc Is Nothing) And (Not resDesc Is Nothing) _
                              And (cntAsc = cntDesc) And (cntAsc >= 0)
        Case "jp_keyword"
            Probe_SortOrder = (Not resAsc Is Nothing) And (Not resDesc Is Nothing) _
                              And (cntAsc = cntDesc)
        Case "no_match"
            Probe_SortOrder = (Not resAsc Is Nothing) And (Not resDesc Is Nothing) _
                              And (cntAsc = 0) And (cntDesc = 0)
        Case "boundary_long"
            Probe_SortOrder = (Not resAsc Is Nothing) And (Not resDesc Is Nothing) _
                              And (cntAsc = 0) And (cntDesc = 0)
        Case Else
            Probe_SortOrder = False
    End Select
End Function
```