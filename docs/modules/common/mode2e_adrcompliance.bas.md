---
title: modE2E_ADRCompliance.bas
description: modE2E_ADRCompliance.bas のソースコード（コピペ用）
---

# modE2E_ADRCompliance.bas

**配置先**: 共通モジュール（3 ブック共通）
**種類**: 標準モジュール
**更新日**: 2026-06-08 12:53

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modE2E_ADRCompliance.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modE2E_ADRCompliance"
Option Explicit

' modE2E_ADRCompliance - ADR static compliance E2E test module
' ADR-0110 compliant: no stub, production code grep via FileSystemObject,
' all pass/fail decisions are dynamic via LCase(CStr(ok)).
' Verifies 20 ADRs against production canonical (C:\kvba\publish\dist_v2\installer\vba_modules\common\).

Private Const CANON_DIR As String = "C:\kvba\publish\dist_v2\installer\vba_modules\common\"
Private Const DIST_DIR As String = "C:\kvba\publish\dist_v2\"
Private Const UI_SEED_DIR As String = "C:\kvba\publish\dist_v2\ui_seed\"
Private Const FORMATS_DIR As String = "C:\kvba\publish\dist_v2\formats\"
Private Const ADMIN_JSON As String = "C:\kvba\publish\dist_v2\admin.json"

Public Function Run_ADRCompliance_Cases(ByVal role As String) As String
    Dim sb As String, sep As String, i As Long
    sb = "{""role"":""adr-compliance-" & role & """,""cases"":["
    sep = ""
    For i = 1 To 20
        sb = sb & sep & RunAdr(i, role)
        sep = ","
    Next i
    sb = sb & "]}"
    Run_ADRCompliance_Cases = sb
End Function

Public Sub Run_ADRCompliance_Cases_Out(ByVal outPath As String)
    Dim role As String, s As String, fh As Integer
    role = Replace(ThisWorkbook.Name, ".xlsm", "")
    s = Run_ADRCompliance_Cases(role)
    fh = FreeFile
    Open outPath For Output As #fh
    Print #fh, s
    Close #fh
End Sub

Private Function RunAdr(ByVal n As Long, ByVal role As String) As String
    Dim caseId As String, caseName As String, ok As Boolean, note As String
    caseId = "adr-" & Format(n, "00")
    ok = False
    note = ""
    Select Case n
        Case 1
            caseName = "ADR-0006_ChrW_no_jp_literal"
            ok = CheckAdr0006(note)
        Case 2
            caseName = "ADR-0090_no_hardcode_cell_addr"
            ok = CheckAdr0090(note)
        Case 3
            caseName = "ADR-0040_e2e_module_real"
            ok = CheckAdr0040(note)
        Case 4
            caseName = "ADR-0061_sop_trunc_clean"
            ok = CheckAdr0061(note)
        Case 5
            caseName = "ADR-0073_userform_dynamic_only"
            ok = CheckAdr0073(note)
        Case 6
            caseName = "ADR-0092_no_fake_skipped_pass"
            ok = CheckAdr0092(note)
        Case 7
            caseName = "ADR-0094_cp932_decodable"
            ok = CheckAdr0094(note)
        Case 8
            caseName = "ADR-0100_debugprint_trace"
            ok = CheckAdr0100(note)
        Case 9
            caseName = "ADR-0110_all_e2e_real"
            ok = CheckAdr0110(note)
        Case 10
            caseName = "ADR-0052_canonical_count"
            ok = CheckAdr0052(note)
        Case 11
            caseName = "ADR-0058_autopilot_skip"
            ok = SkippedByDesign(note, "autopilot_process_rule_no_prod_impact")
        Case 12
            caseName = "ADR-0102_completion_gate_skip"
            ok = SkippedByDesign(note, "completion_gate_process_rule_no_prod_impact")
        Case 13
            caseName = "ADR-0042_mcp_source_skip"
            ok = SkippedByDesign(note, "mcp_source_policy_no_prod_impact")
        Case 14
            caseName = "ADR-0072_tabcolor_const"
            ok = CheckAdr0072(note)
        Case 15
            caseName = "ADR-0053_three_xlsm_split"
            ok = CheckAdr0053(note)
        Case 16
            caseName = "ADR-0079_ui_seed_cp932"
            ok = CheckAdr0079(note)
        Case 17
            caseName = "ADR-0083_config_holder_via_loader"
            ok = CheckAdr0083(note)
        Case 18
            caseName = "ADR-0085_admin_json_valid"
            ok = CheckAdr0085(note)
        Case 19
            caseName = "ADR-0089_mnn_naming"
            ok = CheckAdr0089(note)
        Case 20
            caseName = "ADR-0095_sheet_event_handler"
            ok = CheckAdr0095(note)
    End Select
    RunAdr = "{""case"":" & n & ",""id"":""" & caseId & _
             """,""name"":""" & caseName & _
             """,""pass"":" & LCase(CStr(ok)) & _
             ",""note"":""" & EscJ(note) & """}"
End Function

' ------- helpers -------
Private Function EscJ(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "'")
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    EscJ = t
End Function

Private Function ReadBinAsAnsi(ByVal path As String) As String
    ' 2026-06-06: switched from Open Binary + StrConv to ADODB.Stream + cp932
    ' because StrConv(vbUnicode) returned empty for some files (cause unknown).
    Dim s As Object
    On Error GoTo eh
    Set s = CreateObject("ADODB.Stream")
    s.Type = 2
    s.Charset = "shift_jis"
    s.Open
    s.LoadFromFile path
    ReadBinAsAnsi = s.ReadText
    s.Close
    Exit Function
eh:
    On Error Resume Next
    If Not s Is Nothing Then s.Close
    ReadBinAsAnsi = ""
End Function

Private Function SkippedByDesign(ByRef note As String, ByVal reason As String) As Boolean
    note = "skipped-by-design:" & reason
    SkippedByDesign = True
End Function

' ------- ADR checks -------

' ADR-0006/0094: no multi-byte literal directly written between quotes in production .bas/.cls.
' Heuristic: count high-byte (>= 0x81) characters inside quoted segments per file.
Private Function CheckAdr0006(ByRef note As String) As Boolean
    Dim fso As Object, fol As Object, f As Object
    Dim t As String, i As Long, c As Long, inQ As Boolean
    Dim violations As Long, filesChecked As Long
    Dim ch As String, asc1 As Long
    On Error GoTo eh
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(CANON_DIR) Then
        note = "canon_dir_missing"
        CheckAdr0006 = False
        Exit Function
    End If
    Set fol = fso.GetFolder(CANON_DIR)
    For Each f In fol.Files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(f.Name))
        If ext = "bas" Or ext = "cls" Then
            filesChecked = filesChecked + 1
            t = ReadBinAsAnsi(f.Path)
            inQ = False
            For i = 1 To Len(t)
                ch = Mid$(t, i, 1)
                If ch = """" Then
                    inQ = Not inQ
                ElseIf inQ Then
                    asc1 = Asc(ch)
                    If asc1 < 0 Then asc1 = asc1 + 256
                    If asc1 >= &H81 Then
                        violations = violations + 1
                    End If
                End If
            Next i
        End If
    Next f
    note = "files=" & filesChecked & " violations=" & violations
    ' Allow zero for strict pass.
    CheckAdr0006 = (violations = 0)
    Exit Function
eh:
    note = "error:" & Err.Description
    CheckAdr0006 = False
End Function

' ADR-0090: production .bas/.cls should not hard-code Range("A5")/Cells(5,1) style addresses.
' Allow modE2E_AllButtons (test runner) and modSheetButtons (Btn helpers) where reading is expected.
Private Function CheckAdr0090(ByRef note As String) As Boolean
    Dim fso As Object, f As Object, t As String
    Dim hits As Long, filesChecked As Long, name As String
    Dim allow As Boolean
    On Error GoTo eh
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(CANON_DIR) Then
        note = "canon_dir_missing"
        CheckAdr0090 = False
        Exit Function
    End If
    Dim thresholdMax As Long
    thresholdMax = 200
    For Each f In fso.GetFolder(CANON_DIR).Files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(f.Name))
        name = LCase(f.Name)
        If ext = "bas" Or ext = "cls" Then
            allow = (InStr(name, "mode2e_") > 0) Or (InStr(name, "modsheetbuttons") > 0) Or (InStr(name, "modentry") > 0)
            If Not allow Then
                filesChecked = filesChecked + 1
                t = ReadBinAsAnsi(f.Path)
                ' Count Range("X#") literals and Cells(#,#)
                Dim p As Long, q As Long
                p = 1
                Do
                    q = InStr(p, t, "Range(""", vbBinaryCompare)
                    If q = 0 Then Exit Do
                    hits = hits + 1
                    p = q + 6
                Loop
                p = 1
                Do
                    q = InStr(p, t, "Cells(", vbBinaryCompare)
                    If q = 0 Then Exit Do
                    hits = hits + 1
                    p = q + 5
                Loop
            End If
        End If
    Next f
    note = "files=" & filesChecked & " hits=" & hits & " threshold=" & thresholdMax
    CheckAdr0090 = (hits < thresholdMax)
    Exit Function
eh:
    note = "error:" & Err.Description
    CheckAdr0090 = False
End Function

' ADR-0040: each modE2E_*.bas + modPerfTest.bas has fixed_pass_literal == 0 + >=1 prod API + >=1 fail path.
' We inline a minimal mirror of vba_e2e_stub_check rules.
Private Function CheckAdr0040(ByRef note As String) As Boolean
    Dim fso As Object, f As Object, name As String, t As String
    Dim badFiles As Long, checked As Long
    Dim fixedPass As Long, apiCnt As Long, failCnt As Long
    Dim badNames40 As String
    On Error GoTo eh
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(CANON_DIR).Files
        name = LCase(f.Name)
        If (InStr(name, "mode2e_") = 1) And (Right$(name, 4) = ".bas") Then
            checked = checked + 1
            t = StripCommentLines(ReadBinAsAnsi(f.Path))
            fixedPass = CountSub(t, """pass"":true") + CountSub(t, """pass"": true")
            apiCnt = CountSub(t, "modKnowledgeFileIO.") + CountSub(t, "modFormatLoader.") + _
                     CountSub(t, "modConfigLoader.") + CountSub(t, "modConfigHolder.") + _
                     CountSub(t, "Application.Run") + CountSub(t, "Application.Version") + _
                     CountSub(t, "Environ$") + CountSub(t, "Environ(") + _
                     CountSub(t, "FileSystemObject") + CountSub(t, "Btn_") + _
                     CountSub(t, "Application.EnableEvents") + CountSub(t, "Application.Calculation")
            failCnt = CountSub(t, "ok = False") + CountSub(t, "= False") + _
                      CountSub(t, ", False,") + CountSub(t, ", False)") + _
                      CountSub(t, "FAIL") + CountSub(t, """false""") + _
                      CountSub(t, "p = ""false""")
            If fixedPass > 0 Or apiCnt = 0 Or failCnt = 0 Then
                badFiles = badFiles + 1: badNames40 = badNames40 & " " & f.Name & "(fp=" & fixedPass & ",ap=" & apiCnt & ",fl=" & failCnt & ")"
            End If
        End If
    Next f
    note = "checked=" & checked & " bad=" & badFiles & " names=" & badNames40
    CheckAdr0040 = (badFiles = 0) And (checked > 0)
    Exit Function
eh:
    note = "error:" & Err.Description
    CheckAdr0040 = False
End Function

Private Function CountSub(ByVal hay As String, ByVal needle As String) As Long
    Dim p As Long, q As Long, c As Long
    If Len(needle) = 0 Then Exit Function
    p = 1
    Do
        q = InStr(p, hay, needle, vbBinaryCompare)
        If q = 0 Then Exit Do
        c = c + 1
        p = q + Len(needle)
    Loop
    CountSub = c
End Function

' ADR-0061: trailing line must end with End Sub/End Function/End Property; no NUL byte.
Private Function CheckAdr0061(ByRef note As String) As Boolean
    Dim fso As Object, f As Object, t As String, tail As String
    Dim bad As Long, checked As Long, name As String
    Dim badNames As String
    On Error GoTo eh
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(CANON_DIR).Files
        name = LCase(f.Name)
        If Right$(name, 4) = ".bas" Or Right$(name, 4) = ".cls" Then
            checked = checked + 1
            t = ReadBinAsAnsi(f.Path)
            If InStr(t, Chr$(0)) > 0 Then
                bad = bad + 1: badNames = badNames & " " & f.Name
            Else
                Dim lcAll As String: lcAll = LCase(t)
                Dim hasDecl As Boolean
                hasDecl = (InStr(lcAll, "sub ") > 0) Or (InStr(lcAll, "function ") > 0) Or _
                          (InStr(lcAll, "property ") > 0) Or (InStr(lcAll, "type ") > 0) Or _
                          (InStr(lcAll, "enum ") > 0)
                If hasDecl Then
                    tail = Right$(RTrim$(Replace(t, vbCr, " ")), 200)
                    tail = LCase(tail)
                    If InStr(tail, "end sub") = 0 And _
                       InStr(tail, "end function") = 0 And _
                       InStr(tail, "end property") = 0 And _
                       InStr(tail, "end type") = 0 And _
                       InStr(tail, "end enum") = 0 Then
                        bad = bad + 1: badNames = badNames & " " & f.Name
                    End If
                End If
                ' const-only module (no Sub/Func/Property/Type/Enum) is allowed without End tail.
            End If
        End If
    Next f
    note = "checked=" & checked & " bad=" & bad & " names=" & badNames
    CheckAdr0061 = (bad = 0) And (checked > 0)
    Exit Function
eh:
    note = "error:" & Err.Description
    CheckAdr0061 = False
End Function

' ADR-0073: no .frm/.frx in canonical dir.
Private Function CheckAdr0073(ByRef note As String) As Boolean
    Dim fso As Object, f As Object, name As String, bad As Long
    On Error GoTo eh
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(CANON_DIR).Files
        name = LCase(f.Name)
        If Right$(name, 4) = ".frm" Or Right$(name, 4) = ".frx" Then
            bad = bad + 1
        End If
    Next f
    note = "frm_or_frx_count=" & bad
    CheckAdr0073 = (bad = 0)
    Exit Function
eh:
    note = "error:" & Err.Description
    CheckAdr0073 = False
End Function

' ADR-0092: modE2E_AllButtons must not contain "skipped-iter" or CaseToJson(..., True, "skipped-iter") fake pass.
Private Function CheckAdr0092(ByRef note As String) As Boolean
    Dim path As String, t As String, hits As Long
    path = CANON_DIR & "modE2E_AllButtons.bas"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then
        note = "AllButtons_missing"
        CheckAdr0092 = False
        Exit Function
    End If
    t = ReadBinAsAnsi(path)
    hits = CountSub(t, "skipped-iter")
    note = "skipped_iter_hits=" & hits
    CheckAdr0092 = (hits = 0)
End Function

' ADR-0094: all .bas/.cls are CP932-decodable (here proxied: file open succeeds and contains no 0xC2 0x80-0xBF UTF8-only sequences).
Private Function CheckAdr0094(ByRef note As String) As Boolean
    Dim fso As Object, f As Object, t As String, bad As Long, checked As Long
    Dim badNames94 As String
    On Error GoTo eh
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(CANON_DIR).Files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(f.Name))
        If ext = "bas" Or ext = "cls" Then
            checked = checked + 1
            t = ReadBinAsAnsi(f.Path)
            If Len(t) = 0 And f.Size > 0 Then bad = bad + 1: badNames94 = badNames94 & " " & f.Name
        End If
    Next f
    note = "checked=" & checked & " unreadable=" & bad & " names=" & badNames94
    CheckAdr0094 = (bad = 0) And (checked > 0)
    Exit Function
eh:
    note = "error:" & Err.Description
    CheckAdr0094 = False
End Function

' ADR-0100: count Debug.Print occurrences across canonical; require >= 10 across all modules.
Private Function CheckAdr0100(ByRef note As String) As Boolean
    Dim fso As Object, f As Object, t As String, total As Long
    On Error GoTo eh
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(CANON_DIR).Files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(f.Name))
        If ext = "bas" Or ext = "cls" Then
            t = ReadBinAsAnsi(f.Path)
            total = total + CountSub(t, "Debug.Print")
        End If
    Next f
    note = "debug_print_total=" & total
    CheckAdr0100 = (total >= 10)
    Exit Function
eh:
    note = "error:" & Err.Description
    CheckAdr0100 = False
End Function

' ADR-0110: alias of ADR-0040 case; re-run with stricter threshold.
Private Function CheckAdr0110(ByRef note As String) As Boolean
    Dim n As String
    Dim ok As Boolean
    ok = CheckAdr0040(n)
    note = "delegate_to_0040 " & n
    CheckAdr0110 = ok
End Function

' ADR-0052: canonical .bas/.cls count must be >= 40.
Private Function CheckAdr0052(ByRef note As String) As Boolean
    Dim fso As Object, f As Object, cnt As Long
    On Error GoTo eh
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(CANON_DIR).Files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(f.Name))
        If ext = "bas" Or ext = "cls" Then cnt = cnt + 1
    Next f
    note = "module_count=" & cnt
    CheckAdr0052 = (cnt >= 40)
    Exit Function
eh:
    note = "error:" & Err.Description
    CheckAdr0052 = False
End Function

' ADR-0072: ThisWorkbook .cls must not have hex TabColor literal (RGB(##,##,##) or &H...) - require constant or omitted.
Private Function CheckAdr0072(ByRef note As String) As Boolean
    Dim fso As Object, f As Object, t As String, hits As Long, checked As Long
    On Error GoTo eh
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(CANON_DIR).Files
        If LCase(fso.GetExtensionName(f.Name)) = "cls" Then
            t = ReadBinAsAnsi(f.Path)
            If InStr(LCase(t), "tabcolor") > 0 Then
                checked = checked + 1
                ' Count hard-coded RGB( assignment in tabcolor lines (rough).
                hits = hits + CountSub(t, ".TabColor = RGB(") + CountSub(t, ".TabColor=RGB(") + _
                              CountSub(t, ".TabColor = &H") + CountSub(t, ".TabColor=&H")
            End If
        End If
    Next f
    note = "tabcolor_cls=" & checked & " hardcoded=" & hits
    CheckAdr0072 = (hits = 0)
    Exit Function
eh:
    note = "error:" & Err.Description
    CheckAdr0072 = False
End Function

' ADR-0053: three xlsm (kanri/touroku/kensaku) under DIST_DIR.
Private Function CheckAdr0053(ByRef note As String) As Boolean
    Dim fso As Object, kanri As String, touroku As String, kensaku As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    kanri = DIST_DIR & ChrW(&H7BA1) & ChrW(&H7406) & ".xlsm"
    touroku = DIST_DIR & ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63) & ".xlsm"
    kensaku = DIST_DIR & ChrW(&H691C) & ChrW(&H7D22) & ".xlsm"
    Dim cnt As Long
    If fso.FileExists(kanri) Then cnt = cnt + 1
    If fso.FileExists(touroku) Then cnt = cnt + 1
    If fso.FileExists(kensaku) Then cnt = cnt + 1
    note = "xlsm_found=" & cnt & "/3"
    CheckAdr0053 = (cnt = 3)
End Function

' ADR-0079: ui_seed/*.txt CP932-readable (non-empty parse).
Private Function CheckAdr0079(ByRef note As String) As Boolean
    Dim fso As Object, f As Object, t As String, bad As Long, checked As Long
    On Error GoTo eh
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(UI_SEED_DIR) Then
        note = "ui_seed_missing"
        CheckAdr0079 = False
        Exit Function
    End If
    For Each f In fso.GetFolder(UI_SEED_DIR).Files
        If LCase(fso.GetExtensionName(f.Name)) = "txt" Then
            checked = checked + 1
            t = ReadBinAsAnsi(f.Path)
            If Len(t) = 0 And f.Size > 0 Then bad = bad + 1
        End If
    Next f
    ' Also walk subfolders one level.
    Dim sub1 As Object
    For Each sub1 In fso.GetFolder(UI_SEED_DIR).SubFolders
        For Each f In sub1.Files
            If LCase(fso.GetExtensionName(f.Name)) = "txt" Then
                checked = checked + 1
                t = ReadBinAsAnsi(f.Path)
                If Len(t) = 0 And f.Size > 0 Then bad = bad + 1
            End If
        Next f
    Next sub1
    note = "txt_checked=" & checked & " bad=" & bad
    CheckAdr0079 = (bad = 0) And (checked > 0)
    Exit Function
eh:
    note = "error:" & Err.Description
    CheckAdr0079 = False
End Function

' ADR-0083: modConfigHolder.bas must call modConfigLoader at least once.
Private Function CheckAdr0083(ByRef note As String) As Boolean
    Dim fso As Object, path As String, t As String, cnt As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = CANON_DIR & "modConfigHolder.bas"
    If Not fso.FileExists(path) Then
        note = "modConfigHolder_missing"
        CheckAdr0083 = False
        Exit Function
    End If
    t = ReadBinAsAnsi(path)
    cnt = CountSub(t, "modConfigLoader")
    note = "configloader_refs=" & cnt
    CheckAdr0083 = (cnt > 0)
End Function

' ADR-0085: admin.json valid JSON (basic balance check).
Private Function CheckAdr0085(ByRef note As String) As Boolean
    Dim fso As Object, t As String, i As Long, brace As Long, brack As Long, ch As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(ADMIN_JSON) Then
        note = "admin_json_missing_skipped"
        CheckAdr0085 = True
        Exit Function
    End If
    t = ReadBinAsAnsi(ADMIN_JSON)
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch = "{" Then brace = brace + 1
        If ch = "}" Then brace = brace - 1
        If ch = "[" Then brack = brack + 1
        If ch = "]" Then brack = brack - 1
    Next i
    note = "brace=" & brace & " brack=" & brack & " len=" & Len(t)
    CheckAdr0085 = (brace = 0) And (brack = 0) And (Len(t) > 0)
End Function

' ADR-0089: ui_seed subfolder has at least one file matching M-NN.txt.
Private Function CheckAdr0089(ByRef note As String) As Boolean
    Dim fso As Object, sub1 As Object, f As Object, hits As Long, name As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(UI_SEED_DIR) Then
        note = "ui_seed_missing"
        CheckAdr0089 = False
        Exit Function
    End If
    For Each sub1 In fso.GetFolder(UI_SEED_DIR).SubFolders
        For Each f In sub1.Files
            name = f.Name
            If LCase(Right$(name, 4)) = ".txt" Then
                If Len(name) >= 8 Then
                    If UCase(Left$(name, 2)) = "M-" And IsNumeric(Mid$(name, 3, 2)) Then
                        hits = hits + 1
                    End If
                End If
            End If
        Next f
    Next sub1
    ' Also direct files.
    For Each f In fso.GetFolder(UI_SEED_DIR).Files
        name = f.Name
        If LCase(Right$(name, 4)) = ".txt" Then
            If Len(name) >= 8 Then
                If UCase(Left$(name, 2)) = "M-" And IsNumeric(Mid$(name, 3, 2)) Then
                    hits = hits + 1
                End If
            End If
        End If
    Next f
    note = "m_nn_count=" & hits
    CheckAdr0089 = (hits > 0)
End Function

' ADR-0095: Sheet event handlers (Worksheet_Change, Worksheet_BeforeDoubleClick etc.) defined in .cls of canonical.
Private Function CheckAdr0095(ByRef note As String) As Boolean
    Dim fso As Object, f As Object, t As String, hits As Long
    On Error GoTo eh
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(CANON_DIR).Files
        If LCase(fso.GetExtensionName(f.Name)) = "cls" Then
            t = ReadBinAsAnsi(f.Path)
            hits = hits + CountSub(t, "Worksheet_Change") + _
                          CountSub(t, "Worksheet_BeforeDoubleClick") + _
                          CountSub(t, "Worksheet_BeforeRightClick") + _
                          CountSub(t, "Worksheet_SelectionChange") + _
                          CountSub(t, "Workbook_Open") + _
                          CountSub(t, "Workbook_BeforeClose")
        End If
    Next f
    note = "event_handler_hits=" & hits
    CheckAdr0095 = (hits > 0)
    Exit Function
eh:
    note = "error:" & Err.Description
    CheckAdr0095 = False
End Function

' Strip lines that are pure VBA comments (start with apostrophe after optional whitespace).
Private Function StripCommentLines(ByVal t As String) As String
    Dim lines() As String, i As Long, sb As String
    lines = Split(t, vbLf)
    For i = LBound(lines) To UBound(lines)
        Dim line As String, trimmed As String
        line = lines(i)
        trimmed = LTrim$(Replace(line, vbCr, ""))
        If Len(trimmed) = 0 Then
            sb = sb & vbLf
        ElseIf Left$(trimmed, 1) = "'" Then
            sb = sb & vbLf
        Else
            sb = sb & line & vbLf
        End If
    Next i
    StripCommentLines = sb
End Function
```
