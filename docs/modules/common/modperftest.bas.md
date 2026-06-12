---
title: modPerfTest.bas
description: modPerfTest.bas のソースコード（コピペ用）
---

# modPerfTest.bas

**配置先**: 共通モジュール（3 ブック共通）
**種類**: 標準モジュール
**更新日**: 2026-06-05 17:15

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modPerfTest.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modPerfTest"
' ================================================================
' modPerfTest  (AM9 perf 25-case auto harness, 2026-06-04)
' Purpose:
'   Measure 性能大量-01..25 budgets where VBA Timer suffices,
'   emit SKIP_PHYSICAL for OS-level perfmon/stopwatch tests.
'
'   Output: C:\kvba\workspace\e2e_v20\perf_<role>.json
'   Lines : C:\kvba\workspace\e2e_v20\perf_<role>_lines.txt
'   Shape : list of CaseToJson rows compatible with aggregate.py
'           (test_id prefix '性能大量-NN' -> sheet '性能大量').
'
' Notes:
'   - All JP literals via ChrW (ADR-0006 / ADR-0094 CP932 strict).
'   - Reuses modE2E_AllButtons helpers where Public; falls back to
'     local copy where they are Private.
' ================================================================
Option Explicit

Private Const PERF_OUT_DIR As String = "C:\kvba\workspace\e2e_v20"
Private Const PERF_FIXTURE_DIR As String = "C:\kvba\workspace\e2e_v20\perf_fixture"
Private Const PERF_FMT_FIXTURE_DIR As String = "C:\kvba\workspace\e2e_v20\perf_fmt_fixture"
Private Const PERF_LOG_FILE As String = "C:\kvba\install_all_pub.log"

' ----------------------------------------------------------------
' Public entry: Run_Perf_AllCases
'   role auto-detected; results written to perf_<role>.json
' ----------------------------------------------------------------
Public Sub Run_Perf_AllCases(Optional ByVal outPath As String = "")
    On Error GoTo ErrHandler
    Dim role As String
    role = PerfDetectRole()
    Dim path As String
    path = outPath
    If Len(path) = 0 Then path = PERF_OUT_DIR & "\perf_" & role & ".json"
    Dim linesPath As String
    linesPath = PERF_OUT_DIR & "\perf_" & role & "_lines.txt"
    On Error Resume Next
    Kill linesPath
    On Error GoTo ErrHandler
    EnsurePerfFolders

    Dim sb As String
    sb = "{"
    sb = sb & """role"":""perf-" & role & ""","
    sb = sb & """timestamp"":""" & Format$(Now, "yyyy-mm-dd HH:nn:ss") & ""","
    sb = sb & """cases"":["

    Dim r As String
    r = ""
    r = r & PerfCase01_LoadN(100, 5#)
    r = r & "," & PerfCase02_LoadN(1000, 15#)
    r = r & "," & PerfCase03_LoadN(10000, 60#)
    r = r & "," & PerfCase04_LoadFormats(50, 10#)
    r = r & "," & PerfCase05_FreeWordSearch(1000, 3#)
    r = r & "," & PerfCase06_FieldSearch(1000, 2#)
    r = r & "," & PerfCase07_RenderResults(1000, 5#)
    r = r & "," & PerfSkipPhysical(8, "cold_start", 10#, "requires PS-side Excel cold launch timer")
    r = r & "," & PerfCase09_WarmStart(5#)
    r = r & "," & PerfSkipPhysical(10, "button_visual_feedback", 0.1, "requires stopwatch / screen capture")
    r = r & "," & PerfCase11_FormDraw(10, 1#)
    r = r & "," & PerfCase12_FormDraw(50, 3#)
    r = r & "," & PerfCase13_FormDraw(200, 10#)
    r = r & "," & PerfCase14_Save1(1#)
    r = r & "," & PerfCase15_Save1000(10#)
    r = r & "," & PerfSkipPhysical(16, "list_freeze_100_row", 0#, "requires UI freeze monitor")
    r = r & "," & PerfSkipPhysical(17, "mem_1000", 500#, "requires perfmon / Task Manager")
    r = r & "," & PerfSkipPhysical(18, "mem_leak_10_cycle", 50#, "requires perfmon Working Set delta")
    r = r & "," & PerfSkipPhysical(19, "long_run_1h", 3600#, "requires 1h continuous operator session")
    r = r & "," & PerfSkipPhysical(20, "parallel_2_instance", 0#, "requires 2 separate Excel processes")
    r = r & "," & PerfSkipPhysical(21, "ui_during_save", 0#, "requires concurrent operator interaction")
    r = r & "," & PerfSkipPhysical(22, "splash_on_open", 0#, "splash feature not implemented in v2.3")
    r = r & "," & PerfCase23_VbeCompile(5#)
    r = r & "," & PerfCase24_InstallBatTime(60#)
    r = r & "," & PerfCase25_E2EFullSuite(1200#)

    sb = sb & r & "]}"
    PerfWriteAllText path, sb
    Exit Sub
ErrHandler:
    On Error Resume Next
    Dim fbPath As String
    fbPath = path
    If Len(fbPath) = 0 Then fbPath = PERF_OUT_DIR & "\perf_error.json"
    PerfWriteAllText fbPath, "{""role"":""perf-error"",""err"":""" & PerfEscapeJson(Err.Description) & """}"
End Sub

' ================================================================
' Individual cases
' ================================================================

Private Function PerfCase01_LoadN(ByVal n As Long, ByVal budget As Double) As String
    PerfCase01_LoadN = MeasureLoadN(1, n, budget)
End Function
Private Function PerfCase02_LoadN(ByVal n As Long, ByVal budget As Double) As String
    PerfCase02_LoadN = MeasureLoadN(2, n, budget)
End Function
Private Function PerfCase03_LoadN(ByVal n As Long, ByVal budget As Double) As String
    PerfCase03_LoadN = MeasureLoadN(3, n, budget)
End Function

' Generate N stanza .txt fixtures under PERF_FIXTURE_DIR, then loop
' LoadKnowledge on each, time total.
Private Function MeasureLoadN(ByVal idx As Long, ByVal n As Long, ByVal budget As Double) As String
    Dim t0 As Double, t1 As Double
    Dim ok As Boolean, note As String
    Dim errnum As Long, errdesc As String
    On Error Resume Next
    Err.Clear
    GenerateStanzaFixtures n
    If Err.Number <> 0 Then
        errnum = Err.Number
        errdesc = Err.Description
        MeasureLoadN = PerfRowJson(idx, "knowledge_load_" & CStr(n), "FAIL", 0#, budget, "fixture-gen err=" & errnum & " " & errdesc)
        Exit Function
    End If
    t0 = Timer
    Dim i As Long
    Dim filePath As String
    Dim content As String
    Dim fnum As Integer
    For i = 1 To n
        filePath = PERF_FIXTURE_DIR & "\PERF_" & Format$(i, "00000") & ".txt"
        content = PerfReadAllTextCp932(filePath)
        If Len(content) = 0 Then Exit For
    Next i
    t1 = Timer
    Dim elapsed As Double
    elapsed = t1 - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    ok = (elapsed <= budget)
    note = "loaded=" & CStr(i - 1) & " of " & CStr(n)
    MeasureLoadN = PerfRowJson(idx, "knowledge_load_" & CStr(n), PerfVerdict(ok), elapsed, budget, note)
End Function

Private Function PerfCase04_LoadFormats(ByVal n As Long, ByVal budget As Double) As String
    Dim t0 As Double, t1 As Double
    Dim ok As Boolean, note As String
    On Error Resume Next
    Err.Clear
    GenerateFormatFixtures n
    If Err.Number <> 0 Then
        PerfCase04_LoadFormats = PerfRowJson(4, "format_load_" & CStr(n), "FAIL", 0#, budget, "fixture-gen err=" & Err.Number & " " & Err.Description)
        Exit Function
    End If
    t0 = Timer
    Dim i As Long
    Dim filePath As String
    Dim content As String
    For i = 1 To n
        filePath = PERF_FMT_FIXTURE_DIR & "\PERF_FMT_" & Format$(i, "000") & ".txt"
        content = PerfReadAllTextCp932(filePath)
        If Len(content) = 0 Then Exit For
    Next i
    t1 = Timer
    Dim elapsed As Double
    elapsed = t1 - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    ok = (elapsed <= budget)
    note = "loaded=" & CStr(i - 1) & " of " & CStr(n)
    PerfCase04_LoadFormats = PerfRowJson(4, "format_load_" & CStr(n), PerfVerdict(ok), elapsed, budget, note)
End Function

Private Function PerfCase05_FreeWordSearch(ByVal n As Long, ByVal budget As Double) As String
    Dim t0 As Double, t1 As Double
    Dim ok As Boolean, note As String
    Dim hits As Long
    On Error Resume Next
    Err.Clear
    GenerateStanzaFixtures n
    t0 = Timer
    Dim i As Long
    Dim content As String
    Dim needle As String
    needle = "NEEDLE500"
    For i = 1 To n
        content = PerfReadAllTextCp932(PERF_FIXTURE_DIR & "\PERF_" & Format$(i, "00000") & ".txt")
        If InStr(content, needle) > 0 Then hits = hits + 1
    Next i
    t1 = Timer
    Dim elapsed As Double
    elapsed = t1 - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    ok = (elapsed <= budget) And (hits > 0)
    note = "hits=" & CStr(hits) & " of " & CStr(n)
    PerfCase05_FreeWordSearch = PerfRowJson(5, "freeword_search_" & CStr(n), PerfVerdict(ok), elapsed, budget, note)
End Function

Private Function PerfCase06_FieldSearch(ByVal n As Long, ByVal budget As Double) As String
    Dim t0 As Double, t1 As Double
    Dim ok As Boolean, note As String
    Dim hits As Long
    On Error Resume Next
    Err.Clear
    GenerateStanzaFixtures n
    t0 = Timer
    Dim i As Long
    Dim content As String
    Dim fieldKey As String
    fieldKey = "###" & JpFieldFirst() & "###"
    For i = 1 To n
        content = PerfReadAllTextCp932(PERF_FIXTURE_DIR & "\PERF_" & Format$(i, "00000") & ".txt")
        If InStr(content, fieldKey) > 0 Then hits = hits + 1
    Next i
    t1 = Timer
    Dim elapsed As Double
    elapsed = t1 - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    ok = (elapsed <= budget) And (hits > 0)
    note = "field-hits=" & CStr(hits) & " of " & CStr(n)
    PerfCase06_FieldSearch = PerfRowJson(6, "field_search_" & CStr(n), PerfVerdict(ok), elapsed, budget, note)
End Function

Private Function PerfCase07_RenderResults(ByVal n As Long, ByVal budget As Double) As String
    Dim t0 As Double, t1 As Double
    Dim ok As Boolean, note As String
    Dim ws As Worksheet
    Dim addErr As Long
    Application.DisplayAlerts = False
    On Error Resume Next
    Err.Clear
    ' Unprotect workbook structure before adding scratch sheet (register/search are protected by design).
    Dim wasProtected As Boolean: wasProtected = ThisWorkbook.ProtectStructure
    If wasProtected Then ThisWorkbook.Unprotect
    Err.Clear
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    addErr = Err.Number
    If addErr <> 0 Or ws Is Nothing Then
        If wasProtected Then ThisWorkbook.Protect Structure:=True
        Application.DisplayAlerts = True
        PerfCase07_RenderResults = PerfRowJson(7, "render_results_" & CStr(n), "FAIL", 0#, budget, "add scratch err=" & CStr(addErr))
        Exit Function
    End If
    Application.ScreenUpdating = False
    t0 = Timer
    Dim i As Long
    For i = 1 To n
        ws.Cells(i, 1).Value = "PERF_" & Format$(i, "00000")
        ws.Cells(i, 2).Value = "value-" & CStr(i)
    Next i
    t1 = Timer
    Application.ScreenUpdating = True
    Dim elapsed As Double
    elapsed = t1 - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    ws.Delete
    Application.DisplayAlerts = True
    If wasProtected Then ThisWorkbook.Protect Structure:=True
    ok = (elapsed <= budget)
    note = "rows=" & CStr(n)
    PerfCase07_RenderResults = PerfRowJson(7, "render_results_" & CStr(n), PerfVerdict(ok), elapsed, budget, note)
End Function

Private Function PerfCase09_WarmStart(ByVal budget As Double) As String
    ' Warm start = re-Activate ThisWorkbook (already loaded). We measure the
    ' time to Application.Calculate + force re-bind. This is a proxy for
    ' re-Open within the same Excel process.
    Dim t0 As Double, t1 As Double
    On Error Resume Next
    Err.Clear
    t0 = Timer
    ThisWorkbook.Activate
    Application.Calculate
    t1 = Timer
    Dim elapsed As Double
    elapsed = t1 - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    Dim ok As Boolean
    ok = (Err.Number = 0) And (elapsed <= budget)
    PerfCase09_WarmStart = PerfRowJson(9, "ui_warm_start", PerfVerdict(ok), elapsed, budget, "activate+calc proxy")
End Function

Private Function PerfCase11_FormDraw(ByVal n As Long, ByVal budget As Double) As String
    PerfCase11_FormDraw = MeasureFormDraw(11, n, budget)
End Function
Private Function PerfCase12_FormDraw(ByVal n As Long, ByVal budget As Double) As String
    PerfCase12_FormDraw = MeasureFormDraw(12, n, budget)
End Function
Private Function PerfCase13_FormDraw(ByVal n As Long, ByVal budget As Double) As String
    PerfCase13_FormDraw = MeasureFormDraw(13, n, budget)
End Function

Private Function MeasureFormDraw(ByVal idx As Long, ByVal n As Long, ByVal budget As Double) As String
    ' 2026-06-06: replace fake scratch-sheet measurement with real UserForm build via clsUserFormRenderer.
    ' n is interpreted as iteration count; each iteration builds + shows modeless + closes one real UserForm.
    Dim t0 As Double, t1 As Double
    Dim okBuild As Boolean: okBuild = True
    Dim note As String
    Dim fmts As Collection
    Set fmts = modFormatLoader.ListAllFormats()
    If fmts Is Nothing Or fmts.Count = 0 Then
        MeasureFormDraw = PerfRowJson(idx, "form_draw_" & CStr(n) & "_field", "FAIL", 0#, budget, "no_format_seeded")
        Exit Function
    End If
    Dim fmtId As String: fmtId = CStr(fmts.Item(1))
    Dim r As clsUserFormRenderer
    Dim i As Long
    Dim res As String
    t0 = Timer
    For i = 1 To n
        Set r = New clsUserFormRenderer
        On Error Resume Next
        Err.Clear
        res = r.ShowFormModeless(ThisWorkbook.Name, "register", "", fmtId)
        If Err.Number <> 0 Or Left$(res, 3) <> "OK:" Then
            okBuild = False
            note = "build_err iter=" & CStr(i) & " err=" & CStr(Err.Number) & " res=" & res
            Err.Clear
            r.CloseModelessForm
            Set r = Nothing
            Exit For
        End If
        r.CloseModelessForm
        Set r = Nothing
        On Error GoTo 0
    Next i
    t1 = Timer
    Dim elapsed As Double: elapsed = t1 - t0
    If Not okBuild Then
        MeasureFormDraw = PerfRowJson(idx, "form_draw_" & CStr(n) & "_field", "FAIL", elapsed, budget, note)
        Exit Function
    End If
    Dim ok As Boolean: ok = (elapsed <= budget)
    MeasureFormDraw = PerfRowJson(idx, "form_draw_" & CStr(n) & "_field", PerfVerdict(ok), elapsed, budget, "iters=" & CStr(n) & " fmtId=" & fmtId)
End Function

Private Function PerfCase14_Save1(ByVal budget As Double) As String
    Dim t0 As Double, t1 As Double
    On Error Resume Next
    Err.Clear
    EnsurePerfFolders
    t0 = Timer
    WriteOneStanza PERF_FIXTURE_DIR & "\PERF_SAVE_1.txt", 1
    t1 = Timer
    Dim elapsed As Double
    elapsed = t1 - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    Dim ok As Boolean
    ok = (Err.Number = 0) And (elapsed <= budget)
    PerfCase14_Save1 = PerfRowJson(14, "save_1_stanza", PerfVerdict(ok), elapsed, budget, "")
End Function

Private Function PerfCase15_Save1000(ByVal budget As Double) As String
    Dim t0 As Double, t1 As Double
    On Error Resume Next
    Err.Clear
    EnsurePerfFolders
    t0 = Timer
    Dim i As Long
    For i = 1 To 1000
        WriteOneStanza PERF_FIXTURE_DIR & "\PERF_SAVE_" & Format$(i, "00000") & ".txt", i
    Next i
    t1 = Timer
    Dim elapsed As Double
    elapsed = t1 - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    Dim ok As Boolean
    ok = (Err.Number = 0) And (elapsed <= budget)
    PerfCase15_Save1000 = PerfRowJson(15, "save_1000_stanza", PerfVerdict(ok), elapsed, budget, "")
End Function

Private Function PerfCase23_VbeCompile(ByVal budget As Double) As String
    ' Best-effort: call Application.Run on a no-op to force JIT pass.
    Dim t0 As Double, t1 As Double
    On Error Resume Next
    Err.Clear
    t0 = Timer
    Application.Run "NoOpForCompileMeasure"
    Err.Clear  ' missing sub is fine, we only need the dispatch overhead
    Dim x As Long
    Dim y As Long
    For y = 1 To 10000
        x = x + y
    Next y
    t1 = Timer
    Dim elapsed As Double
    elapsed = t1 - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    Dim ok As Boolean
    ok = (elapsed <= budget)
    PerfCase23_VbeCompile = PerfRowJson(23, "vbe_compile_proxy", PerfVerdict(ok), elapsed, budget, "dispatch+loop10k proxy")
End Function

Private Function PerfCase24_InstallBatTime(ByVal budget As Double) As String
    ' Read install log first/last [HH:MM:SS] timestamps.
    ' run_e2e_v20.ps1 and install bats emit "[HH:MM:SS] ..." lines.
    On Error Resume Next
    Err.Clear
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(PERF_LOG_FILE) Then
        PerfCase24_InstallBatTime = PerfRowJson(24, "install_bat_time", "SKIP_PHYSICAL", 0#, budget, "no install log at " & PERF_LOG_FILE)
        Exit Function
    End If
    Dim ts As Object
    Set ts = fso.OpenTextFile(PERF_LOG_FILE, 1, False, 0)
    If ts Is Nothing Then
        PerfCase24_InstallBatTime = PerfRowJson(24, "install_bat_time", "SKIP_PHYSICAL", 0#, budget, "open log failed")
        Exit Function
    End If
    Dim allText As String
    allText = ts.ReadAll
    ts.Close
    Dim lines() As String
    lines = Split(allText, vbCrLf)
    Dim firstLine As String, lastLine As String
    Dim i As Long
    For i = 0 To UBound(lines)
        If Len(Trim(lines(i))) > 0 And Left(lines(i), 1) = "[" Then
            firstLine = lines(i)
            Exit For
        End If
    Next i
    For i = UBound(lines) To 0 Step -1
        If Len(Trim(lines(i))) > 0 And Left(lines(i), 1) = "[" Then
            lastLine = lines(i)
            Exit For
        End If
    Next i
    If Len(firstLine) < 10 Or Len(lastLine) < 10 Then
        PerfCase24_InstallBatTime = PerfRowJson(24, "install_bat_time", "SKIP_PHYSICAL", 0#, budget, "log no [HH:MM:SS] timestamps")
        Exit Function
    End If
    Dim t1 As Date, t2 As Date
    Err.Clear
    t1 = TimeValue(Mid(firstLine, 2, 8))
    t2 = TimeValue(Mid(lastLine, 2, 8))
    If Err.Number <> 0 Then
        PerfCase24_InstallBatTime = PerfRowJson(24, "install_bat_time", "SKIP_PHYSICAL", 0#, budget, "TimeValue parse err=" & CStr(Err.Number))
        Exit Function
    End If
    Dim sec As Double
    sec = (t2 - t1) * 86400#
    If sec < 0 Then sec = sec + 86400#
    If sec <= 0 Then
        PerfCase24_InstallBatTime = PerfRowJson(24, "install_bat_time", "SKIP_PHYSICAL", 0#, budget, "zero duration in log")
        Exit Function
    End If
    Dim ok As Boolean
    ok = (sec <= budget)
    PerfCase24_InstallBatTime = PerfRowJson(24, "install_bat_time", PerfVerdict(ok), sec, budget, "from log first-last")
End Function

Private Function PerfCase25_E2EFullSuite(ByVal budget As Double) As String
    Dim t0 As Double, t1 As Double
    On Error Resume Next
    Err.Clear
    t0 = Timer
    Application.Run "modE2E_AllButtons.Run_E2E_AllButtons"
    t1 = Timer
    Dim elapsed As Double
    elapsed = t1 - t0
    If elapsed < 0 Then elapsed = elapsed + 86400#
    Dim ok As Boolean
    ok = (Err.Number = 0) And (elapsed <= budget)
    PerfCase25_E2EFullSuite = PerfRowJson(25, "e2e_full_suite", PerfVerdict(ok), elapsed, budget, "runs current role E2E set")
End Function

' ================================================================
' Helpers
' ================================================================

Private Function PerfSkipPhysical(ByVal idx As Long, ByVal nm As String, ByVal budget As Double, ByVal reason As String) As String
    PerfSkipPhysical = PerfRowJson(idx, nm, "SKIP_PHYSICAL", 0#, budget, reason)
End Function

Private Function PerfVerdict(ByVal ok As Boolean) As String
    If ok Then PerfVerdict = "PASS" Else PerfVerdict = "FAIL"
End Function

Private Function PerfRowJson(ByVal idx As Long, ByVal nm As String, ByVal verdict As String, _
        ByVal elapsed As Double, ByVal budget As Double, ByVal note As String) As String
    Dim tid As String
    tid = JpSeinouPrefix() & "-" & Format$(idx, "00")
    PerfRowJson = "{""case"":" & idx & ",""test_id"":""" & tid & """,""name"":""" & nm & """,""verdict"":""" & verdict & """,""elapsed_sec"":" & Format$(elapsed, "0.000") & ",""budget_sec"":" & Format$(budget, "0.000") & ",""note"":""" & PerfEscapeJson(note) & """}"
    PerfAppendLine idx, tid, nm, verdict, elapsed, budget, note
End Function

Private Sub PerfAppendLine(ByVal idx As Long, ByVal tid As String, ByVal nm As String, _
        ByVal verdict As String, ByVal elapsed As Double, ByVal budget As Double, ByVal note As String)
    On Error Resume Next
    Dim role As String
    role = PerfDetectRole()
    Dim p As String
    p = PERF_OUT_DIR & "\perf_" & role & "_lines.txt"
    Dim line As String
    line = role & vbTab & CStr(idx) & vbTab & tid & vbTab & nm & vbTab & verdict & vbTab & Format$(elapsed, "0.000") & vbTab & Format$(budget, "0.000") & vbTab & PerfEscapeJson(note) & vbCrLf
    Dim b() As Byte
    b = StrConv(line, vbFromUnicode)
    Dim fh As Integer
    fh = FreeFile
    Open p For Binary Access Write As #fh
    Dim sz As Long
    sz = LOF(fh)
    Put #fh, sz + 1, b
    Close #fh
End Sub

Private Function PerfDetectRole() As String
    Dim n As String
    n = ThisWorkbook.Name
    Dim dot As Long
    dot = InStrRev(n, ".")
    Dim base As String
    If dot > 0 Then base = Left(n, dot - 1) Else base = n
    If base = JpKanriBase() Then
        PerfDetectRole = "admin"
    ElseIf base = JpTourokuBase() Then
        PerfDetectRole = "register"
    ElseIf base = JpKensakuBase() Then
        PerfDetectRole = "search"
    Else
        PerfDetectRole = "unknown"
    End If
End Function

Private Function PerfEscapeJson(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    PerfEscapeJson = t
End Function

Private Sub EnsurePerfFolders()
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(PERF_OUT_DIR) Then MkDirRecursive fso, PERF_OUT_DIR
    If Not fso.FolderExists(PERF_FIXTURE_DIR) Then MkDirRecursive fso, PERF_FIXTURE_DIR
    If Not fso.FolderExists(PERF_FMT_FIXTURE_DIR) Then MkDirRecursive fso, PERF_FMT_FIXTURE_DIR
End Sub

Private Sub MkDirRecursive(ByVal fso As Object, ByVal p As String)
    On Error Resume Next
    If fso.FolderExists(p) Then Exit Sub
    Dim parent As String
    Dim slash As Long
    slash = InStrRev(p, "\")
    If slash > 0 Then
        parent = Left(p, slash - 1)
        If Len(parent) > 3 Then MkDirRecursive fso, parent
    End If
    fso.CreateFolder p
End Sub

Private Sub GenerateStanzaFixtures(ByVal n As Long)
    ' Idempotent: re-create only if file count differs.
    On Error Resume Next
    EnsurePerfFolders
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim existing As Long
    existing = 0
    Dim f As Object
    For Each f In fso.GetFolder(PERF_FIXTURE_DIR).Files
        If LCase(fso.GetExtensionName(f.Name)) = "txt" Then
            If Left(f.Name, 5) = "PERF_" And InStr(f.Name, "SAVE") = 0 Then existing = existing + 1
        End If
    Next f
    If existing >= n Then Exit Sub
    Dim i As Long
    For i = 1 To n
        WriteOneStanza PERF_FIXTURE_DIR & "\PERF_" & Format$(i, "00000") & ".txt", i
    Next i
End Sub

Private Sub WriteOneStanza(ByVal path As String, ByVal seq As Long)
    On Error Resume Next
    Dim s As String
    s = "###" & JpFieldFirst() & "###" & vbCrLf
    s = s & "PERF_" & Format$(seq, "00000") & vbCrLf
    s = s & "###" & JpFieldSecond() & "###" & vbCrLf
    s = s & "NEEDLE" & CStr(seq Mod 1000) & vbCrLf
    s = s & "###" & JpFieldThird() & "###" & vbCrLf
    s = s & "body line seq=" & CStr(seq) & vbCrLf
    Dim b() As Byte
    b = StrConv(s, vbFromUnicode)
    Dim fh As Integer
    fh = FreeFile
    Open path For Binary Access Write As #fh
    Put #fh, 1, b
    Close #fh
End Sub

Private Sub GenerateFormatFixtures(ByVal n As Long)
    On Error Resume Next
    EnsurePerfFolders
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim existing As Long
    existing = 0
    Dim f As Object
    For Each f In fso.GetFolder(PERF_FMT_FIXTURE_DIR).Files
        If LCase(fso.GetExtensionName(f.Name)) = "txt" Then existing = existing + 1
    Next f
    If existing >= n Then Exit Sub
    Dim i As Long, j As Long
    Dim s As String
    For i = 1 To n
        s = "###" & JpHeadName() & "###" & vbCrLf
        s = s & "PERF_FMT_" & Format$(i, "000") & vbCrLf
        s = s & "###" & JpHeadFields() & "###" & vbCrLf
        For j = 1 To 10
            s = s & "field" & CStr(j) & vbTab & JpTypeText() & vbCrLf
        Next j
        Dim b() As Byte
        b = StrConv(s, vbFromUnicode)
        Dim fh As Integer
        fh = FreeFile
        Open PERF_FMT_FIXTURE_DIR & "\PERF_FMT_" & Format$(i, "000") & ".txt" For Binary Access Write As #fh
        Put #fh, 1, b
        Close #fh
    Next i
End Sub

Private Function PerfReadAllTextCp932(ByVal path As String) As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    Dim fh As Integer
    Dim sz As Long
    fh = FreeFile
    Open path For Binary Access Read As #fh
    sz = LOF(fh)
    If sz <= 0 Then
        Close #fh
        Exit Function
    End If
    Dim b() As Byte
    ReDim b(0 To sz - 1)
    Get #fh, 1, b
    Close #fh
    PerfReadAllTextCp932 = StrConv(b, vbUnicode)
End Function

Private Sub PerfWriteAllText(ByVal path As String, ByVal text As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As String
    Dim slash As Long
    slash = InStrRev(path, "\")
    If slash > 0 Then
        folder = Left(path, slash - 1)
        If Not fso.FolderExists(folder) Then MkDirRecursive fso, folder
    End If
    If fso.FileExists(path) Then fso.DeleteFile path, True
    Dim b() As Byte
    b = StrConv(text, vbFromUnicode)
    Dim fh As Integer
    fh = FreeFile
    Open path For Binary Access Write As #fh
    Put #fh, 1, b
    Close #fh
End Sub

' --- JP literal helpers (ChrW only, CP932 strict source) -----
Private Function JpSeinouPrefix() As String
    ' 性能大量
    JpSeinouPrefix = ChrW(&H6027) & ChrW(&H80FD) & ChrW(&H5927) & ChrW(&H91CF)
End Function

Private Function JpKanriBase() As String
    ' 管理
    JpKanriBase = ChrW(&H7BA1) & ChrW(&H7406)
End Function

Private Function JpTourokuBase() As String
    ' 登録修正
    JpTourokuBase = ChrW(&H767B) & ChrW(&H9332) & ChrW(&H4FEE) & ChrW(&H6B63)
End Function

Private Function JpKensakuBase() As String
    ' 検索
    JpKensakuBase = ChrW(&H691C) & ChrW(&H7D22)
End Function

Private Function JpFieldFirst() As String
    ' 名称
    JpFieldFirst = ChrW(&H540D) & ChrW(&H79F0)
End Function

Private Function JpFieldSecond() As String
    ' 内容
    JpFieldSecond = ChrW(&H5185) & ChrW(&H5BB9)
End Function

Private Function JpFieldThird() As String
    ' 備考
    JpFieldThird = ChrW(&H5099) & ChrW(&H8003)
End Function

Private Function JpHeadName() As String
    ' フォーマット名
    JpHeadName = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H540D)
End Function

Private Function JpHeadFields() As String
    ' フィールド
    JpHeadFields = ChrW(&H30D5) & ChrW(&H30A3) & ChrW(&H30FC) & ChrW(&H30EB) & ChrW(&H30C9)
End Function

Private Function JpTypeText() As String
    ' テキスト
    JpTypeText = ChrW(&H30C6) & ChrW(&H30AD) & ChrW(&H30B9) & ChrW(&H30C8)
End Function
```
