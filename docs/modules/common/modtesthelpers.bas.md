---
title: modTestHelpers.bas
description: modTestHelpers.bas 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# modTestHelpers.bas

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 讓呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`modTestHelpers.bas`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
```vb
Attribute VB_Name = "modTestHelpers"
' ================================================================
' Module:  modTestHelpers (S2-CARRY-11, Sprint 2)
' Scope:   Shared test harness helpers (Save validation, fixture I/O,
'          dict assertions). Aggregates 3 patterns duplicated across
'          modKnowledgeFileIOTest / modFormatLoaderTest /
'          modUILoaderTest / modConfigLoaderTest:
'            1) test data dir ensure + clean
'            2) Save return-code assertion (OK=0 / Rejected!=0)
'            3) fixture seed + on-disk file presence assertions
' Version: v2.1 (2026-05-19 initial)
' Layer:   L0 (test helper, ASCII only per ADR-0062 P2)
' Notes:
'   - Existing tests are NOT modified (constraint 7); new helpers are
'     opt-in for new harnesses and Sprint 3+ tests.
'   - Counter is module-private; new harnesses Call ResetCounters at
'     start, then BumpPass / BumpFail / GetPassCount / GetFailCount.
'   - All Sub/Function names ASCII; helper strings ASCII only.
' ================================================================
Option Explicit

Private m_passCount As Long
Private m_failCount As Long

' ----------------------------------------------------------------
' Counter helpers
' ----------------------------------------------------------------

Public Sub ResetCounters()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1433] modTestHelpers.ResetCounters ENTER"  ' [ADR-0100]
    m_passCount = 0
    m_failCount = 0
End Sub

Public Function GetPassCount() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1434] modTestHelpers.GetPassCount ENTER"  ' [ADR-0100]
    GetPassCount = m_passCount
End Function

Public Function GetFailCount() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1435] modTestHelpers.GetFailCount ENTER"  ' [ADR-0100]
    GetFailCount = m_failCount
End Function

Public Sub BumpPass(ByVal testName As String, ByVal subName As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1436] modTestHelpers.BumpPass ENTER"  ' [ADR-0100]
    m_passCount = m_passCount + 1
    Debug.Print "  [PASS] " & testName & " - " & subName
End Sub

Public Sub BumpFail(ByVal testName As String, ByVal subName As String, ByVal msg As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1437] modTestHelpers.BumpFail ENTER"  ' [ADR-0100]
    m_failCount = m_failCount + 1
    Debug.Print "  [FAIL] " & testName & " - " & subName & " " & msg
End Sub

' ----------------------------------------------------------------
' Fixture directory helpers (CleanFixtureDir / SeedFixture / DeleteIfExists)
' ----------------------------------------------------------------

Public Sub EnsureDir(ByVal path As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1438] modTestHelpers.EnsureDir ENTER"  ' [ADR-0100]
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
End Sub

Public Sub CleanFixtureDir(ByVal path As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1439] modTestHelpers.CleanFixtureDir ENTER"  ' [ADR-0100]
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then Exit Sub
    Dim folder As Object
    Set folder = fso.GetFolder(path)
    Dim f As Object
    On Error Resume Next
    For Each f In folder.Files
        fso.DeleteFile f.Path
    Next f
    On Error GoTo 0
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1440] modTestHelpers.CleanFixtureDir EXIT-OK"  ' [ADR-0100]
End Sub

Public Sub DeleteIfExists(ByVal path As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1441] modTestHelpers.DeleteIfExists ENTER"  ' [ADR-0100]
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If fso.FileExists(path) Then fso.DeleteFile path
    On Error GoTo 0
End Sub

Public Function SeedFixture(ByVal dirPath As String, ByVal fileName As String, _
                            ByVal payload As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1442] modTestHelpers.SeedFixture ENTER"  ' [ADR-0100]
    EnsureDir dirPath
    Dim full As String
    full = dirPath & fileName
    Dim h As Integer
    h = FreeFile
    Open full For Output As #h
    Print #h, payload
    Close #h
    SeedFixture = full
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1443] modTestHelpers.SeedFixture EXIT-OK"  ' [ADR-0100]
End Function

Public Function FixtureExists(ByVal path As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1444] modTestHelpers.FixtureExists ENTER"  ' [ADR-0100]
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FixtureExists = fso.FileExists(path)
End Function

' ----------------------------------------------------------------
' Save validation assertions (Save returns Long 0=OK / !=0 rejected)
' ----------------------------------------------------------------

Public Sub AssertSaveOK(ByVal testName As String, ByVal subName As String, _
                        ByVal ret As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1445] modTestHelpers.AssertSaveOK ENTER"  ' [ADR-0100]
    If ret = 0 Then
        BumpPass testName, subName
    Else
        BumpFail testName, subName, "expected Save OK (0) got " & ret
    End If
End Sub

Public Sub AssertSaveRejected(ByVal testName As String, ByVal subName As String, _
                              ByVal ret As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1446] modTestHelpers.AssertSaveRejected ENTER"  ' [ADR-0100]
    If ret <> 0 Then
        BumpPass testName, subName
    Else
        BumpFail testName, subName, "expected Save rejected (<>0) got 0"
    End If
End Sub

Public Sub AssertSaveRejectedWithCode(ByVal testName As String, ByVal subName As String, _
                                      ByVal expectedCode As Long, ByVal ret As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1447] modTestHelpers.AssertSaveRejectedWithCode ENTER"  ' [ADR-0100]
    If ret = expectedCode Then
        BumpPass testName, subName
    Else
        BumpFail testName, subName, "expected code " & expectedCode & " got " & ret
    End If
End Sub

Public Sub AssertFileExists(ByVal testName As String, ByVal subName As String, _
                            ByVal path As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1448] modTestHelpers.AssertFileExists ENTER"  ' [ADR-0100]
    If FixtureExists(path) Then
        BumpPass testName, subName
    Else
        BumpFail testName, subName, "file not found: " & path
    End If
End Sub

Public Sub AssertFileAbsent(ByVal testName As String, ByVal subName As String, _
                            ByVal path As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1449] modTestHelpers.AssertFileAbsent ENTER"  ' [ADR-0100]
    If Not FixtureExists(path) Then
        BumpPass testName, subName
    Else
        BumpFail testName, subName, "file still present: " & path
    End If
End Sub

' ----------------------------------------------------------------
' Generic equality assertions (mirror the duplicated AssertEqual* in tests)
' ----------------------------------------------------------------

Public Sub AssertEqualStr(ByVal testName As String, ByVal subName As String, _
                          ByVal expected As String, ByVal actual As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1450] modTestHelpers.AssertEqualStr ENTER"  ' [ADR-0100]
    If expected = actual Then
        BumpPass testName, subName
    Else
        BumpFail testName, subName, "expected='" & expected & "' actual='" & actual & "'"
    End If
End Sub

Public Sub AssertEqualLong(ByVal testName As String, ByVal subName As String, _
                           ByVal expected As Variant, ByVal actual As Variant)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1451] modTestHelpers.AssertEqualLong ENTER"  ' [ADR-0100]
    If expected = actual Then
        BumpPass testName, subName
    Else
        BumpFail testName, subName, "expected=" & expected & " actual=" & actual
    End If
End Sub

Public Sub AssertEqualBool(ByVal testName As String, ByVal subName As String, _
                           ByVal expected As Boolean, ByVal actual As Boolean)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1452] modTestHelpers.AssertEqualBool ENTER"  ' [ADR-0100]
    If expected = actual Then
        BumpPass testName, subName
    Else
        BumpFail testName, subName, "expected=" & expected & " actual=" & actual
    End If
End Sub
```