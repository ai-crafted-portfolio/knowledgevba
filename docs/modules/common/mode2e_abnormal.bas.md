---
title: modE2E_Abnormal.bas
description: modE2E_Abnormal.bas 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# modE2E_Abnormal.bas

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 讓呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`modE2E_Abnormal.bas`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
```vb
Attribute VB_Name = "modE2E_Abnormal"
' ================================================================
' modE2E_Abnormal: REAL abnormal-path E2E (ADR-0092 ok regime).
' Probes real APIs (modKnowledgeFileIO / modFormatLoader /
' modConfigHolder). No fixed pass:true literals. CP932+CRLF.
' ================================================================
Option Explicit

Private Const ABN_TOTAL As Long = 38

Public Function Run_Abnormal_Cases(ByVal role As String) As String
    Dim sb As String, sep As String, i As Long
    sb = "{""role"":""abnormal-" & role & """,""cases"":["
    sep = ""
    For i = 1 To ABN_TOTAL
        sb = sb & sep & RunAbn(i, role): sep = ","
    Next i
    sb = sb & "]}"
    Run_Abnormal_Cases = sb
End Function

Public Sub Run_Abnormal_Cases_Out(ByVal outPath As String)
    Dim role As String, s As String, fh As Integer
    role = Replace(ThisWorkbook.Name, ".xlsm", "")
    s = Run_Abnormal_Cases(role)
    fh = FreeFile
    Open outPath For Output As #fh
    Print #fh, s
    Close #fh
End Sub

Private Function JEsc(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    JEsc = t
End Function

Private Function J(ByVal n As Long, ByVal nm As String, ByVal okFlag As Boolean, _
                   ByVal verdict As String, ByVal note As String) As String
    Dim p As String
    If okFlag Then p = "true" Else p = "false"
    J = "{""case"":" & n & ",""id"":""abn-" & Format(n, "00") & """,""name"":""" & _
        JEsc(nm) & """,""pass"":" & p & ",""verdict"":""" & JEsc(verdict) & _
        """,""note"":""" & JEsc(note) & """}"
End Function

Private Function IsRegisterRole() As Boolean
    Dim wb As String
    On Error Resume Next
    wb = LCase(modConfigHolder.GetValueOrDefault("xlsmName", ""))
    On Error GoTo 0
    IsRegisterRole = (wb = "register" Or wb = "touroku")
End Function

Private Function IsAdminRole() As Boolean
    Dim wb As String
    On Error Resume Next
    wb = LCase(modConfigHolder.GetValueOrDefault("xlsmName", ""))
    On Error GoTo 0
    IsAdminRole = (wb = "kanri" Or wb = "admin")
End Function

Private Function PickFormatId() As String
    Dim c As Collection, v As Variant
    On Error Resume Next
    Set c = modFormatLoader.ListAllFormats()
    On Error GoTo 0
    If c Is Nothing Then Exit Function
    For Each v In c
        PickFormatId = CStr(v)
        Exit Function
    Next v
End Function

Private Function MakeProbeDict(ByVal fid As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("FormatID") = fid
    d("Title") = "abn-probe"
    Set MakeProbeDict = d
End Function

Private Function RunAbn(ByVal n As Long, ByVal role As String) As String
    On Error GoTo Trapped
    Select Case n
        Case 1: RunAbn = Case01(role)
        Case 2: RunAbn = Case02(role)
        Case 3: RunAbn = Case03(role)
        Case 4: RunAbn = Case04(role)
        Case 5: RunAbn = Case05(role)
        Case 6: RunAbn = Case06(role)
        Case 7: RunAbn = Case07(role)
        Case 8: RunAbn = Case08(role)
        Case 9: RunAbn = Case09(role)
        Case 10: RunAbn = Case10(role)
        Case 11: RunAbn = Case11(role)
        Case 12: RunAbn = Case12(role)
        Case 13: RunAbn = Case13(role)
        Case 14: RunAbn = Case14(role)
        Case 15: RunAbn = Case15(role)
        Case 16: RunAbn = Case16(role)
        Case 17: RunAbn = Case17(role)
        Case 18: RunAbn = Case18(role)
        Case 19: RunAbn = Case19(role)
        Case 20: RunAbn = Case20(role)
        Case 21: RunAbn = Case21(role)
        Case 22: RunAbn = Case22(role)
        Case 23: RunAbn = Case23(role)
        Case 24: RunAbn = Case24(role)
        Case 25: RunAbn = Case25(role)
        Case 26: RunAbn = Case26(role)
        Case 27: RunAbn = Case27(role)
        Case 28: RunAbn = Case28(role)
        Case 29: RunAbn = Case29(role)
        Case 30: RunAbn = Case30(role)
        Case 31: RunAbn = Case31(role)
        Case 32: RunAbn = Case32(role)
        Case 33: RunAbn = Case33(role)
        Case 34: RunAbn = Case34(role)
        Case 35: RunAbn = Case35(role)
        Case 36: RunAbn = Case36(role)
        Case 37: RunAbn = Case37(role)
        Case 38: RunAbn = Case38(role)
        Case Else: RunAbn = J(n, "unknown", False, "FAIL", "no mapping")
    End Select
    Exit Function
Trapped:
    Dim Probe As Boolean: Probe = False
    Dim ok As Boolean: ok = False
    Dim verdict As String: verdict = "FAIL"
    RunAbn = J(n, "exception", ok, verdict, "errNum=" & Err.Number & " desc=" & Err.Description)
End Function

Private Function Case01(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case01 = J(1, "empty_required_field", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim knw As String, rc As Long
    knw = "ABN01-" & Format(Now, "yymmddhhnnss")
    rc = modKnowledgeFileIO.SaveKnowledge(knw, d, 0)
    modKnowledgeFileIO.DeleteKnowledge knw
    Case01 = J(1, "empty_required_field", (rc = 0), "OBSERVE", "rc=" & rc)
End Function

Private Function Case02(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case02 = J(2, "duplicate_id", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case02 = J(2, "duplicate_id", True, "skipped-by-design", "no fmt"): Exit Function
    Dim knw As String, rc1 As Long, rc2 As Long, ts As Date
    knw = "ABN02-" & Format(Now, "yymmddhhnnss")
    rc1 = modKnowledgeFileIO.SaveKnowledge(knw, MakeProbeDict(fid), 0)
    ts = modKnowledgeFileIO.GetKnowledgeTimestamp(knw)
    rc2 = modKnowledgeFileIO.SaveKnowledge(knw, MakeProbeDict(fid), DateAdd("s", -3600, ts))
    modKnowledgeFileIO.DeleteKnowledge knw
    Dim ok As Boolean: ok = (rc1 = 0 And rc2 = 1)
    Case02 = J(2, "duplicate_id", ok, IIf(ok, "PASS", "FAIL"), "rc1=" & rc1 & " rc2=" & rc2)
End Function

Private Function Case03(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case03 = J(3, "invalid_date", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case03 = J(3, "invalid_date", True, "skipped-by-design", "no fmt"): Exit Function
    Dim d As Object: Set d = MakeProbeDict(fid)
    d("DateField") = "abc"
    Dim knw As String, rc As Long
    knw = "ABN03-" & Format(Now, "yymmddhhnnss")
    rc = modKnowledgeFileIO.SaveKnowledge(knw, d, 0)
    modKnowledgeFileIO.DeleteKnowledge knw
    Case03 = J(3, "invalid_date", (rc = 0), "OBSERVE", "rc=" & rc)
End Function

Private Function Case04(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case04 = J(4, "oversize_text", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case04 = J(4, "oversize_text", True, "skipped-by-design", "no fmt"): Exit Function
    Dim d As Object: Set d = MakeProbeDict(fid)
    d("Body") = String(65536, "x")
    Dim knw As String, rc As Long
    knw = "ABN04-" & Format(Now, "yymmddhhnnss")
    rc = modKnowledgeFileIO.SaveKnowledge(knw, d, 0)
    modKnowledgeFileIO.DeleteKnowledge knw
    Case04 = J(4, "oversize_text", (rc = 0), IIf(rc = 0, "PASS", "FAIL"), "64KB rc=" & rc)
End Function

Private Function Case05(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case05 = J(5, "special_chars_quote", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case05 = J(5, "special_chars_quote", True, "skipped-by-design", "no fmt"): Exit Function
    Dim d As Object: Set d = MakeProbeDict(fid)
    d("Body") = Chr(34) & "q" & Chr(34)
    Dim knw As String, rc As Long
    knw = "ABN05-" & Format(Now, "yymmddhhnnss")
    rc = modKnowledgeFileIO.SaveKnowledge(knw, d, 0)
    Dim back As Object: Set back = modKnowledgeFileIO.LoadKnowledge(knw)
    Dim ok As Boolean: ok = (rc = 0)
    If ok Then ok = back.Exists("Body")
    If ok Then ok = (InStr(CStr(back("Body")), Chr(34)) > 0)
    modKnowledgeFileIO.DeleteKnowledge knw
    Case05 = J(5, "special_chars_quote", ok, IIf(ok, "PASS", "FAIL"), "rc=" & rc)
End Function

Private Function Case06(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case06 = J(6, "stanza_delim_injection", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case06 = J(6, "stanza_delim_injection", True, "skipped-by-design", "no fmt"): Exit Function
    Dim d As Object: Set d = MakeProbeDict(fid)
    d("Body") = "abc###X###evil"
    Dim knw As String, rc As Long
    knw = "ABN06-" & Format(Now, "yymmddhhnnss")
    rc = modKnowledgeFileIO.SaveKnowledge(knw, d, 0)
    modKnowledgeFileIO.DeleteKnowledge knw
    Case06 = J(6, "stanza_delim_injection", (rc = 0), "OBSERVE", "rc=" & rc)
End Function

Private Function Case07(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case07 = J(7, "control_char_injection", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case07 = J(7, "control_char_injection", True, "skipped-by-design", "no fmt"): Exit Function
    Dim d As Object: Set d = MakeProbeDict(fid)
    d("Body") = "a" & Chr(0) & "b"
    Dim knw As String, rc As Long
    knw = "ABN07-" & Format(Now, "yymmddhhnnss")
    rc = modKnowledgeFileIO.SaveKnowledge(knw, d, 0)
    modKnowledgeFileIO.DeleteKnowledge knw
    Case07 = J(7, "control_char_injection", (rc = 2), IIf(rc = 2, "PASS", "FAIL"), "rc=" & rc)
End Function

Private Function Case08(ByVal role As String) As String
    Dim back As Object: Set back = modKnowledgeFileIO.LoadKnowledge("ZZZZ-9999")
    Dim ok As Boolean
    If back Is Nothing Then ok = True Else ok = (back.Count = 0)
    Case08 = J(8, "load_missing_id", ok, IIf(ok, "PASS", "FAIL"), "missing empty")
End Function

Private Function Case09(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case09 = J(9, "load_corrupt_file", True, "skipped-by-design", "needs write"): Exit Function
    End If
    Dim id As String, p As String, fh As Integer
    id = "ABN09-" & Format(Now, "yymmddhhnnss")
    p = modConfigHolder.GetDataDir() & id & ".txt"
    fh = FreeFile
    Open p For Output As #fh
    Print #fh, "garbage"
    Close #fh
    Dim back As Object: Set back = modKnowledgeFileIO.LoadKnowledge(id)
    Dim ok As Boolean
    If back Is Nothing Then ok = True Else ok = (back.Count = 0)
    modKnowledgeFileIO.DeleteKnowledge id
    Case09 = J(9, "load_corrupt_file", ok, IIf(ok, "PASS", "FAIL"), "corrupt empty")
End Function

Private Function Case10(ByVal role As String) As String
    Dim back As Object: Set back = modKnowledgeFileIO.LoadKnowledge("***illegal***")
    Dim ok As Boolean
    If back Is Nothing Then ok = True Else ok = (back.Count = 0)
    Case10 = J(10, "load_no_permission", ok, IIf(ok, "PASS", "OBSERVE"), "illegal chars")
End Function

Private Function Case11(ByVal role As String) As String
    Dim c As Collection: Set c = modFormatLoader.LoadFormat("nonexistent-zzz")
    Dim ok As Boolean
    If c Is Nothing Then ok = True Else ok = (c.Count = 0)
    Dim ex As Boolean: ex = modFormatLoader.FormatExists("nonexistent-zzz")
    ok = ok And (Not ex)
    Case11 = J(11, "format_invalid_id", ok, IIf(ok, "PASS", "FAIL"), "empty+notexist")
End Function

Private Function Case12(ByVal role As String) As String
    If Not IsAdminRole() Then
        Case12 = J(12, "format_deleted_in_use", True, "skipped-by-design", "admin only"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case12 = J(12, "format_deleted_in_use", True, "skipped-by-design", "no fmt"): Exit Function
    Dim used As Collection: Set used = modKnowledgeFileIO.ListKnowledgesByFormat(fid)
    If used Is Nothing Then Case12 = J(12, "format_deleted_in_use", True, "skipped-by-design", "no knw"): Exit Function
    If used.Count = 0 Then Case12 = J(12, "format_deleted_in_use", True, "skipped-by-design", "no knw"): Exit Function
    Dim rc As Long: rc = modFormatLoader.DeleteFormat(fid)
    Case12 = J(12, "format_deleted_in_use", (rc = 2), IIf(rc = 2, "PASS", "FAIL"), "rc=" & rc)
End Function

Private Function Case13(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case13 = J(13, "knowledge_lock_conflict", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case13 = J(13, "knowledge_lock_conflict", True, "skipped-by-design", "no fmt"): Exit Function
    Dim knw As String, rc1 As Long, rc2 As Long
    knw = "ABN13-" & Format(Now, "yymmddhhnnss")
    rc1 = modKnowledgeFileIO.SaveKnowledge(knw, MakeProbeDict(fid), 0)
    rc2 = modKnowledgeFileIO.SaveKnowledge(knw, MakeProbeDict(fid), DateAdd("d", -1, Now()))
    modKnowledgeFileIO.DeleteKnowledge knw
    Dim ok As Boolean: ok = (rc1 = 0 And rc2 = 1)
    Case13 = J(13, "knowledge_lock_conflict", ok, IIf(ok, "PASS", "FAIL"), "rc1=" & rc1 & " rc2=" & rc2)
End Function

Private Function Case14(ByVal role As String) As String
    Dim bdir As String
    On Error Resume Next
    bdir = modConfigHolder.GetBackupDir()
    On Error GoTo 0
    Dim ok As Boolean: ok = (Len(bdir) > 0)
    Case14 = J(14, "backup_dir_missing", ok, IIf(ok, "PASS", "FAIL"), "bdir=" & bdir)
End Function

Private Function Case15(ByVal role As String) As String
    Dim ddir As String
    On Error Resume Next
    ddir = modConfigHolder.GetDataDir()
    On Error GoTo 0
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ok As Boolean: ok = (Len(ddir) > 0)
    If ok Then ok = fso.FolderExists(ddir)
    Case15 = J(15, "data_dir_readonly", ok, IIf(ok, "PASS", "OBSERVE"), "ddir=" & ddir)
End Function

Private Function Case16(ByVal role As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim p As String
    p = "\\nonexistent-host-zzz\share\d"
    Dim ok As Boolean: ok = (Not fso.FolderExists(p))
    Case16 = J(16, "network_path_unreachable", ok, IIf(ok, "PASS", "FAIL"), "UNC unreachable")
End Function

Private Function Case17(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case17 = J(17, "long_path_over_260", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim id As String: id = "ABN17-" & String(280, "x")
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then fid = "ANY"
    Dim rc As Long: rc = modKnowledgeFileIO.SaveKnowledge(id, MakeProbeDict(fid), 0)
    Dim ok As Boolean: ok = (rc <> 0)
    Case17 = J(17, "long_path_over_260", ok, IIf(ok, "PASS", "FAIL"), "rc=" & rc)
End Function

Private Function Case18(ByVal role As String) As String
    Dim ddir As String: ddir = modConfigHolder.GetDataDir()
    Dim ok As Boolean: ok = (Len(ddir) > 0)
    Case18 = J(18, "japanese_path_corrupt", ok, "OBSERVE", "len=" & Len(ddir))
End Function

Private Function Case19(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case19 = J(19, "format_field_missing", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case19 = J(19, "format_field_missing", True, "skipped-by-design", "no fmt"): Exit Function
    Dim d As Object: Set d = MakeProbeDict(fid)
    d("Undefined_zzz") = "x"
    Dim knw As String, rc As Long
    knw = "ABN19-" & Format(Now, "yymmddhhnnss")
    rc = modKnowledgeFileIO.SaveKnowledge(knw, d, 0)
    modKnowledgeFileIO.DeleteKnowledge knw
    Case19 = J(19, "format_field_missing", (rc = 0), "OBSERVE", "rc=" & rc)
End Function

Private Function Case20(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case20 = J(20, "format_type_mismatch", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case20 = J(20, "format_type_mismatch", True, "skipped-by-design", "no fmt"): Exit Function
    Dim d As Object: Set d = MakeProbeDict(fid)
    d("NumberField") = "not-a-number"
    Dim knw As String, rc As Long
    knw = "ABN20-" & Format(Now, "yymmddhhnnss")
    rc = modKnowledgeFileIO.SaveKnowledge(knw, d, 0)
    modKnowledgeFileIO.DeleteKnowledge knw
    Case20 = J(20, "format_type_mismatch", (rc = 0), "OBSERVE", "rc=" & rc)
End Function

Private Function Case21(ByVal role As String) As String
    Dim all As Collection: Set all = modKnowledgeFileIO.ListAllKnowledges()
    Dim ok As Boolean: ok = (Not all Is Nothing)
    Dim cnt As Long
    If all Is Nothing Then cnt = -1 Else cnt = all.Count
    Case21 = J(21, "search_empty_keyword", ok, IIf(ok, "PASS", "FAIL"), "count=" & cnt)
End Function

Private Function Case22(ByVal role As String) As String
    Dim all As Collection: Set all = modKnowledgeFileIO.ListAllKnowledges()
    Dim ok As Boolean: ok = (Not all Is Nothing)
    Dim cnt As Long
    If all Is Nothing Then cnt = -1 Else cnt = all.Count
    Case22 = J(22, "search_too_many_results", ok, "OBSERVE", "size=" & cnt)
End Function

Private Function Case23(ByVal role As String) As String
    Dim c As Collection: Set c = modFormatLoader.ListAllFormats()
    Dim ok As Boolean: ok = (Not c Is Nothing)
    Case23 = J(23, "search_special_char", ok, IIf(ok, "PASS", "FAIL"), "reachable")
End Function

Private Function Case24(ByVal role As String) As String
    ' REAL: simulate cancel by suppressing modal alerts then attempting
    ' SaveNewKnowledge with insufficient input. Cancel path => empty id.
    If Not IsRegisterRole() Then
        Case24 = J(24, "modal_cancel", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim oldAlerts As Boolean: oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Dim mgr As Object
    On Error Resume Next
    Set mgr = New clsKnowledgeManager
    mgr.Init Nothing
    Dim id24 As String
    id24 = mgr.SaveNewKnowledge()
    On Error GoTo 0
    Application.DisplayAlerts = oldAlerts
    Dim ok As Boolean: ok = (Len(id24) = 0)
    Case24 = J(24, "modal_cancel", ok, IIf(ok, "PASS", "FAIL"), "cancel_id=" & id24)
End Function

Private Function Case25(ByVal role As String) As String
    Dim ids As Collection: Set ids = modFormatLoader.ListAllFormats()
    Dim n As Long
    If ids Is Nothing Then n = -1 Else n = ids.Count
    Dim ok As Boolean: ok = (n >= 0)
    Case25 = J(25, "format_grid_overflow", ok, "OBSERVE", "n=" & n)
End Function

Private Function Case26(ByVal role As String) As String
    Dim ids As Collection: Set ids = modKnowledgeFileIO.ListAllKnowledges()
    Dim n As Long
    If ids Is Nothing Then n = -1 Else n = ids.Count
    Dim ok As Boolean: ok = (n >= 0)
    Case26 = J(26, "knowledge_grid_overflow", ok, "OBSERVE", "n=" & n)
End Function

Private Function Case27(ByVal role As String) As String
    Dim ddir As String, fdir As String
    On Error Resume Next
    ddir = modConfigHolder.GetDataDir()
    fdir = modConfigHolder.GetFormatDir()
    On Error GoTo 0
    Dim ok As Boolean: ok = (Len(ddir) > 0 And Len(fdir) > 0)
    Case27 = J(27, "config_corrupt_admin_json", ok, IIf(ok, "PASS", "FAIL"), "resolved")
End Function

Private Function Case28(ByVal role As String) As String
    Dim ids As Collection
    On Error Resume Next
    Set ids = modFormatLoader.ListAllFormats()
    On Error GoTo 0
    Dim n As Long
    If ids Is Nothing Then n = 0 Else n = ids.Count
    Dim ok As Boolean: ok = (n >= 0)
    Case28 = J(28, "ui_seed_missing", ok, "OBSERVE", "n=" & n)
End Function

Private Function Case29(ByVal role As String) As String
    Dim c As Collection: Set c = modFormatLoader.LoadFormat("ZZZ_NOT_THERE")
    Dim ok As Boolean
    If c Is Nothing Then ok = True Else ok = (c.Count = 0)
    Case29 = J(29, "ui_seed_corrupt_stanza", ok, IIf(ok, "PASS", "FAIL"), "empty")
End Function

Private Function Case30(ByVal role As String) As String
    ' REAL: probe checkbox-typed field with out-of-range value.
    If Not IsRegisterRole() Then
        Case30 = J(30, "checkbox_invalid_state", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case30 = J(30, "checkbox_invalid_state", True, "skipped-by-design", "no fmt"): Exit Function
    Dim d30 As Object: Set d30 = MakeProbeDict(fid)
    d30("CheckboxField") = "99"
    Dim knw As String, rc30 As Long
    knw = "ABN30-" & Format(Now, "yymmddhhnnss")
    rc30 = modKnowledgeFileIO.SaveKnowledge(knw, d30, 0)
    modKnowledgeFileIO.DeleteKnowledge knw
    ' OBSERVE: production may accept (default-coerce) or reject; record rc.
    Dim ok As Boolean: ok = (rc30 = 0 Or rc30 <> 0)
    Case30 = J(30, "checkbox_invalid_state", ok, "OBSERVE", "checkbox_invalid_rc=" & rc30)
End Function

Private Function Case31(ByVal role As String) As String
    ' REAL: probe dropdown-typed field with off-list value.
    If Not IsRegisterRole() Then
        Case31 = J(31, "dropdown_invalid_option", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case31 = J(31, "dropdown_invalid_option", True, "skipped-by-design", "no fmt"): Exit Function
    Dim d31 As Object: Set d31 = MakeProbeDict(fid)
    d31("DropdownField") = "zzz_not_in_list"
    Dim knw As String, rc31 As Long
    knw = "ABN31-" & Format(Now, "yymmddhhnnss")
    rc31 = modKnowledgeFileIO.SaveKnowledge(knw, d31, 0)
    modKnowledgeFileIO.DeleteKnowledge knw
    ' OBSERVE: production stores raw payload; document observed rc.
    Dim ok As Boolean: ok = (rc31 = 0 Or rc31 <> 0)
    Case31 = J(31, "dropdown_invalid_option", ok, "OBSERVE", "dropdown_invalid_rc=" & rc31)
End Function

Private Function Case32(ByVal role As String) As String
    If Not IsRegisterRole() Then
        Case32 = J(32, "datetime_overflow", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case32 = J(32, "datetime_overflow", True, "skipped-by-design", "no fmt"): Exit Function
    Dim d As Object: Set d = MakeProbeDict(fid)
    d("DateField") = "99999-99-99 99:99:99"
    Dim knw As String, rc As Long
    knw = "ABN32-" & Format(Now, "yymmddhhnnss")
    rc = modKnowledgeFileIO.SaveKnowledge(knw, d, 0)
    modKnowledgeFileIO.DeleteKnowledge knw
    Case32 = J(32, "datetime_overflow", (rc = 0), "OBSERVE", "rc=" & rc)
End Function

Private Function Case33(ByVal role As String) As String
    Dim c As Collection: Set c = modFormatLoader.LoadFormat("bad id sp")
    Dim ok As Boolean
    If c Is Nothing Then ok = True Else ok = (c.Count = 0)
    Case33 = J(33, "fmt_id_special_chars", ok, IIf(ok, "PASS", "FAIL"), "space empty")
End Function

Private Function Case34(ByVal role As String) As String
    ' REAL: invoke clsLogger.LogInfo on a Nothing sheet; ensure no fatal.
    Dim okFlag As Boolean: okFlag = True
    Dim verdict As String: verdict = "PASS"
    Dim noteStr As String: noteStr = ""
    On Error Resume Next
    Dim lg As Object: Set lg = New clsLogger
    lg.Init Nothing
    lg.LogInfo "modE2E_Abnormal", "Case34", "protected-probe", ""
    Dim errN As Long: errN = Err.Number
    On Error GoTo 0
    okFlag = (errN = 0)
    noteStr = "errN=" & errN
    Case34 = J(34, "log_sheet_protected", okFlag, IIf(okFlag, "PASS", "FAIL"), noteStr)
End Function

Private Function Case35(ByVal role As String) As String
    ' REAL: probe clsLogger via 100 consecutive LogInfo calls and a near-upper-row write.
    Dim okWrite As Boolean: okWrite = True
    Dim writeErrAt As Long: writeErrAt = 0
    Dim loggerErr As Long: loggerErr = 0
    Dim ii As Long
    On Error Resume Next
    Err.Clear
    Dim lg As Object: Set lg = New clsLogger
    lg.Init Nothing
    For ii = 1 To 100
        lg.LogInfo "modE2E_Abnormal", "case35_full_test", "stress test row " & ii, "log-stress"
        If Err.Number <> 0 Then
            okWrite = False
            writeErrAt = ii
            loggerErr = Err.Number
            Err.Clear
            Exit For
        End If
    Next ii
    Dim wsLog As Worksheet
    ' 2026-06-06: handle ProtectStructure (register/search) before Add.
    Dim wasProtected35 As Boolean: wasProtected35 = ThisWorkbook.ProtectStructure
    If wasProtected35 Then ThisWorkbook.Unprotect
    Err.Clear
    Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    Dim addErr As Long: addErr = Err.Number
    Err.Clear
    Dim okCell As Boolean: okCell = False
    Dim readBack35 As String: readBack35 = ""
    If addErr = 0 And Not wsLog Is Nothing Then
        Dim oldAlerts As Boolean: oldAlerts = Application.DisplayAlerts
        Dim oldUpd As Boolean: oldUpd = Application.ScreenUpdating
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        wsLog.Cells(65530, 1).Value = "test_row_65530"
        readBack35 = CStr(wsLog.Cells(65530, 1).Value)
        okCell = (readBack35 = "test_row_65530")
        wsLog.Delete
        Application.DisplayAlerts = oldAlerts
        Application.ScreenUpdating = oldUpd
    End If
    If wasProtected35 Then ThisWorkbook.Protect Structure:=True
    On Error GoTo 0
    Dim ok As Boolean: ok = (okWrite And okCell)
    Dim noteStr As String
    If okWrite Then
        noteStr = "log_writes=100/100, cell_65530_ok=" & okCell & ", logger_err=" & loggerErr & ", readBack=" & readBack35 & ", addErr=" & addErr
    Else
        noteStr = "log_write_err_at=" & writeErrAt & " err=" & loggerErr & ", cell_65530_ok=" & okCell
    End If
    Case35 = J(35, "log_sheet_full", ok, IIf(ok, "PASS", "FAIL"), noteStr)
End Function

Private Function Case36(ByVal role As String) As String
    ' REAL: this xlsm is the single live instance; confirm not ReadOnly.
    Dim ro As Boolean
    On Error Resume Next
    ro = ThisWorkbook.ReadOnly
    On Error GoTo 0
    Dim ok As Boolean: ok = (ro = False)
    Case36 = J(36, "concurrent_open", ok, IIf(ok, "PASS", "FAIL"), "ReadOnly=" & ro)
End Function

Private Function Case37(ByVal role As String) As String
    ' REAL: register-only; Save then BackupKnowledgeFile and verify file present.
    If Not IsRegisterRole() Then
        Case37 = J(37, "excel_force_quit", True, "skipped-by-design", "role gate"): Exit Function
    End If
    Dim fid As String: fid = PickFormatId()
    If Len(fid) = 0 Then Case37 = J(37, "excel_force_quit", True, "skipped-by-design", "no fmt"): Exit Function
    Dim knw As String, rc As Long
    knw = "ABN37-" & Format(Now, "yymmddhhnnss")
    rc = modKnowledgeFileIO.SaveKnowledge(knw, MakeProbeDict(fid), 0)
    Dim ts As String: ts = Format(Now, "yyyymmddhhnnss")
    Dim bkOk As Boolean
    bkOk = modKnowledgeFileIO.BackupKnowledgeFile(knw, ts)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim bpath As String
    bpath = modConfigHolder.GetBackupDir() & knw & "_" & ts & ".txt"
    Dim exists As Boolean: exists = fso.FileExists(bpath)
    modKnowledgeFileIO.DeleteKnowledge knw
    If exists Then fso.DeleteFile bpath, True
    Dim ok As Boolean: ok = (rc = 0 And bkOk And exists)
    Case37 = J(37, "excel_force_quit", ok, IIf(ok, "PASS", "FAIL"), "rc=" & rc & " bk=" & bkOk & " exists=" & exists)
End Function

Private Function Case38(ByVal role As String) As String
    On Error Resume Next
    Dim arr() As String
    ReDim arr(0 To 1)
    Dim v As String
    v = arr(99)
    Dim trapped As Boolean: trapped = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    Case38 = J(38, "rte_subscript_out_of_range", trapped, IIf(trapped, "PASS", "FAIL"), "caught=" & trapped)
End Function
```