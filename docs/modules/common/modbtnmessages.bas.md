---
title: modBtnMessages.bas
description: modBtnMessages.bas のソースコード（コピペ用）
---

# modBtnMessages.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-05 01:27 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modBtnMessages.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modBtnMessages"
Option Explicit

' ============================================================
' modBtnMessages
' Role: ボタン前提条件 NG / 実行時エラーの自然言語メッセージを集約。
' Coding rule: ADR-0094 D1 (CP932 strict + Python script for write)
' ============================================================

' Public Function GetMessage
' Args: msgId (例: "MSG-BTN-PRE-001"), args (置換用パラメータ {0} {1} ...)
' Returns: 自然言語メッセージ文字列 (placeholder 置換済)
Public Function GetMessage(ByVal msgId As String, ParamArray args() As Variant) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0899] modBtnMessages.GetMessage ENTER"  ' [ADR-0100]
    Dim tmpl As String
    Dim i As Long
    Dim ph As String
    Select Case msgId
        Case "MSG-BTN-PRE-001"
            tmpl = ChrW(&H8A2D) & ChrW(&H5B9A) & ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H672A) & ChrW(&H8AAD) & ChrW(&H8FBC) & ": config.txt " & ChrW(&H304C) & ChrW(&H8AAD) & ChrW(&H307F) & ChrW(&H8FBC) & ChrW(&H307E) & ChrW(&H308C) & ChrW(&H3066) & ChrW(&H3044) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002) & ChrW(&H30D6) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H3092) & ChrW(&H958B) & ChrW(&H304D) & ChrW(&H76F4) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
        Case "MSG-BTN-PRE-002"
            tmpl = ChrW(&H30B7) & ChrW(&H30FC) & ChrW(&H30C8) & ChrW(&H4E0D) & ChrW(&H5B58) & ChrW(&H5728) & ": " & ChrW(&H30B7) & ChrW(&H30FC) & ChrW(&H30C8) & " ""{0}"" " & ChrW(&H304C) & ChrW(&H898B) & ChrW(&H3064) & ChrW(&H304B) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002) & "Setup " & ChrW(&H30DE) & ChrW(&H30AF) & ChrW(&H30ED) & ChrW(&H3092) & ChrW(&H5B9F) & ChrW(&H884C) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H30B7) & ChrW(&H30FC) & ChrW(&H30C8) & ChrW(&H3092) & ChrW(&H518D) & ChrW(&H751F) & ChrW(&H6210) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
        Case "MSG-BTN-PRE-003"
            tmpl = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H672A) & ChrW(&H767B) & ChrW(&H9332) & ": " & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & " ID ""{0}"" " & ChrW(&H304C) & ChrW(&H5B9A) & ChrW(&H7FA9) & ChrW(&H3055) & ChrW(&H308C) & ChrW(&H3066) & ChrW(&H3044) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002) & ChrW(&H7BA1) & ChrW(&H7406) & ".xlsm " & ChrW(&H3067) & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H3092) & ChrW(&H8FFD) & ChrW(&H52A0) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
        Case "MSG-BTN-PRE-004"
            tmpl = ChrW(&H5165) & ChrW(&H529B) & ChrW(&H6B04) & ChrW(&H7A7A) & ": ""{0}"" " & ChrW(&H3092) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
        Case "MSG-BTN-PRE-005"
            tmpl = ChrW(&H30C7) & ChrW(&H30FC) & ChrW(&H30BF) & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30EB) & ChrW(&H30C0) & ChrW(&H4E0D) & ChrW(&H5B58) & ChrW(&H5728) & ": ""{0}"" " & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30EB) & ChrW(&H30C0) & ChrW(&H304C) & ChrW(&H5B58) & ChrW(&H5728) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002) & "config.txt " & ChrW(&H306E) & " data_dir " & ChrW(&H3092) & ChrW(&H78BA) & ChrW(&H8A8D) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
        Case "MSG-BTN-PRE-006"
            tmpl = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30EB) & ChrW(&H30C0) & ChrW(&H4E0D) & ChrW(&H5B58) & ChrW(&H5728) & ": ""{0}"" " & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30EB) & ChrW(&H30C0) & ChrW(&H304C) & ChrW(&H5B58) & ChrW(&H5728) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002) & "config.txt " & ChrW(&H306E) & " format_dir " & ChrW(&H3092) & ChrW(&H78BA) & ChrW(&H8A8D) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
        Case Else
            tmpl = ChrW(&H672A) & ChrW(&H5B9A) & ChrW(&H7FA9) & ChrW(&H30E1) & ChrW(&H30C3) & ChrW(&H30BB) & ChrW(&H30FC) & ChrW(&H30B8) & " ID: " & msgId
    End Select
    ' Placeholder 置換
    If Not IsMissing(args) Then
        For i = LBound(args) To UBound(args)
            ph = "{" & CStr(i) & "}"
            tmpl = Replace(tmpl, ph, CStr(args(i)))
        Next i
    End If
    GetMessage = tmpl
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0900] modBtnMessages.GetMessage EXIT-OK"  ' [ADR-0100]
End Function

' Public Function GetErrorMessage
' Args: errNum (VBA Err.Number)
' Returns: err num を自然言語に翻訳した文字列
Public Function GetErrorMessage(ByVal errNum As Long) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0901] modBtnMessages.GetErrorMessage ENTER"  ' [ADR-0100]
    Select Case errNum
        Case 9
            GetErrorMessage = ChrW(&H914D) & ChrW(&H5217) & ChrW(&H30A4) & ChrW(&H30F3) & ChrW(&H30C7) & ChrW(&H30C3) & ChrW(&H30AF) & ChrW(&H30B9) & ChrW(&H304C) & ChrW(&H7BC4) & ChrW(&H56F2) & ChrW(&H5916) & ChrW(&H3067) & ChrW(&H3059) & ChrW(&H3002) & ChrW(&H524D) & ChrW(&H63D0) & ChrW(&H6761) & ChrW(&H4EF6) & " (" & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H5B9A) & ChrW(&H7FA9) & ChrW(&H306A) & ChrW(&H3069) & ") " & ChrW(&H3092) & ChrW(&H78BA) & ChrW(&H8A8D) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
        Case 13
            GetErrorMessage = ChrW(&H578B) & ChrW(&H304C) & ChrW(&H4E00) & ChrW(&H81F4) & ChrW(&H3057) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H6B04) & ChrW(&H306E) & ChrW(&H5024) & ChrW(&H306E) & ChrW(&H578B) & ChrW(&H3092) & ChrW(&H78BA) & ChrW(&H8A8D) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
        Case 91
            GetErrorMessage = ChrW(&H524D) & ChrW(&H63D0) & ChrW(&H3068) & ChrW(&H306A) & ChrW(&H308B) & ChrW(&H8A2D) & ChrW(&H5B9A) & "/" & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30FC) & ChrW(&H30DE) & ChrW(&H30C3) & ChrW(&H30C8) & ChrW(&H5B9A) & ChrW(&H7FA9) & ChrW(&H304C) & ChrW(&H8AAD) & ChrW(&H307F) & ChrW(&H8FBC) & ChrW(&H307E) & ChrW(&H308C) & ChrW(&H3066) & ChrW(&H3044) & ChrW(&H307E) & ChrW(&H305B) & ChrW(&H3093) & ChrW(&H3002) & "Setup " & ChrW(&H30DE) & ChrW(&H30AF) & ChrW(&H30ED) & ChrW(&H3092) & ChrW(&H5B9F) & ChrW(&H884C) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
        Case 424
            GetErrorMessage = ChrW(&H30AA) & ChrW(&H30D6) & ChrW(&H30B8) & ChrW(&H30A7) & ChrW(&H30AF) & ChrW(&H30C8) & ChrW(&H304C) & ChrW(&H5FC5) & ChrW(&H8981) & ChrW(&H3067) & ChrW(&H3059) & ChrW(&H3002) & ChrW(&H30B7) & ChrW(&H30FC) & ChrW(&H30C8) & ChrW(&H307E) & ChrW(&H305F) & ChrW(&H306F) & ChrW(&H30B3) & ChrW(&H30F3) & ChrW(&H30C8) & ChrW(&H30ED) & ChrW(&H30FC) & ChrW(&H30EB) & ChrW(&H304C) & ChrW(&H898B) & ChrW(&H3064) & ChrW(&H304B) & ChrW(&H3089) & ChrW(&H306A) & ChrW(&H3044) & ChrW(&H53EF) & ChrW(&H80FD) & ChrW(&H6027) & ChrW(&H304C) & ChrW(&H3042) & ChrW(&H308A) & ChrW(&H307E) & ChrW(&H3059) & ChrW(&H3002)
        Case 1004
            GetErrorMessage = "Excel " & ChrW(&H64CD) & ChrW(&H4F5C) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC) & ChrW(&H3067) & ChrW(&H3059) & ChrW(&H3002) & ChrW(&H30B7) & ChrW(&H30FC) & ChrW(&H30C8) & ChrW(&H306E) & ChrW(&H4FDD) & ChrW(&H8B77) & ChrW(&H3084) & ChrW(&H30BB) & ChrW(&H30EB) & ChrW(&H53C2) & ChrW(&H7167) & ChrW(&H3092) & ChrW(&H78BA) & ChrW(&H8A8D) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H304F) & ChrW(&H3060) & ChrW(&H3055) & ChrW(&H3044) & ChrW(&H3002)
        Case Else
            GetErrorMessage = ChrW(&H4E88) & ChrW(&H671F) & ChrW(&H305B) & ChrW(&H306C) & ChrW(&H30A8) & ChrW(&H30E9) & ChrW(&H30FC) & " (err " & CStr(errNum) & ")"
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0902] modBtnMessages.GetErrorMessage EXIT-OK"  ' [ADR-0100]
End Function
```
