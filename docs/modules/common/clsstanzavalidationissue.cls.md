---
title: ClsStanzaValidationIssue.cls
description: ClsStanzaValidationIssue.cls のソースコード（コピペ用）
---

# ClsStanzaValidationIssue.cls

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `ClsStanzaValidationIssue.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "ClsStanzaValidationIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' クラス: ClsStanzaValidationIssue（Phase 2 ユーティリティ層）
' 概要:   schema 検証 issue 1 件を表す（severity / LogID / 位置情報 / メッセージ）
' 依存先: なし
' 関連 schema: v2_stanza_parser_spec.md §4
' Severity: "ERROR" / "WARN"（v2_stanza_parser_spec.md §3 準拠）
' ================================================================
Option Explicit

Public Severity As String         ' "ERROR" or "WARN"
Public LogID As String            ' "UI-SCHEMA-XXX" 等（ADR-0043 §4.1）
Public FilePath As String
Public LineNumber As Long
Public SectionName As String
Public KeyName As String
Public Message As String          ' 人間可読

' ================================================================
' 関数名: Init
' 概要:   全フィールドを一括設定
' ================================================================
Public Sub Init( _
    ByVal severity_ As String, _
    ByVal logId_ As String, _
    ByVal filePath_ As String, _
    ByVal lineNumber_ As Long, _
    ByVal sectionName_ As String, _
    ByVal keyName_ As String, _
    ByVal message_ As String _
)
    Me.Severity = severity_
    Me.LogID = logId_
    Me.FilePath = filePath_
    Me.LineNumber = lineNumber_
    Me.SectionName = sectionName_
    Me.KeyName = keyName_
    Me.Message = message_
End Sub

' ================================================================
' 関数名: Format
' 概要:   1 行ログ形式に整形（[Severity] LogID@File:Line Section.Key - Message）
' 戻り値: String
' ================================================================
Public Function Format() As String
    Dim s As String
    s = "[" & Severity & "] " & LogID
    If Len(FilePath) > 0 Then
        s = s & " @" & FilePath
        If LineNumber > 0 Then s = s & ":" & LineNumber
    End If
    If Len(SectionName) > 0 Then
        s = s & " (" & SectionName
        If Len(KeyName) > 0 Then s = s & "." & KeyName
        s = s & ")"
    End If
    s = s & " - " & Message
    Format = s
End Function
```
