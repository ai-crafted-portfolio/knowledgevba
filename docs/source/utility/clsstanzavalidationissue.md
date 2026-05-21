---
title: ClsStanzaValidationIssue.cls
---

# ClsStanzaValidationIssue.cls

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | スタンザ検証の指摘 1 件を表す値オブジェクト |
| 行数 | 68 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`ClsStanzaValidationIssue.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
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
