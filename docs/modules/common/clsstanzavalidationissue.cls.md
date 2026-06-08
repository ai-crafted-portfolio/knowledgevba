---
title: ClsStanzaValidationIssue.cls
description: ClsStanzaValidationIssue.cls 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# ClsStanzaValidationIssue.cls

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 繧ｯ繝ｩ繧ｹ繝｢繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`ClsStanzaValidationIssue.cls`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0356] ClsStanzaValidationIssue.Init ENTER"  ' [ADR-0100]
    Me.Severity = severity_
    Me.LogID = logId_
    Me.FilePath = filePath_
    Me.LineNumber = lineNumber_
    Me.SectionName = sectionName_
    Me.KeyName = keyName_
    Me.Message = message_
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0357] ClsStanzaValidationIssue.Init EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' 関数名: Format
' 概要:   1 行ログ形式に整形（[Severity] LogID@File:Line Section.Key - Message）
' 戻り値: String
' ================================================================
Public Function Format() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0358] ClsStanzaValidationIssue.Format ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0359] ClsStanzaValidationIssue.Format EXIT-OK"  ' [ADR-0100]
End Function
```