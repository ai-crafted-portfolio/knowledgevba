---
title: ClsStanzaValidationResult.cls
description: ClsStanzaValidationResult.cls のソースコード（コピペ用）
---

# ClsStanzaValidationResult.cls

**配置先**: `共通モジュール (3 ブック全て)` 用の VBA モジュール  
**種類**: クラス モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\`
- ファイル名: `ClsStanzaValidationResult.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。

---

## ソースコード

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "ClsStanzaValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' クラス: ClsStanzaValidationResult（Phase 2 ユーティリティ層）
' 概要:   schema 検証結果の集約（ERROR/WARN 件数 + ClsStanzaValidationIssue のリスト）
' 依存先: ClsStanzaValidationIssue
' 関連 schema: v2_stanza_parser_spec.md §4
' ================================================================
Option Explicit

Private m_errors As Collection    ' ClsStanzaValidationIssue (Severity=ERROR)
Private m_warnings As Collection  ' ClsStanzaValidationIssue (Severity=WARN)

' ================================================================
' 関数名: Init
' 概要:   インスタンス初期化
' ================================================================
Public Sub Init()
    Set m_errors = New Collection
    Set m_warnings = New Collection
End Sub

' ================================================================
' 関数名: AddIssue
' 概要:   issue を追加（Severity に応じて Errors/Warnings に振り分け）
' 引数:   ByVal issue As ClsStanzaValidationIssue
' ================================================================
Public Sub AddIssue(ByVal issue As ClsStanzaValidationIssue)
    If m_errors Is Nothing Then Init
    If issue.Severity = "ERROR" Then
        m_errors.Add issue
    ElseIf issue.Severity = "WARN" Then
        m_warnings.Add issue
    End If
End Sub

' ----------------------------------------------------------------
' Property: IsValid（ERROR 0 件なら TRUE）
' ----------------------------------------------------------------
Public Property Get IsValid() As Boolean
    If m_errors Is Nothing Then
        IsValid = True
    Else
        IsValid = (m_errors.Count = 0)
    End If
End Property

' ----------------------------------------------------------------
' Property: ErrorCount
' ----------------------------------------------------------------
Public Property Get ErrorCount() As Long
    If m_errors Is Nothing Then
        ErrorCount = 0
    Else
        ErrorCount = m_errors.Count
    End If
End Property

' ----------------------------------------------------------------
' Property: WarningCount
' ----------------------------------------------------------------
Public Property Get WarningCount() As Long
    If m_warnings Is Nothing Then
        WarningCount = 0
    Else
        WarningCount = m_warnings.Count
    End If
End Property

' ----------------------------------------------------------------
' Property: Errors（参照）
' ----------------------------------------------------------------
Public Property Get Errors() As Collection
    If m_errors Is Nothing Then Init
    Set Errors = m_errors
End Property

' ----------------------------------------------------------------
' Property: Warnings（参照）
' ----------------------------------------------------------------
Public Property Get Warnings() As Collection
    If m_warnings Is Nothing Then Init
    Set Warnings = m_warnings
End Property

' ================================================================
' 関数名: FormatSummary
' 概要:   人間可読サマリ（"PASS (0 errors, 2 warnings)" 等）
' 戻り値: String
' ================================================================
Public Function FormatSummary() As String
    Dim status As String
    If IsValid Then
        status = "PASS"
    Else
        status = "FAIL"
    End If
    FormatSummary = status & " (" & ErrorCount & " errors, " & WarningCount & " warnings)"
End Function
```
