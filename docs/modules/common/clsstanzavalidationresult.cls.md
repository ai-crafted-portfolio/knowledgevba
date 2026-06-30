---
title: ClsStanzaValidationResult.cls
description: ClsStanzaValidationResult.cls のソースコード（コピペ用）
---

# ClsStanzaValidationResult.cls

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: クラスモジュール
**更新日**: 2026-06-30 14:44 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `ClsStanzaValidationResult.cls`
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0360] ClsStanzaValidationResult.Init ENTER"  ' [ADR-0100]
    Set m_errors = New Collection
    Set m_warnings = New Collection
End Sub

' ================================================================
' 関数名: AddIssue
' 概要:   issue を追加（Severity に応じて Errors/Warnings に振り分け）
' 引数:   ByVal issue As ClsStanzaValidationIssue
' ================================================================
Public Sub AddIssue(ByVal issue As ClsStanzaValidationIssue)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0361] ClsStanzaValidationResult.AddIssue ENTER"  ' [ADR-0100]
    If m_errors Is Nothing Then Init
    If issue.Severity = "ERROR" Then
        m_errors.Add issue
    ElseIf issue.Severity = "WARN" Then
        m_warnings.Add issue
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0362] ClsStanzaValidationResult.AddIssue EXIT-OK"  ' [ADR-0100]
End Sub

' ----------------------------------------------------------------
' Property: IsValid（ERROR 0 件なら TRUE）
' ----------------------------------------------------------------
Public Property Get IsValid() As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0363] ClsStanzaValidationResult.IsValid ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0364] ClsStanzaValidationResult.ErrorCount ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0365] ClsStanzaValidationResult.WarningCount ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0366] ClsStanzaValidationResult.Errors ENTER"  ' [ADR-0100]
    If m_errors Is Nothing Then Init
    Set Errors = m_errors
End Property

' ----------------------------------------------------------------
' Property: Warnings（参照）
' ----------------------------------------------------------------
Public Property Get Warnings() As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0367] ClsStanzaValidationResult.Warnings ENTER"  ' [ADR-0100]
    If m_warnings Is Nothing Then Init
    Set Warnings = m_warnings
End Property

' ================================================================
' 関数名: FormatSummary
' 概要:   人間可読サマリ（"PASS (0 errors, 2 warnings)" 等）
' 戻り値: String
' ================================================================
Public Function FormatSummary() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0368] ClsStanzaValidationResult.FormatSummary ENTER"  ' [ADR-0100]
    Dim status As String
    If IsValid Then
        status = "PASS"
    Else
        status = "FAIL"
    End If
    FormatSummary = status & " (" & ErrorCount & " errors, " & WarningCount & " warnings)"
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0369] ClsStanzaValidationResult.FormatSummary EXIT-OK"  ' [ADR-0100]
End Function
```
