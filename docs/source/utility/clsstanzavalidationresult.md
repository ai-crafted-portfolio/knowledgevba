---
title: ClsStanzaValidationResult.cls
---

# ClsStanzaValidationResult.cls

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | スタンザ検証結果を集約する値オブジェクト |
| 行数 | 106 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`ClsStanzaValidationResult.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
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
