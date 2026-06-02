---
title: ClsStanzaSection.cls
description: ClsStanzaSection.cls のソースコード（コピペ用）
---

# ClsStanzaSection.cls

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `ClsStanzaSection.cls`
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
Attribute VB_Name = "ClsStanzaSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' クラス: ClsStanzaSection（Phase 2 ユーティリティ層）
' 概要:   スタンザ 1 セクションを表す（セクション名 + key→value 辞書 + 元ファイル行番号）
' 依存先: なし
' 関連 schema: v2_stanza_parser_spec.md §1.2
' 使い方:
'   Dim sec As New ClsStanzaSection
'   sec.Init "SHEET", 1
'   sec.SetValue "SheetName", "M-05"
'   Debug.Print sec.GetValue("SheetName")
' ================================================================
Option Explicit

Private m_sectionName As String
Private m_lineNumber As Long
Private m_keyValues As Object  ' Scripting.Dictionary

' ================================================================
' 関数名: Init
' 概要:   インスタンス初期化（セクション名 + 行番号）
' 引数:   ByVal sectionName As String   - セクション名（角括弧内、例 "SHEET"）
'         ByVal lineNumber As Long      - 元ファイルの開始行番号
' ================================================================
Public Sub Init(ByVal sectionName As String, ByVal lineNumber As Long)
    m_sectionName = sectionName
    m_lineNumber = lineNumber
    Set m_keyValues = CreateObject("Scripting.Dictionary")
End Sub

' セクション名を取得（読み取り専用 Property）
Public Property Get SectionName() As String
    SectionName = m_sectionName
End Property

' 行番号を取得
Public Property Get LineNumber() As Long
    LineNumber = m_lineNumber
End Property

' KeyValues 辞書を取得（参照、外部からの直接操作は非推奨）
Public Property Get KeyValues() As Object
    Set KeyValues = m_keyValues
End Property

' ================================================================
' 関数名: SetValue
' 概要:   key=value を辞書に追加（既存 key は上書き）
' 引数:   ByVal key As String   - key 名
'         ByVal value As String - 値
' ================================================================
Public Sub SetValue(ByVal key As String, ByVal value As String)
    If m_keyValues Is Nothing Then
        Set m_keyValues = CreateObject("Scripting.Dictionary")
    End If
    m_keyValues(key) = value
End Sub

' ================================================================
' 関数名: GetValue
' 概要:   key に対応する値を取得（未定義時は空文字）
' 引数:   ByVal key As String - key 名
' 戻り値: String - 値（未定義時 vbNullString）
' ================================================================
Public Function GetValue(ByVal key As String) As String
    If m_keyValues Is Nothing Then
        GetValue = vbNullString
    ElseIf m_keyValues.Exists(key) Then
        GetValue = CStr(m_keyValues(key))
    Else
        GetValue = vbNullString
    End If
End Function

' ================================================================
' 関数名: HasKey
' 概要:   key が定義されているか判定
' ================================================================
Public Function HasKey(ByVal key As String) As Boolean
    If m_keyValues Is Nothing Then
        HasKey = False
    Else
        HasKey = m_keyValues.Exists(key)
    End If
End Function

' ================================================================
' 関数名: KeyCount
' 概要:   key 件数を取得
' ================================================================
Public Function KeyCount() As Long
    If m_keyValues Is Nothing Then
        KeyCount = 0
    Else
        KeyCount = m_keyValues.Count
    End If
End Function
```
