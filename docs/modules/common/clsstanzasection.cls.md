---
title: ClsStanzaSection.cls
description: ClsStanzaSection.cls 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# ClsStanzaSection.cls

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 繧ｯ繝ｩ繧ｹ繝｢繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`ClsStanzaSection.cls`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0347] ClsStanzaSection.Init ENTER"  ' [ADR-0100]
    m_sectionName = sectionName
    m_lineNumber = lineNumber
    Set m_keyValues = CreateObject("Scripting.Dictionary")
End Sub

' セクション名を取得（読み取り専用 Property）
Public Property Get SectionName() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0348] ClsStanzaSection.SectionName ENTER"  ' [ADR-0100]
    SectionName = m_sectionName
End Property

' 行番号を取得
Public Property Get LineNumber() As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0349] ClsStanzaSection.LineNumber ENTER"  ' [ADR-0100]
    LineNumber = m_lineNumber
End Property

' KeyValues 辞書を取得（参照、外部からの直接操作は非推奨）
Public Property Get KeyValues() As Object
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0350] ClsStanzaSection.KeyValues ENTER"  ' [ADR-0100]
    Set KeyValues = m_keyValues
End Property

' ================================================================
' 関数名: SetValue
' 概要:   key=value を辞書に追加（既存 key は上書き）
' 引数:   ByVal key As String   - key 名
'         ByVal value As String - 値
' ================================================================
Public Sub SetValue(ByVal key As String, ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0351] ClsStanzaSection.SetValue ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0352] ClsStanzaSection.GetValue ENTER"  ' [ADR-0100]
    If m_keyValues Is Nothing Then
        GetValue = vbNullString
    ElseIf m_keyValues.Exists(key) Then
        GetValue = CStr(m_keyValues(key))
    Else
        GetValue = vbNullString
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0353] ClsStanzaSection.GetValue EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' 関数名: HasKey
' 概要:   key が定義されているか判定
' ================================================================
Public Function HasKey(ByVal key As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0354] ClsStanzaSection.HasKey ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0355] ClsStanzaSection.KeyCount ENTER"  ' [ADR-0100]
    If m_keyValues Is Nothing Then
        KeyCount = 0
    Else
        KeyCount = m_keyValues.Count
    End If
End Function
```