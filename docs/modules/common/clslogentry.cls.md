---
title: clsLogEntry.cls
description: clsLogEntry.cls 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# clsLogEntry.cls

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 繧ｯ繝ｩ繧ｹ繝｢繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`clsLogEntry.cls`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsLogEntry（ValueObject）
' 概要:   ログ1行の値を保持する軽量クラス
' 依存先: なし
' ================================================================

Private m_timestamp As String
Private m_moduleName As String
Private m_functionName As String
Private m_level As String
Private m_message As String

' --- Property Get/Let ---

Public Property Get Timestamp() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0537] clsLogEntry.Timestamp ENTER"  ' [ADR-0100]
    Timestamp = m_timestamp
End Property

Public Property Let Timestamp(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0538] clsLogEntry.Timestamp ENTER"  ' [ADR-0100]
    m_timestamp = value
End Property

Public Property Get ModuleName() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0539] clsLogEntry.ModuleName ENTER"  ' [ADR-0100]
    ModuleName = m_moduleName
End Property

Public Property Let ModuleName(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0540] clsLogEntry.ModuleName ENTER"  ' [ADR-0100]
    m_moduleName = value
End Property

Public Property Get FunctionName() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0541] clsLogEntry.FunctionName ENTER"  ' [ADR-0100]
    FunctionName = m_functionName
End Property

Public Property Let FunctionName(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0542] clsLogEntry.FunctionName ENTER"  ' [ADR-0100]
    m_functionName = value
End Property

Public Property Get Level() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0543] clsLogEntry.Level ENTER"  ' [ADR-0100]
    Level = m_level
End Property

Public Property Let Level(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0544] clsLogEntry.Level ENTER"  ' [ADR-0100]
    m_level = value
End Property

Public Property Get Message() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0545] clsLogEntry.Message ENTER"  ' [ADR-0100]
    Message = m_message
End Property

Public Property Let Message(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0546] clsLogEntry.Message ENTER"  ' [ADR-0100]
    m_message = value
End Property

' ================================================================
' 関数名: Init
' 概要:   全プロパティを一括設定する初期化メソッド
' 引数:   ts    - 日時文字列
'         mod_  - モジュール名
'         func  - 関数名
'         lvl   - ログレベル
'         msg   - メッセージ
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal ts As String, _
                 ByVal mod_ As String, _
                 ByVal func As String, _
                 ByVal lvl As String, _
                 ByVal msg As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0547] clsLogEntry.Init ENTER"  ' [ADR-0100]
    m_timestamp = ts
    m_moduleName = mod_
    m_functionName = func
    m_level = lvl
    m_message = msg
End Sub
```