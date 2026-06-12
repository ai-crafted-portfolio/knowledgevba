---
title: clsLogEntry.cls
description: clsLogEntry.cls のソースコード（コピペ用）
---

# clsLogEntry.cls

**配置先**: 共通モジュール（3 ブック共通）
**種類**: クラスモジュール
**更新日**: 2026-06-03 23:22

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsLogEntry.cls`
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
