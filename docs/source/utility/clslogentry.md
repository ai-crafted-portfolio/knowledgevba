---
title: clsLogEntry.cls
---

# clsLogEntry.cls

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | ログ 1 行分の値オブジェクト |
| 行数 | 86 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsLogEntry.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
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
    Timestamp = m_timestamp
End Property

Public Property Let Timestamp(ByVal value As String)
    m_timestamp = value
End Property

Public Property Get ModuleName() As String
    ModuleName = m_moduleName
End Property

Public Property Let ModuleName(ByVal value As String)
    m_moduleName = value
End Property

Public Property Get FunctionName() As String
    FunctionName = m_functionName
End Property

Public Property Let FunctionName(ByVal value As String)
    m_functionName = value
End Property

Public Property Get Level() As String
    Level = m_level
End Property

Public Property Let Level(ByVal value As String)
    m_level = value
End Property

Public Property Get Message() As String
    Message = m_message
End Property

Public Property Let Message(ByVal value As String)
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
    m_timestamp = ts
    m_moduleName = mod_
    m_functionName = func
    m_level = lvl
    m_message = msg
End Sub
```
