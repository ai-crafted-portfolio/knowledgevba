---
title: clsLogEntry.cls
---

# clsLogEntry.cls

| 項目 | 値 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | 1 行分のログレコード (値オブジェクト) |
| 行数 | 86 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsLogEntry` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

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
