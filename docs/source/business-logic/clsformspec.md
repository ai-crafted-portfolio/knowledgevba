---
title: clsFormSpec.cls
---

# clsFormSpec.cls

| 項目 | 内容 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | UserForm 1 つの宣言情報を保持する値オブジェクト |
| 行数 | 97 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsFormSpec.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsFormSpec (ビジネスロジック層)
' 概要:   1 つの UserForm の宣言情報を保持。
'         FormTitle / Width / Height + Controls Collection。
'         AddControl で clsControlSpec を追加する DSL を提供。
' 依存先: clsControlSpec
' 備考:   modFormBuilder.BuildAndShow に渡して動的フォーム生成する。
' ================================================================

Private m_formTitle As String
Private m_width As Long
Private m_height As Long
Private m_controls As Collection

Private Sub Class_Initialize()
    Set m_controls = New Collection
    m_width = 480
    m_height = 360
    m_formTitle = "Untitled"
End Sub

' --- FormTitle ---
Public Property Get FormTitle() As String
    FormTitle = m_formTitle
End Property
Public Property Let FormTitle(ByVal v As String)
    m_formTitle = v
End Property

' --- Width ---
Public Property Get Width() As Long
    Width = m_width
End Property
Public Property Let Width(ByVal v As Long)
    m_width = v
End Property

' --- Height ---
Public Property Get Height() As Long
    Height = m_height
End Property
Public Property Let Height(ByVal v As Long)
    m_height = v
End Property

' --- Controls (read-only Collection) ---
Public Property Get Controls() As Collection
    Set Controls = m_controls
End Property

' ================================================================
' 関数名: AddControl
' 概要:   clsControlSpec を生成して Controls Collection に追加する DSL。
'         同名 Control が既にあれば例外を投げる (Collection の挙動)。
' 引数:   cType / nm / l / t / w / h / cap / onClk - clsControlSpec.Init と同じ
' 戻り値: clsControlSpec - 追加された ControlSpec (チェイン用)
' ================================================================
Public Function AddControl(ByVal cType As String, ByVal nm As String, _
                            ByVal l As Long, ByVal t As Long, _
                            ByVal w As Long, ByVal h As Long, _
                            Optional ByVal cap As String = "", _
                            Optional ByVal onClk As String = "") As clsControlSpec
    Dim cs As clsControlSpec
    Set cs = New clsControlSpec
    cs.Init cType, nm, l, t, w, h, cap, onClk
    m_controls.Add cs, nm
    Set AddControl = cs
End Function

' ================================================================
' 関数名: ControlByName
' 概要:   指定名の Control Spec を取得 (なければ Nothing)
' ================================================================
Public Function ControlByName(ByVal nm As String) As clsControlSpec
    On Err