---
title: clsScreenSpec.cls
---

# clsScreenSpec.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | 1 画面分の宣言情報 (タイトル / セクション / ボタン / フィールド) を保持するルート spec |
| 行数 | 135 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsScreenSpec` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScreenSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsScreenSpec（画面層 — データ）
' 概要:   1 画面分の構成（タイトル/セクション/ボタン/フィールド）を保持。
'         画面修正は本クラスのインスタンスを書き換えるだけで完結する設計。
' 依存先: clsSectionSpec, clsButtonSpec, clsFieldSpec
' ================================================================

Private m_screenId As String          ' "M-01"
Private m_sheetName As String         ' "メイン"
Private m_title As String             ' "【画面モック v2】..."
Private m_titleColorHex As String     ' "#1F3864"
Private m_backToMainAddr As String    ' "K17" など空文字許容
Private m_sections As Collection      ' of clsSectionSpec
Private m_buttons As Collection       ' of clsButtonSpec
Private m_fields As Collection        ' of clsFieldSpec
Private m_headerRowAddr As String     ' 一覧系: ヘッダ開始位置 ("B10")
Private m_headerLabels As Variant     ' 一覧系: ヘッダラベル配列
Private m_emptyStateAddr As String    ' "0 件" 表示位置
Private m_emptyStateText As String

Private Sub Class_Initialize()
    Set m_sections = New Collection
    Set m_buttons = New Collection
    Set m_fields = New Collection
End Sub

Public Property Get ScreenId() As String
    ScreenId = m_screenId
End Property
Public Property Let ScreenId(ByVal value As String)
    m_screenId = value
End Property

Public Property Get SheetName() As String
    SheetName = m_sheetName
End Property
Public Property Let SheetName(ByVal value As String)
    m_sheetName = value
End Property

Public Property Get Title() As String
    Title = m_title
End Property
Public Property Let Title(ByVal value As String)
    m_title = value
End Property

Public Property Get TitleColorHex() As String
    TitleColorHex = m_titleColorHex
End Property
Public Property Let TitleColorHex(ByVal value As String)
    m_titleColorHex = value
End Property

Public Property Get BackToMainAddr() As String
    BackToMainAddr = m_backToMainAddr
End Property
Public Property Let BackToMainAddr(ByVal value As String)
    m_backToMainAddr = value
End Property

Public Property Get Sections() As Collection
    Set Sections = m_sections
End Property

Public Property Get Buttons() As Collection
    Set Buttons = m_buttons
End Property

Public Property Get Fields() As Collection
    Set Fields = m_fields
End Property

Public Property Get HeaderRowAddr() As String
    HeaderRowAddr = m_headerRowAddr
End Property
Public Property Let HeaderRowAddr(ByVal value As String)
    m_headerRowAddr = value
End Property

Public Property Get HeaderLabels() As Variant
    HeaderLabels = m_headerLabels
End Property
Public Property Let HeaderLabels(ByVal value As Variant)
    m_headerLabels = value
End Property

Public Property Get EmptyStateAddr() As String
    EmptyStateAddr = m_emptyStateAddr
End Property
Public Property Let EmptyStateAddr(ByVal value As String)
    m_emptyStateAddr = value
End Property

Public Property Get EmptyStateText() As String
    EmptyStateText = m_emptyStateText
End Property
Public Property Let EmptyStateText(ByVal value As String)
    m_emptyStateText = value
End Property

' ================================================================
' 関数名: AddSection
' 概要:   セクション帯を追加
' ================================================================
Public Sub AddSection(ByVal sec As clsSectionSpec)
    m_sections.Add sec
End Sub

' ================================================================
' 関数名: AddButton
' 概要:   ボタンを追加
' ================================================================
Public Sub AddButton(ByVal btn As clsButtonSpec)
    m_buttons.Add btn
End Sub

' ================================================================
' 関数名: AddField
' 概要:   フィールドを追加
' ================================================================
Public Sub AddField(ByVal fld As clsFieldSpec)
    m_fields.Add fld
End Sub
```
