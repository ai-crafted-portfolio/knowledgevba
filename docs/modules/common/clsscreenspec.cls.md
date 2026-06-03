---
title: clsScreenSpec.cls
description: clsScreenSpec.cls のソースコード（コピペ用）
---

# clsScreenSpec.cls

**配置先**: 共通モジュール（3 ブック共通）
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsScreenSpec.cls`
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
Attribute VB_Name = "clsScreenSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsScreenSpec（画面層 ? データ）
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

' --- ADR-0045 FIX-RNDR-02: SubTitle (タイトル行直下の補助情報 行 2 への分離) ---
Private m_subTitleText As String      ' "FMT001 (新規)" 等。空文字なら描画スキップ
Private m_subTitleAddr As String      ' "A2" 等。既定値 "A2" (Property Get で対応)

' --- ADR-0045 FIX-RNDR-03: ColumnWidths (列幅 SSOT)
'     M-07/M-08/M-10/M-14 のテーブル右端列切れ問題 (C-5) 解消用 ---
Private m_columnWidths As Variant     ' Array of Double。例: Array(4, 12, 12, 10, 10, 22, 10, 8, 6, 6, 8)
                                      ' i 番目要素 = (i+1) 列目 (A=1, B=2, ...) の幅 (文字数単位)
                                      ' 未設定 (Empty) の場合は ColumnWidths 設定スキップ (legacy 動作)

Private Sub Class_Initialize()
    Set m_sections = New Collection
    Set m_buttons = New Collection
    Set m_fields = New Collection
    m_subTitleText = ""
    m_subTitleAddr = "A2"       ' 既定: タイトル行 (行 1) の直下
    m_columnWidths = Empty
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

' --- ADR-0045 FIX-RNDR-02: SubTitle Property ---
Public Property Get SubTitleText() As String
    SubTitleText = m_subTitleText
End Property
Public Property Let SubTitleText(ByVal value As String)
    m_subTitleText = value
End Property

Public Property Get SubTitleAddr() As String
    SubTitleAddr = m_subTitleAddr
End Property
Public Property Let SubTitleAddr(ByVal value As String)
    m_subTitleAddr = value
End Property

' --- ADR-0045 FIX-RNDR-03: ColumnWidths Property ---
' 値は Array(width_a, width_b, width_c, ...) の Variant 配列。文字数単位 (Excel 既定の ColumnWidth と同単位)
' 未設定 (Empty) なら ColumnWidths は触らず Excel デフォルトのまま
Public Property Get ColumnWidths() As Variant
    ColumnWidths = m_columnWidths
End Property
Public Property Let ColumnWidths(ByVal value As Variant)
    m_columnWidths = value
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
