---
title: clsScreenSpec.cls
description: clsScreenSpec.cls のソースコード（コピペ用）
---

# clsScreenSpec.cls

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: クラスモジュール
**更新日**: 2026-06-05 01:27 JST

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

Private m_screenId As String          ' 'M-01'
Private m_sheetName As String         ' 'メイン'
Private m_title As String             ' '【画面モック v2】...'
Private m_titleColorHex As String     ' '#1F3864'
Private m_backToMainAddr As String    ' 'K17' など空文字許容
Private m_sections As Collection      ' of clsSectionSpec
Private m_buttons As Collection       ' of clsButtonSpec
Private m_fields As Collection        ' of clsFieldSpec
Private m_headerRowAddr As String     ' 一覧系: ヘッダ開始位置 ('B10')
Private m_headerLabels As Variant     ' 一覧系: ヘッダラベル配列
Private m_emptyStateAddr As String    ' '0 件' 表示位置
Private m_emptyStateText As String

' --- ADR-0045 FIX-RNDR-02: SubTitle (タイトル行直下の補助情報 行 2 への分離) ---
Private m_subTitleText As String      ' 'FMT001 (新規)' 等。空文字なら描画スキップ
Private m_subTitleAddr As String      ' 'A2' 等。既定値 'A2' (Property Get で対応)

' --- ADR-0045 FIX-RNDR-03: ColumnWidths (列幅 SSOT)
'     M-07/M-08/M-10/M-14 のテーブル右端列切れ問題 (C-5) 解消用 ---
Private m_columnWidths As Variant     ' Array of Double。例: Array(4, 12, 12, 10, 10, 22, 10, 8, 6, 6, 8)
                                      ' i 番目要素 = (i+1) 列目 (A=1, B=2, ...) の幅 (文字数単位)
                                      ' 未設定 (Empty) の場合は ColumnWidths 設定スキップ (legacy 動作)

Private Sub Class_Initialize()
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0585] clsScreenSpec.Class_Initialize ENTER"  ' [ADR-0100]
    Set m_sections = New Collection
    Set m_buttons = New Collection
    Set m_fields = New Collection
    m_subTitleText = ""
    m_subTitleAddr = "A2"       ' 既定: タイトル行 (行 1) の直下
    m_columnWidths = Empty
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0586] clsScreenSpec.Class_Initialize EXIT-OK"  ' [ADR-0100]
End Sub

Public Property Get ScreenId() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0587] clsScreenSpec.ScreenId ENTER"  ' [ADR-0100]
    ScreenId = m_screenId
End Property
Public Property Let ScreenId(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0588] clsScreenSpec.ScreenId ENTER"  ' [ADR-0100]
    m_screenId = value
End Property

Public Property Get SheetName() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0589] clsScreenSpec.SheetName ENTER"  ' [ADR-0100]
    SheetName = m_sheetName
End Property
Public Property Let SheetName(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0590] clsScreenSpec.SheetName ENTER"  ' [ADR-0100]
    m_sheetName = value
End Property

Public Property Get Title() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0591] clsScreenSpec.Title ENTER"  ' [ADR-0100]
    Title = m_title
End Property
Public Property Let Title(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0592] clsScreenSpec.Title ENTER"  ' [ADR-0100]
    m_title = value
End Property

Public Property Get TitleColorHex() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0593] clsScreenSpec.TitleColorHex ENTER"  ' [ADR-0100]
    TitleColorHex = m_titleColorHex
End Property
Public Property Let TitleColorHex(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0594] clsScreenSpec.TitleColorHex ENTER"  ' [ADR-0100]
    m_titleColorHex = value
End Property

Public Property Get BackToMainAddr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0595] clsScreenSpec.BackToMainAddr ENTER"  ' [ADR-0100]
    BackToMainAddr = m_backToMainAddr
End Property
Public Property Let BackToMainAddr(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0596] clsScreenSpec.BackToMainAddr ENTER"  ' [ADR-0100]
    m_backToMainAddr = value
End Property

Public Property Get Sections() As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0597] clsScreenSpec.Sections ENTER"  ' [ADR-0100]
    Set Sections = m_sections
End Property

Public Property Get Buttons() As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0598] clsScreenSpec.Buttons ENTER"  ' [ADR-0100]
    Set Buttons = m_buttons
End Property

Public Property Get Fields() As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0599] clsScreenSpec.Fields ENTER"  ' [ADR-0100]
    Set Fields = m_fields
End Property

Public Property Get HeaderRowAddr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0600] clsScreenSpec.HeaderRowAddr ENTER"  ' [ADR-0100]
    HeaderRowAddr = m_headerRowAddr
End Property
Public Property Let HeaderRowAddr(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0601] clsScreenSpec.HeaderRowAddr ENTER"  ' [ADR-0100]
    m_headerRowAddr = value
End Property

Public Property Get HeaderLabels() As Variant
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0602] clsScreenSpec.HeaderLabels ENTER"  ' [ADR-0100]
    HeaderLabels = m_headerLabels
End Property
Public Property Let HeaderLabels(ByVal value As Variant)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0603] clsScreenSpec.HeaderLabels ENTER"  ' [ADR-0100]
    m_headerLabels = value
End Property

Public Property Get EmptyStateAddr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0604] clsScreenSpec.EmptyStateAddr ENTER"  ' [ADR-0100]
    EmptyStateAddr = m_emptyStateAddr
End Property
Public Property Let EmptyStateAddr(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0605] clsScreenSpec.EmptyStateAddr ENTER"  ' [ADR-0100]
    m_emptyStateAddr = value
End Property

Public Property Get EmptyStateText() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0606] clsScreenSpec.EmptyStateText ENTER"  ' [ADR-0100]
    EmptyStateText = m_emptyStateText
End Property
Public Property Let EmptyStateText(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0607] clsScreenSpec.EmptyStateText ENTER"  ' [ADR-0100]
    m_emptyStateText = value
End Property

' --- ADR-0045 FIX-RNDR-02: SubTitle Property ---
Public Property Get SubTitleText() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0608] clsScreenSpec.SubTitleText ENTER"  ' [ADR-0100]
    SubTitleText = m_subTitleText
End Property
Public Property Let SubTitleText(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0609] clsScreenSpec.SubTitleText ENTER"  ' [ADR-0100]
    m_subTitleText = value
End Property

Public Property Get SubTitleAddr() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0610] clsScreenSpec.SubTitleAddr ENTER"  ' [ADR-0100]
    SubTitleAddr = m_subTitleAddr
End Property
Public Property Let SubTitleAddr(ByVal value As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0611] clsScreenSpec.SubTitleAddr ENTER"  ' [ADR-0100]
    m_subTitleAddr = value
End Property

' --- ADR-0045 FIX-RNDR-03: ColumnWidths Property ---
' 値は Array(width_a, width_b, width_c, ...) の Variant 配列。文字数単位 (Excel 既定の ColumnWidth と同単位)
' 未設定 (Empty) なら ColumnWidths は触らず Excel デフォルトのまま
Public Property Get ColumnWidths() As Variant
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0612] clsScreenSpec.ColumnWidths ENTER"  ' [ADR-0100]
    ColumnWidths = m_columnWidths
End Property
Public Property Let ColumnWidths(ByVal value As Variant)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0613] clsScreenSpec.ColumnWidths ENTER"  ' [ADR-0100]
    m_columnWidths = value
End Property

' ================================================================
' 関数名: AddSection
' 概要:   セクション帯を追加
' ================================================================
Public Sub AddSection(ByVal sec As clsSectionSpec)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0614] clsScreenSpec.AddSection ENTER"  ' [ADR-0100]
    m_sections.Add sec
End Sub

' ================================================================
' 関数名: AddButton
' 概要:   ボタンを追加
' ================================================================
Public Sub AddButton(ByVal btn As clsButtonSpec)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0615] clsScreenSpec.AddButton ENTER"  ' [ADR-0100]
    m_buttons.Add btn
End Sub

' ================================================================
' 関数名: AddField
' 概要:   フィールドを追加
' ================================================================
Public Sub AddField(ByVal fld As clsFieldSpec)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0616] clsScreenSpec.AddField ENTER"  ' [ADR-0100]
    m_fields.Add fld
End Sub
```
