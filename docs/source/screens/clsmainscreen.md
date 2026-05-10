---
title: clsMainScreen.cls
---

# clsMainScreen.cls

| 項目 | 値 |
|---|---|
| 層 | 画面層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | メイン画面 (M-01) の ScreenSpec 構築 |
| 行数 | 64 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsMainScreen` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMainScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsMainScreen（画面層）
' 概要:   M-01 メイン画面の構築・再描画。12 ボタンを 3 グループに整理。
' 依存先: IScreenRenderer, clsScreenSpec, modScreenSpecRegistry, modCommon
' ================================================================

Private m_renderer As IScreenRenderer
Private m_spec As clsScreenSpec

Public Sub Init(ByVal renderer As IScreenRenderer, ByVal spec As clsScreenSpec)
    Set m_renderer = renderer
    Set m_spec = spec
End Sub

' ================================================================
' 関数名: Setup
' 概要:   メイン画面の全要素（タイトル/グループ帯/12 ボタン+ヒント/凡例）を描画
' ================================================================
Public Sub Setup()
    On Error GoTo ErrHandler
    m_renderer.BindSheet m_spec.SheetName
    m_renderer.RenderTitle m_spec.ScreenId, m_spec.Title, m_spec.TitleColorHex

    ' ヘッダ（タスク表示 + バージョン）
    m_renderer.RenderLabel "B6", ChrW(&H1F4DA) & " ナレッジ管理システム", COLOR_SECTION_BLUE
    m_renderer.RenderLabel "B7", "現在のタスク:", ""
    m_renderer.RenderLabel "D7", "(未選択)", "#E7E6E6"
    m_renderer.RenderLabel "H7", "バージョン:", ""
    m_renderer.RenderLabel "J7", "v2.0", "#E7E6E6"

    ' グループ帯 + ボタン群
    Dim sec As clsSectionSpec
    For Each sec In m_spec.Sections
        m_renderer.RenderSection sec.Address, sec.Label, sec.ColorHex
    Next sec

    Dim btn As clsButtonSpec
    For Each btn In m_spec.Buttons
        m_renderer.RenderButton btn
    Next btn

    ' 凡例
    m_renderer.RenderHint "B21", "凡例: 青=主操作 / 緑=遷移 / 灰=副 / (橙=破壊 ― 各画面内に配置)"
    m_renderer.RenderHint "B22", ChrW(&H203B) & " 各画面の右上に必ず " & ChrW(&H2190) & "メインに戻る 緑ボタンが配置されています"

    Exit Sub
ErrHandler:
    Debug.Print "[clsMainScreen.Setup] ERROR: " & Err.Description
End Sub

Public Sub Render()
    Call Setup
End Sub

```
