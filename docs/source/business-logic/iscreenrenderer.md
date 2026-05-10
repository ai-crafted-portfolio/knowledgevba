---
title: IScreenRenderer.cls
---

# IScreenRenderer.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | 画面描画インターフェース (画面層クラスが実装) |
| 行数 | 58 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `IScreenRenderer` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IScreenRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' インタフェース: IScreenRenderer（画面層）
' 概要:   画面要素を物理的に描画する責務を抽象化するインタフェース。
'         実装1: clsSheetRenderer（シート埋込型 — 今回採用）
'         実装2: clsUserFormRenderer（フォーム型 — 将来切替先のスタブ）
' 依存先: clsScreenSpec, clsButtonSpec, clsFieldSpec, clsSectionSpec
' 備考:   VBA の Implements パターン用。Public Sub のシグネチャだけ定義し
'         具体実装はそれぞれのクラスで Private Sub IScreenRenderer_xxx として書く。
' ================================================================

' ----- 画面初期化 -----
Public Sub BindSheet(ByVal sheetName As String)
End Sub

Public Sub ClearScreen()
End Sub

' ----- タイトル/セクション -----
Public Sub RenderTitle(ByVal screenId As String, ByVal title As String, ByVal colorHex As String)
End Sub

Public Sub RenderSection(ByVal sectionAddr As String, ByVal sectionLabel As String, ByVal colorHex As String)
End Sub

' ----- ボタン -----
Public Sub RenderButton(ByVal btnSpec As Object)
End Sub

' ----- ラベル/フィールド -----
Public Sub RenderLabel(ByVal cellAddr As String, ByVal labelText As String, ByVal colorHex As String)
End Sub

Public Sub RenderInputField(ByVal cellAddr As String, ByVal fieldSpec As Object)
End Sub

Public Sub RenderRequiredMark(ByVal cellAddr As String)
End Sub

Public Sub RenderHint(ByVal cellAddr As String, ByVal hintText As String)
End Sub

' ----- データ系 -----
Public Sub RenderHeaderRow(ByVal startAddr As String, ByVal headerLabels As Variant, ByVal colorHex As String)
End Sub

Public Sub RenderEmptyState(ByVal cellAddr As String, ByVal message As String)
End Sub

```
