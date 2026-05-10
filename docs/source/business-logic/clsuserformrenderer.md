---
title: clsUserFormRenderer.cls
---

# clsUserFormRenderer.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | ScreenSpec → UserForm への動的描画 |
| 行数 | 81 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsUserFormRenderer` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUserFormRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IScreenRenderer

' ================================================================
' クラス: clsUserFormRenderer（画面層 — 実装2 / スタブ）
' 概要:   IScreenRenderer の UserForm 型実装。今回は未実装のスタブ。
'         将来「フォーム入出力型」に画面方式を切り替える時の入口を確保する。
'         切替時はこのクラスのメソッドを実装すれば、画面クラス本体（clsXxxScreen）
'         は無修正で動作する（Strategy パターンの恩恵）。
' 依存先: なし（スタブ）
' 備考:   全メソッドが NotImplemented エラーを Raise する。
' ================================================================

Private Const NOT_IMPL_NUM As Long = vbObjectError + 1
Private Const NOT_IMPL_SRC As String = "clsUserFormRenderer"
Private Const NOT_IMPL_MSG As String = _
    "NotImplemented: 将来 UserForm 切替時に実装してください（design.md §9 参照）"

Private Sub IScreenRenderer_BindSheet(ByVal sheetName As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_ClearScreen()
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderTitle(ByVal screenId As String, _
                                         ByVal title As String, _
                                         ByVal colorHex As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderSection(ByVal sectionAddr As String, _
                                           ByVal sectionLabel As String, _
                                           ByVal colorHex As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderButton(ByVal btnSpec As Object)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderLabel(ByVal cellAddr As String, _
                                         ByVal labelText As String, _
                                         ByVal colorHex As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderInputField(ByVal cellAddr As String, _
                                              ByVal fieldSpec As Object)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderRequiredMark(ByVal cellAddr As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderHint(ByVal cellAddr As String, ByVal hintText As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderHeaderRow(ByVal startAddr As String, _
                                             ByVal headerLabels As Variant, _
                                             ByVal colorHex As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_RenderEmptyState(ByVal cellAddr As String, _
                                              ByVal message As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

```
