---
title: modScreenRender.bas
---

# modScreenRender.bas

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | 画面描画の薄いエントリ (Render(spec, sheet)) |
| 行数 | 87 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modScreenRender` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modScreenRender"
Option Explicit

' ================================================================
' モジュール: modScreenRender（画面層 — ユーティリティ）
' 概要:   各画面クラス（clsXxxScreen）が共通で使う「spec を Renderer に
'         流し込む」標準描画ロジック。各画面クラスはこれを Call することで
'         「タイトル / セクション帯 / ボタン / フィールドラベル / 一覧ヘッダ /
'          ←メインに戻る ボタン / 空状態」を一括で描画できる。
' 依存先: IScreenRenderer, clsScreenSpec, modCommon
' 備考:   各画面クラスの Render() 実装は本モジュールへの薄い委譲となる。
'         画面固有のロジック（OnButtonClick, OnFieldChange 等）は
'         各画面クラス側で実装する。
' ================================================================

' ================================================================
' 関数名: RenderStandardScreen
' 概要:   標準的な画面構築の流れを実行
' 引数:   renderer - IScreenRenderer 実装
'         spec     - clsScreenSpec データ
' ================================================================
Public Sub RenderStandardScreen(ByVal renderer As IScreenRenderer, _
                                 ByVal spec As clsScreenSpec)
    On Error GoTo ErrHandler

    renderer.BindSheet spec.SheetName

    ' 1) タイトル帯
    If Len(spec.Title) > 0 Then
        renderer.RenderTitle spec.ScreenId, spec.Title, spec.TitleColorHex
    End If

    ' 2) セクション帯
    Dim sec As clsSectionSpec
    For Each sec In spec.Sections
        renderer.RenderSection sec.Address, sec.Label, sec.ColorHex
    Next sec

    ' 3) ボタン群
    Dim btn As clsButtonSpec
    For Each btn In spec.Buttons
        renderer.RenderButton btn
    Next btn

    ' 4) フィールドラベル群（データが空でも常時表示）
    Dim fld As clsFieldSpec
    For Each fld In spec.Fields
        renderer.RenderInputField "", fld
    Next fld

    ' 5) 一覧ヘッダ行（M-02/M-07/M-08 等の一覧系のみ）
    If Len(spec.HeaderRowAddr) > 0 Then
        renderer.RenderHeaderRow spec.HeaderRowAddr, spec.HeaderLabels, COLOR_BTN_PRIMARY
    End If

    ' 6) 空状態メッセージ（一覧系のみ）
    If Len(spec.EmptyStateAddr) > 0 Then
        renderer.RenderEmptyState spec.EmptyStateAddr, spec.EmptyStateText
    End If

    ' 7) ←メインに戻る ボタン（メイン以外の全画面に必置）
    If Len(spec.BackToMainAddr) > 0 Then
        Call PlaceBackToMainButton(renderer, spec.BackToMainAddr)
    End If

    Exit Sub
ErrHandler:
    Debug.Print "[modScreenRender.RenderStandardScreen] ERROR " & spec.ScreenId & ": " & Err.Description
End Sub

' ================================================================
' 関数名: PlaceBackToMainButton
' 概要:   「←メインに戻る」緑ボタンを指定セルに配置
' ================================================================
Private Sub PlaceBackToMainButton(ByVal renderer As IScreenRenderer, _
                                    ByVal cellAddr As String)
    Dim btn As clsButtonSpec
    Set btn = New clsButtonSpec
    btn.Init "Btn_BackToMain", _
              ChrW(&H2190) & " メインに戻る", _
              cellAddr, _
              COLOR_BTN_NAV, _
              "ナビゲーション", _
              "", _
              ""
    renderer.RenderButton btn
End Sub

```
