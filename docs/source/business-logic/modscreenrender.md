---
title: modScreenRender.bas
---

# modScreenRender.bas

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | 標準画面の描画委譲 (各 clsXxxScreen.Setup から呼ばれる共通処理) |
| 行数 | 90 行 |

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
' 依存先: IScreenRenderer, clsScreenSpec, modCommon, clsLogger
' 備考:   v21 (E2E rerun) で ENTER/EXIT/step ログを注入。
'         各画面クラスから呼ぶ公開ヘルパー LogScreenTrace/LogScreenError も提供。
' ================================================================

Private Const MOD_NAME As String = "modScreenRender"

' ================================================================
' 関数名: RenderStandardScreen
' 概要:   標準的な画面構築の流れを実行
' 引数:   renderer - IScreenRenderer 実装
'         spec     - clsScreenSpec データ
' ================================================================
Public Sub RenderStandardScreen(ByVal renderer As IScreenRenderer, _
                                 ByVal spec As clsScreenSpec)
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Dim sid As String : sid = ""
    On Error Resume Next
    sid = spec.ScreenId
    Err.Clear
    On Error GoTo ErrHandler

    Call LogScreenTrace(MOD_NAME, "RenderStandardScreen", "ENTER sid=" & sid)

    stepName = "BindSheet " & spec.SheetName
    renderer.BindSheet spec.SheetName

    ' 1) タイトル帯
    If Len(spec.Title) > 0 Then
        stepName = "RenderTitle"
        renderer.RenderTitle spec.ScreenId, spec.Title, spec.TitleColorHex
    End If

    ' 2) セクション帯
    stepName = "Sections"
    Dim sec As clsSectionSpec
    Dim secCount As Long : secCount = 0
    For Each sec In spec.Sections
        renderer.RenderSection sec.Address, sec.Label, sec.ColorHex
        secCount = secCount + 1
    Next sec

    ' 3) ボタン群
    stepName = "Buttons"
    Dim btn As clsButtonSpec
    Dim btnCount As Long : btnCount = 0
    For Each btn In spec.Buttons
        renderer.RenderButton btn
        btnCount = btnCount + 1
    Next btn

    ' 4) フィールドラベル群（データが空でも常時表示）
    stepName = "Fields"
    Dim fld As clsFieldSpec
    Dim fldCount As Long : fldCount = 0
    For Each fld In spec.Fields
        renderer.RenderInputField "", fld
        fldCount = fldCount + 1
    Next fld

    ' 5) 一覧ヘッダ行（M-02/M-07/M-08 等の一覧系のみ）
    If Len(spec.HeaderRowAddr) > 0 Then
        stepName = "HeaderRow " & spec.HeaderRowAddr
        renderer.RenderHeaderRow spec.HeaderRowAddr, spec.HeaderLabels, COLOR_BTN_PRIMARY
    End If

    ' 6) 空状態メッセージ（一覧系のみ）
    If Len(spec.EmptyStateAddr) > 0 Then
        stepName = "EmptyState " & spec.EmptyStateAddr
        renderer.RenderEmptyState spec.EmptyStateAddr, spec.EmptyStateText
    End If

    ' 7) ←メインに戻る ボタン（メイン以外の全画面に必置）
    If Len(spec.BackToMainAddr) > 0 Then
        stepName = "BackToMainButton " & spec.BackToMainAddr
        Call PlaceBackToMainButton(renderer, spec.BackToMainAddr)
    End If

    Call LogScreenTrace(MOD_NAME, "RenderStandardScreen", _
                         "EXIT sid=" & sid & " sec=" & secCount & " btn
```
