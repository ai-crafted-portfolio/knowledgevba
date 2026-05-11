---
title: modFactory.bas
---

# modFactory.bas

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | 画面層クラスと Renderer を生成するファクトリ。Init 注入の口を集約 |
| 行数 | 109 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modFactory` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modFactory"
Option Explicit

' ================================================================
' モジュール: modFactory（画面層 — ファクトリ）
' 概要:   Renderer / Screen クラスのインスタンス生成を集約する。
'         エントリポイントは「kind を文字列で指定して」インスタンスを得る。
' 依存先: clsSheetRenderer, clsUserFormRenderer,
'         clsMainScreen, clsKnowledgeRegisterScreen, ...（14 画面クラス）,
'         modScreenSpecRegistry
' ================================================================

Public Const RENDERER_KIND_SHEET As String = "sheet"
Public Const RENDERER_KIND_FORM As String = "form"

' ================================================================
' 関数名: CreateRenderer
' 概要:   Renderer 実装を生成（Strategy パターンの起点）
' 引数:   kind - "sheet" → clsSheetRenderer / "form" → clsUserFormRenderer (stub)
' 戻り値: IScreenRenderer
' ================================================================
Public Function CreateRenderer(ByVal kind As String) As IScreenRenderer
    Select Case LCase$(kind)
        Case RENDERER_KIND_SHEET
            Set CreateRenderer = New clsSheetRenderer
        Case RENDERER_KIND_FORM
            Set CreateRenderer = New clsUserFormRenderer
        Case Else
            Set CreateRenderer = New clsSheetRenderer  ' 既定
    End Select
End Function

' ================================================================
' 関数名: CreateScreen
' 概要:   screenId 対応の画面クラスインスタンスを Init 済みで返す
' 引数:   screenId - "M-01" 〜 "M-14"
'         renderer - IScreenRenderer 実装
' 戻り値: Object - 各 clsXxxScreen インスタンス（Init 済み）
' 備考:   全画面クラスは Init / Setup / Render の同一シグネチャを持つので
'         Object 型で返してダックタイピングする。
' ================================================================
Public Function CreateScreen(ByVal screenId As String, _
                              ByVal renderer As IScreenRenderer) As Object
    Dim spec As clsScreenSpec
    Set spec = modScreenSpecRegistry.GetScreenSpec(screenId)
    If spec Is Nothing Then
        Set CreateScreen = Nothing
        Exit Function
    End If

    Select Case screenId
        Case "M-01"
            Dim m As clsMainScreen: Set m = New clsMainScreen
            m.Init renderer, spec
            Set CreateScreen = m
        Case "M-02"
            Dim fl As clsFormatListScreen: Set fl = New clsFormatListScreen
            fl.Init renderer, spec
            Set CreateScreen = fl
        Case "M-03"
            Dim fd As clsFormatDesignScreen: Set fd = New clsFormatDesignScreen
            fd.Init renderer, spec
            Set CreateScreen = fd
        Case "M-04"
            Dim fp As clsFormatPreviewScreen: Set fp = New clsFormatPreviewScreen
            fp.Init renderer, spec
            Set CreateScreen = fp
        Case "M-05"
            Dim kr As clsKnowledgeRegisterScreen: Set kr = New clsKnowledgeRegisterScreen
            kr.Init renderer, spec
            Set CreateScreen = kr
        Case "M-06"
            Dim ke As clsKnowledgeEditScreen: Set ke = New clsKnowledgeEditScreen
            ke.Init renderer, spec
            Set CreateScreen = ke
        Case "M-07"
            Dim kl As clsKnowledgeListScreen: Set kl = New clsKnowledgeListScreen
            kl.Init renderer, spec
            Set CreateScreen = kl
        Case "M-08"
            Dim sc As clsSearchScreen: Set sc = New clsSearchScreen
            sc.Init renderer, spec
            Set CreateScreen = sc
        Case "M-09"
            Dim kv As clsKnowledgeViewScreen: Set kv = New clsKnowledgeViewScreen
            kv.Init renderer, spec
            Set CreateScreen = kv
        Case "M-10"
            Dim st As clsStorageScreen: Set st = New clsStorageScreen
            st.Init renderer, spec
            Set CreateScreen = st
        Case "M-11"
            Dim ss As clsSystemSettingsScreen: Set ss = New clsSystemSettingsScreen
            ss.Init renderer, spec
            Set CreateScreen = ss
        Case "M-12"
            Dim mg As clsMigrationScreen: Set mg = New clsMigrationScreen
            mg.Init renderer, spec
            Set CreateScreen = mg
        Case "M-13"
            Dim ff As clsFileFormatScreen: Set ff = New clsFileFormatScreen
            ff.Init renderer, spec
            Set CreateScreen = ff
        Case "M-14"
            Dim lg As clsLogScreen: Set lg = New clsLogScreen
            lg.Init renderer, spec
            Set CreateScreen = lg
    End Select
End Function
```
