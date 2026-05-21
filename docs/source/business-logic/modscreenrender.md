---
title: modScreenRender.bas
---

# modScreenRender.bas

| 項目 | 内容 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 3 ブック共通 |
| 役割 | 各画面クラス共通のシート描画入口とログ補助 |
| 行数 | 96 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modScreenRender.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modScreenRender"
Option Explicit

' ================================================================
' モジュール: modScreenRender (画面層 ユーティリティ)
' 概要: 各画面クラスが共通で使うシート描画 entry + 画面層ログヘルパー
' 依存先: IScreenRenderer (v2 8-method), clsScreenSpec, modCommon, clsLogger
' 備考: 2026-05-20 v1→v2 migration。逐次 Render* パイプライン
'       (RenderTitle / RenderSection / RenderButton 等) を ApplyFromStanza
'       1 呼出に集約。ui_layout / coords 切替え部は削除 (architecture §807 #27)。
'       色指定 (COLOR_BTN_*) は modUILoader / UI スタンザ側に委譲。
' ================================================================

Private Const MOD_NAME As String = "modScreenRender"

' ================================================================
' 関数名: RenderStandardScreen
' 概要: 1 画面分の物理 UI をシートに再構築する。
'       v2: BindSheet → ClearScreen → ApplyFromStanza の 3 手順。
' 引数: renderer  - IScreenRenderer (v2 8-method)
'       spec      - clsScreenSpec (ScreenId / SheetName 等を参照)
'       xlsmName  - UI スタンザ配置先 xlsm 名 ("登録・修正" / "設定" / "管理")
'                   未指定 ("") の場合は ApplyFromStanza にそのまま
'                   渡し (modUILoader 内部で空ディレクトリ判定)。
'                   clsScreenSpec に xlsmName が無いため Optional 引数とし、
'                   呼出元 (画面層 cls) は無指定のまま無改修で通。
' ================================================================
Public Sub RenderStandardScreen(ByVal renderer As IScreenRenderer, _
                                 ByVal spec As clsScreenSpec, _
                                 Optional ByVal xlsmName As String = "")
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

    ' --- 画面遷移時に既存描画を全消し (clsSheetRenderer.ClearScreen が cell + shape 削除) ---
    stepName = "ClearScreen"
    renderer.ClearScreen

    ' --- v2: 物理 UI 構築は ApplyFromStanza 1 呼出に集約 (modUILoader 経由) ---
    ' 逐次 Render* (Title / Section / Button / Field / HeaderRow 等) は UI スタンザ
    ' .txt 定義側に委譲。色指定や coords 切替えも modUILoader 内で解決する。
    stepName = "ApplyFromStanza " & sid
    renderer.ApplyFromStanza xlsmName, sid

    Call LogScreenTrace(MOD_NAME, "RenderStandardScreen", "EXIT sid=" & sid)
    Exit Sub

ErrHandler:
    Call LogScreenError(MOD_NAME, "RenderStandardScreen", _
                         "sid=" & sid & " step=" & stepName, Err.Number, Err.Description)
End Sub

' ================================================================
' 関数名: LogScreenTrace
' 概要: 画面層共通の Trace ログ出力ラッパ (v1/v2 非依存)
' ================================================================
Public Sub LogScreenTrace(ByVal className As String, _
                            ByVal funcName As String, _
                            ByVal message As String, _
                            Optional ByVal logId As String = "")
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ThisWorkbook.Worksheets("LOG")
    If Not lg Is Nothing Then
        lg.LogTrace className, funcName, message, "", logId
    End If
End Sub

' ================================================================
' 関数名: LogScreenError
' 概要: 画面層共通の Error ログ出力ラッパ (v1/v2 非依存)
' ================================================================
Public Sub LogScreenError(ByVal className As String, _
                            ByVal funcName As String, _
                            ByVal stepName As String, _
                            ByVal errNum As Long, _
                            ByVal errDesc As String, _
                            Optional ByVal logId As String = "")
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ThisWorkbook.Worksheets("LOG")
    If Not lg Is Nothing Then
        lg.LogError className, funcName, "step=" & stepName & " Err#" & errNum & " " & errDesc, "", logId
    End If
End Sub
```
