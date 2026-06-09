---
title: modScreenRender.bas
description: modScreenRender.bas のソースコード（コピペ用）
---

# modScreenRender.bas

**配置先**: 共通モジュール（3 ブック共通）
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modScreenRender.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1331] modScreenRender.RenderStandardScreen ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    Dim stepName As String : stepName = "begin"
    Dim sid As String : sid = ""
    On Error Resume Next
    sid = spec.ScreenId
    Err.Clear
    On Error GoTo ErrHandler

    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1334] modScreenRender.RenderStandardScreen STEP-1 pre LogScreenTrace"  ' [ADR-0100]
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

    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1335] modScreenRender.RenderStandardScreen STEP-2 pre LogScreenTrace"  ' [ADR-0100]
    Call LogScreenTrace(MOD_NAME, "RenderStandardScreen", "EXIT sid=" & sid)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1332] modScreenRender.RenderStandardScreen EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1336] modScreenRender.RenderStandardScreen STEP-3 pre LogScreenError"  ' [ADR-0100]
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1333] modScreenRender.RenderStandardScreen EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1337] modScreenRender.LogScreenTrace ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ThisWorkbook.Worksheets("LOG")
    If Not lg Is Nothing Then
        lg.LogTrace className, funcName, message, "", logId
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1338] modScreenRender.LogScreenTrace EXIT-OK"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1339] modScreenRender.LogScreenError ENTER"  ' [ADR-0100]
    On Error Resume Next
    Dim lg As clsLogger
    Set lg = New clsLogger
    lg.Init ThisWorkbook.Worksheets("LOG")
    If Not lg Is Nothing Then
        lg.LogError className, funcName, "step=" & stepName & " Err#" & errNum & " " & errDesc, "", logId
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1340] modScreenRender.LogScreenError EXIT-OK"  ' [ADR-0100]
End Sub
```
