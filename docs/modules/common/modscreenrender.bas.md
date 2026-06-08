---
title: modScreenRender.bas
description: modScreenRender.bas 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# modScreenRender.bas

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 讓呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`modScreenRender.bas`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
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