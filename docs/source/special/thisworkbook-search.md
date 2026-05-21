---
title: ThisWorkbook（検索.xlsm）
---

# ThisWorkbook（検索.xlsm）

| 項目 | 内容 |
|---|---|
| 層 | 特殊モジュール |
| 種別 | ドキュメントモジュール (ThisWorkbook) |
| 配置ブック | 検索.xlsm |
| 役割 | 検索.xlsm 用の ThisWorkbook。起動時の設定読込・セットアップ・終了ログ |
| 行数 | 97 行 |

## 取り込み先

ブックに最初から存在する `ThisWorkbook` モジュール（VBE のプロジェクトツリーにある `ThisWorkbook`）に貼り付けます。新規モジュールとしては取り込めません。先頭の `VERSION 1.0 CLASS` から `Attribute` 行までのファイル先頭ブロックは貼り付けず、その下の本体だけをコードペインに貼り付けて既存の内容を置き換えてください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ================================================================
' ThisWorkbook 特殊モジュール (検索.xlsm 専用、閲覧専用 xlsm)
' 配置:   ThisWorkbook.cls として検索.xlsm の VBE に直接 Import (VBE 仕様で Import 不可、コピペのみ)
' 関連:   clsSetupOrchestrator.cls v2, modConfigLoader.bas, modConfigHolder.bas
' Version: v2.1 (2026-05-16 EOD、Q1-Q57 反映)
' Phase: 6
' ADR:   ADR-0053 §2.1 (検索.xlsm = M-07/M-08/M-09 + LOG、起動時先頭 = M-08)
' v2.1:  Q44 で起動時 ActiveSheet = M-08 確定 (M-01 メニュー削除)
'        Q38 で M-09 ナレッジ表示は閲覧専用、編集ボタンなし (編集は登録修正.xlsm M-06 経由)
'        検索仕様: M-08 はキーワード + 多軸フィルタ検索（8 入力フィールド、canonical clsSearchEngine.cls + SSOT-Q32 裁定、screen_design_v2 §2.7）/ 100 件上限 (Q37)
'        Round 25 訂正: Q38 の閲覧専用画面は M-09 (旧採番 M-06)。多軸検索の根拠は canonical clsSearchEngine.cls + SSOT-Q32 裁定 (Q36 当初確定は free text 1 つ・多軸は Sprint 2 送り)。
' ================================================================
Option Explicit

Private Const XLSM_NAME As String = "検索"
Private Const STARTUP_SHEET As String = "M-08"  ' v2.1 Q44 確定 (M-01 メニュー削除、起動時 = 業務画面直行)

' ================================================================
' Workbook_Open
' 概要:   xlsm 起動時の setup を実行 (検索.xlsm 用)
' 手順:   1. modConfigLoader で xlsm 別対応の config.txt を read し modConfigHolder にセット (Q8)
'         2. clsLogger.Init (ログシート + debugLevel ERROR 既定、Q7)
'         3. clsSetupOrchestrator.RunFullSetup("検索")
'            - シート確保 (M-07/M-08/M-09/LOG)
'            - UI スタンザ適用 (modUILoader.ApplyUiToSheet、Q20)
'            - Workbook.Protect Structure + シート保護
'            - ActiveSheet = M-08 (Q44)
' ================================================================
Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    Call orch.RunFullSetup(XLSM_NAME)

    ' S5-LOG-02: SAVE-EXIT-OK-II-001 (Workbook_Open success exit, screen 検索)
    On Error Resume Next
    Dim oLogger001 As clsLogger
    Set oLogger001 = New clsLogger
    oLogger001.Init ThisWorkbook.Worksheets("LOG")
    oLogger001.LogInfo "ThisWorkbook_検索", "Workbook_Open", "Workbook_Open 正常完了: " & XLSM_NAME, "", "SAVE-EXIT-OK-II-001"
    On Error GoTo 0

    ' Q44: startup ActiveSheet = STARTUP_SHEET (production spec mirror)
    On Error Resume Next
    ThisWorkbook.Worksheets(STARTUP_SHEET).Activate
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Dim msg As String
    msg = "検索.xlsm 起動エラー: " & Err.Description & vbCrLf & _
          "config.txt の存在 / debugLevel 値 / シート構成を確認してください"
    Debug.Print msg
    MsgBox msg, vbCritical, "Workbook_Open"
End Sub

' ================================================================
' Workbook_BeforeClose
' 概要:   閉じる前に最終ログ
' R6-03 / R6-08: 3 xlsm 共通の LOG 書込仕様 (col 1-5)。
'   col 1 = timestamp / col 2 = "ThisWorkbook" / col 3 = "BeforeClose"
'   col 4 = "INFO" / col 5 = message ("xlsm 終了: " & XLSM_NAME)
' ================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ' ログシートに「終了」記録
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("LOG")
    If Not ws Is Nothing Then
        Dim r As Long
        r = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1
        If r < 9 Then r = 9
        ws.Cells(r, 1).Value = Format$(Now(), "yyyy-mm-dd hh:nn:ss")
        ws.Cells(r, 2).Value = "ThisWorkbook"
        ws.Cells(r, 3).Value = "BeforeClose"
        ws.Cells(r, 4).Value = "INFO"
        ws.Cells(r, 5).Value = "xlsm 終了: " & XLSM_NAME
    End If
    On Error GoTo 0
End Sub
```
