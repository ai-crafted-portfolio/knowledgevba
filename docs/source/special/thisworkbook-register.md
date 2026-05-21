---
title: ThisWorkbook（登録修正.xlsm）
---

# ThisWorkbook（登録修正.xlsm）

| 項目 | 内容 |
|---|---|
| 層 | 特殊モジュール |
| 種別 | ドキュメントモジュール (ThisWorkbook) |
| 配置ブック | 登録修正.xlsm |
| 役割 | 登録修正.xlsm 用の ThisWorkbook。起動時の設定読込・セットアップ・終了ログ |
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
' ThisWorkbook 特殊モジュール (登録修正.xlsm 専用)
' 配置:   ThisWorkbook.cls として登録修正.xlsm の VBE に直接 Import (VBE 仕様で Import 不可、コピペのみ)
' 関連:   clsSetupOrchestrator.cls v2, modConfigLoader.bas, modConfigHolder.bas
' Version: v2.1 (2026-05-16 EOD、Q1-Q57 反映)
' Phase: 5
' ADR:   ADR-0053 §2.1 (登録修正.xlsm = M-05/M-06 + LOG)
' v2.1:  Q44 で起動時 ActiveSheet = M-05 確定 (M-01 メニュー削除)
'        Q44 で M-01 メイン画面廃止、本 xlsm の業務画面は M-05 + M-06（M-07/M-09/M-10/M-11/M-12 の削除案は Round 1 R1-13/R1-14 で取消済）
' ================================================================
Option Explicit

Private Const XLSM_NAME As String = "登録修正"
Private Const STARTUP_SHEET As String = "M-05"  ' v2.1 Q44 確定 (M-01 メニュー削除、起動時 = 業務画面直行)

' ================================================================
' Workbook_Open
' 概要:   xlsm 起動時の setup を実行
' 手順:   1. modConfigLoader で xlsm 別対応の config.txt を read し modConfigHolder にセット (Q8)
'         2. clsLogger.Init (ログシート + debugLevel ERROR 既定、Q7)
'         3. clsSetupOrchestrator.RunFullSetup("登録修正")
'            - シート確保 (M-05/M-06/LOG)
'            - UI スタンザ適用 (modUILoader.ApplyUiToSheet、Q20)
'            - Workbook.Protect Structure + シート保護
'            - ActiveSheet = M-05 (Q44)
' 備考:   Application.EnableEvents の On/Off 制御、エラー時 EnableEvents 復元保証
' ================================================================
Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim orch As clsSetupOrchestrator
    Set orch = New clsSetupOrchestrator
    Call orch.RunFullSetup(XLSM_NAME)

    ' S5-LOG-02: SAVE-EXIT-OK-II-002 (Workbook_Open success exit, screen 登録修正)
    On Error Resume Next
    Dim oLogger002 As clsLogger
    Set oLogger002 = New clsLogger
    oLogger002.Init ThisWorkbook.Worksheets("LOG")
    oLogger002.LogInfo "ThisWorkbook_登録修正", "Workbook_Open", "Workbook_Open 正常完了: " & XLSM_NAME, "", "SAVE-EXIT-OK-II-002"
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
    ' 起動時エラーは MsgBox + ログ (ログシートが未確保ならコンソールに Debug.Print)
    Dim msg As String
    msg = "登録修正.xlsm 起動エラー: " & Err.Description & vbCrLf & _
          "config.txt の存在 / debugLevel 値 / シート構成 を確認してください"
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
