---
title: modEntrySettings.bas
---

# modEntrySettings.bas

| 項目 | 値 |
|---|---|
| 層 | エントリポイント層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | 設定シートのマクロボタン受け口 (dataFolder 切替) |
| 行数 | 35 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modEntrySettings` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modEntrySettings"
Option Explicit

' ================================================================
' モジュール: modEntrySettings（エントリポイント層）
' 概要:       設定シート関連のボタンに割り当てるマクロ群
' 依存先:     clsLogger, modEntryMain, modCommon
' ================================================================

' ================================================================
' 関数名: Btn_ResetLog
' 概要:   設定シート「▶リセット」ボタン
'         ログシートの2行目以降（データ行）を全削除
' ================================================================
Public Sub Btn_ResetLog()
    On Error GoTo ErrHandler
    
    If Not ConfirmAction("ログリセット", _
                           "ログシートの記録を全て削除します") Then
        Exit Sub
    End If
    
    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.ClearLog
    logger.LogInfo "modEntrySettings", "Btn_ResetLog", _
                    "ログをリセットしました"
    
    Call ShowInfo("ログリセット", "ログをリセットしました")
    Exit Sub

ErrHandler:
    Call ShowError("ログリセット", Err.Description, _
                    "再度ボタンを押してください")
End Sub
```

## 関連

- 呼び出す: `clsStorageResolver`, `clsLogger`
- 呼び出される: `ボタン: Btn_PickFolder`
