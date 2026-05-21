---
title: modCommon.bas
---

# modCommon.bas

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 3 ブック共通 |
| 役割 | 全ブック共通の定数群 |
| 行数 | 46 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modCommon.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modCommon"
' ================================================================
' モジュール: modCommon（v2.1、共通定数）
' 概要:       knowledgevba v2.1 で共有する定数群
' Version:    v2.1（2026-05-17、最小実装、test ハーネス compile 用途）
' ================================================================
Option Explicit

' バックアップ保持日数（Q34）
Public Const BACKUP_RETENTION_DAYS As Long = 90

' knowledge .txt 新スタンザ形式 区切り識別子（Q47）
Public Const KNW_STANZA_PREFIX As String = "###"
Public Const KNW_STANZA_SUFFIX As String = "###"

' 検索結果最大表示件数（Q37）
Public Const SEARCH_MAX_RESULTS As Long = 100

' debugLevel 既定値（Q7）
Public Const DEFAULT_DEBUG_LEVEL As String = "ERROR"

' 起動時 ActiveSheet（ADR-0053 §2.1 表 / §9 M-2、R6-01 是正）
' SSOT: ADR-0053 §2.1。登録修正=M-05、検索=M-08、管理=M-02
Public Const STARTUP_SHEET_TOUROKU As String = "M-05"
Public Const STARTUP_SHEET_KENSAKU As String = "M-08"
Public Const STARTUP_SHEET_KANRI As String = "M-02"

' ログシート行（v1 から維持）
Public Const LOG_DATA_START_ROW As Long = 9

' スタンザ書式（v1 から維持）
Public Const CHARSET_SJIS As String = "Shift_JIS"
Public Const STANZA_SEP As String = "==="

' xlsm 名（v2.1）
Public Const XLSM_TOUROKU As String = "登録修正"
Public Const XLSM_KENSAKU As String = "検索"
Public Const XLSM_KANRI As String = "管理"

' debugLevel enum 6 値（modConfigHolder にもあるが Public Const 整合のため本書で重複定義）
Public Const DEBUG_LEVEL_OFF As Long = 0
Public Const DEBUG_LEVEL_ERROR As Long = 1
Public Const DEBUG_LEVEL_WARN As Long = 2
Public Const DEBUG_LEVEL_INFO As Long = 3
Public Const DEBUG_LEVEL_DEBUG As Long = 4
Public Const DEBUG_LEVEL_TRACE As Long = 5
```
