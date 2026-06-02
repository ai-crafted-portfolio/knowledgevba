---
title: modFormatColumns.bas
description: modFormatColumns.bas のソースコード（コピペ用）
---

# modFormatColumns.bas

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modFormatColumns.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modFormatColumns"
Option Explicit

' ================================================================
' モジュール: modFormatColumns (release_v2 / v12 新規追加)
' 概要:       フォーマット一覧シートの列番号定数を Public 公開する。
'             v11 までは clsFormatManager.cls 内で Public Const 定義していたが、
'             VBA 仕様上「クラスモジュールで Public Const は不可」のため
'             real Excel が compile-time に reject (ADR 0027)。
'             本モジュール (.bas = 標準モジュール) に移動して合法化。
' 依存先:     なし
' 備考:       clsFormatManager 内部、modDemoSeeder 外部の両方から参照可能。
'             値は v11 までと同一で互換性維持。
' 関連 ADR:   0017 (D-2 Public 化), 0027 (cls 内 Public Const 禁止対応)
' ================================================================

' --- フォーマット一覧シート 列番号 ---
Public Const FL_COL_NO As Long = 1
Public Const FL_COL_FMT_ID As Long = 2
Public Const FL_COL_FMT_NAME As Long = 3
Public Const FL_COL_ID_PATTERN As Long = 4
Public Const FL_COL_CURRENT_NUM As Long = 5
Public Const FL_COL_NEXT_NUM As Long = 6
Public Const FL_COL_VERSION As Long = 7
Public Const FL_COL_FIELD_COUNT As Long = 8
Public Const FL_COL_KNW_COUNT As Long = 9
Public Const FL_COL_CREATED As Long = 10
Public Const FL_COL_UPDATED As Long = 11

' --- フォーマット一覧シート データ開始行 (1=ヘッダー領域) ---
Public Const FL_START_ROW As Long = 3
```
