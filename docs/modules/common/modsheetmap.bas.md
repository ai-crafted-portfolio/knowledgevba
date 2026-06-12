---
title: modSheetMap.bas
description: modSheetMap.bas のソースコード（コピペ用）
---

# modSheetMap.bas

**配置先**: 共通モジュール（3 ブック共通）
**種類**: 標準モジュール
**更新日**: 2026-06-03 23:22

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modSheetMap.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modSheetMap"
Option Explicit

' ================================================================
' モジュール: modSheetMap（ユーティリティ層）
' 概要:       Sheet 名 ⇔ 画面 ID (M-01?M-14) の双方向 mapping を
'             一元提供する SSOT モジュール。
'             SSOT-27 (Rank B B0 FAIL) 解消のため Sprint 1 W1-P0-12
'             / S1-P0-23 で新規追加 (2026-05-15)。
' 設計参照:   modScreenSpecRegistry.bas BuildXxxSpec 群の
'             ScreenId / SheetName ペアを SSOT とする
'             (M-01: メイン / M-02: フォーマット一覧 / M-03: フォーマット設計 /
'              M-04: フォーマットプレビュー / M-05: ナレッジ登録 /
'              M-06: ナレッジ修正 / M-07: ナレッジ一覧 / M-08: 検索 /
'              M-09: ナレッジ表示 / M-10: 格納先設定 / M-11: 設定 /
'              M-12: 既存データへのフィールド反映 /
'              M-13: 廃止 (ADR-0072 §2.1) / M-14: ログ)
'             ※ M-15 は現行 SSOT に存在しない (将来予約)。
' 依存先:     modCommon (SHEET_* 定数)
' 利用例:
'     Dim sid As String
'     sid = SheetToScreenId(ActiveSheet.Name)  ' "M-08" 等
'     If sid = "" Then ... ' 未知シートはハンドリング
' ================================================================

' ================================================================
' 関数名: SheetToScreenId
' 概要:   シート名から画面 ID (M-01?M-14) を返す。
'         未知シート名は "" (空文字) を返す。
' 引数:   sheetName - シート名 (例: "メイン")
' 戻り値: String - 画面 ID (例: "M-01") / 未知なら ""
' 備考:   Select Case 実装 (Dictionary 初期化コスト回避 / Const 配列限界回避)。
'         比較は完全一致 (大文字小文字区別、trim せず)。
' ================================================================
Public Function SheetToScreenId(ByVal sheetName As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1368] modSheetMap.SheetToScreenId ENTER"  ' [ADR-0100]
    Select Case sheetName
        Case SHEET_MAIN:           SheetToScreenId = "M-01"
        Case SHEET_FORMAT_LIST:    SheetToScreenId = "M-02"
        Case SHEET_FORMAT_DESIGN:  SheetToScreenId = "M-03"
        Case SHEET_FORMAT_PREVIEW: SheetToScreenId = "M-04"
        Case SHEET_KNW_SAVE:       SheetToScreenId = "M-05"
        Case SHEET_KNW_EDIT:       SheetToScreenId = "M-06"
        Case SHEET_KNW_LIST:       SheetToScreenId = "M-07"
        Case SHEET_SEARCH:         SheetToScreenId = "M-08"
        Case SHEET_KNW_DISPLAY:    SheetToScreenId = "M-09"
        Case SHEET_STORAGE:        SheetToScreenId = "M-10"
        Case SHEET_SETTINGS:       SheetToScreenId = "M-11"
        Case SHEET_MIGRATION:      SheetToScreenId = "M-12"
        ' M-13 廃止 (ADR-0072 §2.1): SHEET_FILE_FORMAT は v2.3 で除外、mapping なし
        Case SHEET_LOG:            SheetToScreenId = "M-14"
        Case Else:                 SheetToScreenId = ""
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1369] modSheetMap.SheetToScreenId EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' 関数名: ScreenIdToSheet
' 概要:   画面 ID (M-01?M-14) からシート名を返す (逆引き)。
'         未知 ID は "" (空文字) を返す。
' 引数:   screenId - 画面 ID (例: "M-01")
' 戻り値: String - シート名 (例: "メイン") / 未知なら ""
' 備考:   SheetToScreenId と完全一致対応。modCommon の SHEET_* 定数を返す。
' ================================================================
Public Function ScreenIdToSheet(ByVal screenId As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1370] modSheetMap.ScreenIdToSheet ENTER"  ' [ADR-0100]
    Select Case screenId
        Case "M-01": ScreenIdToSheet = SHEET_MAIN
        Case "M-02": ScreenIdToSheet = SHEET_FORMAT_LIST
        Case "M-03": ScreenIdToSheet = SHEET_FORMAT_DESIGN
        Case "M-04": ScreenIdToSheet = SHEET_FORMAT_PREVIEW
        Case "M-05": ScreenIdToSheet = SHEET_KNW_SAVE
        Case "M-06": ScreenIdToSheet = SHEET_KNW_EDIT
        Case "M-07": ScreenIdToSheet = SHEET_KNW_LIST
        Case "M-08": ScreenIdToSheet = SHEET_SEARCH
        Case "M-09": ScreenIdToSheet = SHEET_KNW_DISPLAY
        Case "M-10": ScreenIdToSheet = SHEET_STORAGE
        Case "M-11": ScreenIdToSheet = SHEET_SETTINGS
        Case "M-12": ScreenIdToSheet = SHEET_MIGRATION
        ' M-13 廃止 (ADR-0072 §2.1): ScreenIdToSheet("M-13") は "" を返す (Case Else)
        Case "M-14": ScreenIdToSheet = SHEET_LOG
        Case Else:   ScreenIdToSheet = ""
    End Select
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1371] modSheetMap.ScreenIdToSheet EXIT-OK"  ' [ADR-0100]
End Function
```
