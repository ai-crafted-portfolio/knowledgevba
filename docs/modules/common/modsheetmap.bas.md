---
title: modSheetMap.bas
description: modSheetMap.bas 縺ｮ繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝会ｼ医さ繝斐・逕ｨ・・---

# modSheetMap.bas

**驟咲ｽｮ蜈・*: `蜈ｨ繝悶ャ繧ｯ蜈ｱ騾啻 逕ｨ縺ｮ VBA 繝｢繧ｸ繝･繝ｼ繝ｫ
**遞ｮ鬘・*: 讓呎ｺ悶Δ繧ｸ繝･繝ｼ繝ｫ

---

## 繝輔ぃ繧､繝ｫ縺ｨ縺励※菫晏ｭ・
繝｡繝｢蟶ｳ・医∪縺溘・莉ｻ諢上・繝・く繧ｹ繝医お繝・ぅ繧ｿ・峨↓荳九・繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝牙・譁・ｒ雋ｼ繧贋ｻ倥￠縲・*`modSheetMap.bas`** 縺ｨ縺・≧蜷榊燕縺ｧ `installer\vba_modules\common\` 驟堺ｸ九↓菫晏ｭ倥＠縺ｦ縺上□縺輔＞縲よ枚蟄励さ繝ｼ繝峨・ ANSI・・hift-JIS・峨∵隼陦後・ CRLF 縺ｫ縺励※縺上□縺輔＞縲・
---

## 繧ｽ繝ｼ繧ｹ繧ｳ繝ｼ繝・
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