---
title: ThisWorkbook.cls (Cowork デモ版)
---

# ThisWorkbook.cls (Cowork デモ版)

| 項目 | 値 |
|---|---|
| 層 | Cowork デモ専用 |
| 種別 | ドキュメントモジュール (ThisWorkbook) |
| 役割 | デモ版 ThisWorkbook (Workbook_Open で modAutoInit を 1 ティック遅延起動) |
| 行数 | 60 行 |

## 配置先

VBE のプロジェクトツリーで `ThisWorkbook` モジュールを開き、コードペインに**置換貼り付け**します。新規モジュールとしてインポートしないでください。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

**注**: 自動生成された VBA ファイルヘッダ部（`VERSION`, `BEGIN..END`, `Attribute VB_*`）は除外済みです。下のコードを VBE の ThisWorkbook コードペインにそのまま貼り付けて使えます（ヘッダ部は VBE が自動管理します）。

```vbnet linenums="1"
Option Explicit

' ================================================================
' クラス: ThisWorkbook (ドキュメントモジュール、Cowork デモ v3)
' 概要:   ブックイベントを薄く受け取り、auto-init は modAutoInit に委譲。
'         - Workbook_Open: Application.OnTime で 1 ティック遅延の
'           RunAutoInitDeferred を予約するだけ。
'           本体ロジック (シート判定 / Setup / Seed / アクティブ化) は
'           全て modAutoInit.RunAutoInitDeferred 内。
'         - Workbook_BeforeClose: 終了ログのみ。
' 依存先: modAutoInit (RunAutoInitDeferred), modEntryMain (BuildLogger)
' 備考:   D-1 部分対応: ThisWorkbook を「OnTime を呼ぶだけ」にすることで
'         本番 release_v2 の ThisWorkbook (auto-init 無し) との差分を最小化。
'         本番展開時はこの Workbook_Open 1 行を削除すれば等価になる。
' ================================================================

' ================================================================
' 関数名: Workbook_Open
' 概要:   ブックオープン時のイベント。
'         auto-init は Application.OnTime 経由で 1 ティック遅延起動する
'         (C-1: Workbook_Open context 内での Application.Run 不安定回避)。
' 引数:   なし
' 戻り値: なし
' ================================================================
Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    Application.OnTime Now, "RunAutoInitDeferred"
    Exit Sub
ErrHandler:
    MsgBox "起動に失敗しました: " & Err.Description, _
           vbCritical, "Workbook_Open エラー"
End Sub

' ================================================================
' 関数名: Workbook_BeforeClose
' 概要:   ブックを閉じる直前のイベント。終了ログのみ。
' 引数:   Cancel - キャンセルフラグ
' 戻り値: なし
' ================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo ErrHandler

    Dim logger As clsLogger
    Set logger = BuildLogger()
    logger.LogInfo "ThisWorkbook", "Workbook_BeforeClose", "システム終了"
    Exit Sub

ErrHandler:
    ' 終了時エラーは握りつぶす (ブックを閉じる妨げにしない)
End Sub
```

## 関連

- 呼び出す: `modAutoInit`, `modEntryMain`
- 呼び出される: `Excel イベントから直接呼び出し`
