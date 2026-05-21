---
title: clsUserFormRenderer.cls
---

# clsUserFormRenderer.cls

| 項目 | 内容 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | 画面描画インターフェースの UserForm 実装（将来用の入口） |
| 行数 | 60 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsUserFormRenderer.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "clsUserFormRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' クラス: clsUserFormRenderer（v2.1、Phase 3 着手時の小物、Stub）
' 概要:   IScreenRenderer の UserForm 型実装、v2.1 でも未実装 Stub
'         将来 UserForm 化する場合の入口を確保（Strategy pattern）
' Version: v2.1（2026-05-16 EOD、Q5/Q20 整合）
' 関連:   clsSheetRenderer（thin shim 化）, IScreenRenderer.cls
' v2.1 主要更新:
'   - 8 method NotImplemented Stub に縮退（v1 14 method → v2.1 8 method）
'   - Q5: UI 定義 .txt は read-only、Save API なし（将来 UserForm 化時も同様）
' ================================================================
Implements IScreenRenderer
Option Explicit

Private Const NOT_IMPL_NUM As Long = vbObjectError + 1
Private Const NOT_IMPL_SRC As String = "clsUserFormRenderer"
Private Const NOT_IMPL_MSG As String = _
    "NotImplemented: UserForm 切替を実装する場合に本 cls を完成させてください"

' --- IScreenRenderer v2 実装（8 メソッド、全 NotImplemented Stub） ---

Private Sub IScreenRenderer_BindSheet(ByVal sheetName As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_ClearScreen()
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_ApplyFromStanza(ByVal xlsmName As String, ByVal screenId As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_ShowSheet()
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_HideSheet()
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_ActivateSheet()
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_ProtectSheet(ByVal level As String)
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub

Private Sub IScreenRenderer_UnprotectSheet()
    Err.Raise NOT_IMPL_NUM, NOT_IMPL_SRC, NOT_IMPL_MSG
End Sub
```
