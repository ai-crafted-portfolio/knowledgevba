---
title: clsCellIO.cls
---

# clsCellIO.cls

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | セル値の読み書きヘルパー（ワークシート／辞書モックの両対応） |
| 行数 | 57 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsCellIO.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
End
Attribute VB_Name = "clsCellIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' Cell value read/write helper (Sprint2 SRP split, ASCII-only)
' target = Worksheet (production) OR Scripting.Dictionary (mock)
' VB_PredeclaredId = True : callable as clsCellIO.Xxx
' ============================================================
Option Explicit

' Read cell value as String (handles Dict mock, Worksheet, errors, Date)
Public Function ReadCellValue(ByVal target As Object, ByVal cellAddr As String) As String
    If TypeName(target) = "Dictionary" Then
        If target.Exists(cellAddr) Then
            Dim dv As Variant
            dv = target(cellAddr)
            If IsNull(dv) Then
                ReadCellValue = ""
            Else
                ReadCellValue = CStr(dv)
            End If
        Else
            ReadCellValue = ""
        End If
    Else
        Dim r As Range
        Set r = target.Range(cellAddr)
        Dim v As Variant
        v = r.Cells(1, 1).Value
        If IsError(v) Then
            ReadCellValue = CStr(CVErr(v))
        ElseIf IsNull(v) Then
            ReadCellValue = ""
        ElseIf IsDate(v) Then
            ReadCellValue = Format(v, "yyyy-mm-dd")
        Else
            ReadCellValue = CStr(v)
        End If
    End If
End Function

' Write String value to target cell
Public Sub WriteCellValue(ByVal target As Object, ByVal cellAddr As String, ByVal value As String)
    If TypeName(target) = "Dictionary" Then
        target(cellAddr) = value
    Else
        Dim r As Range
        Set r = target.Range(cellAddr)
        r.Cells(1, 1).Value = value
    End If
End Sub
```
