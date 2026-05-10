---
title: modDateUtil.bas
---

# modDateUtil.bas

| 項目 | 値 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | 日付・時刻処理の純粋関数群 (LibreOffice 互換配慮) |
| 行数 | 103 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modDateUtil` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modDateUtil"
Option Explicit

' ================================================================
' モジュール: modDateUtil（ユーティリティ層）
' 概要:       日付・時刻処理の純粋関数群
'             LibreOffice Basic互換（"hh:nn:ss"問題を回避）
' 依存先:     なし
' ================================================================

' ================================================================
' 関数名: NowStr
' 概要:   現在時刻を "yyyy-mm-dd HH:MM:SS" 形式で返す
'         LibreOffice Basic の Format(..., "nn") 問題を回避するため自前実装
' 引数:   なし
' 戻り値: String - "2026-04-19 15:52:19" 形式の文字列
' 備考:   Format(Now, "hh:nn:ss") は LibreOffice で "nn" が曜日略称に
'         解釈されるため使わない
' ================================================================
Public Function NowStr() As String
    Dim d As Date
    d = Now()
    NowStr = Pad2(Year(d)) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d)) & " " & _
             Pad2(Hour(d)) & ":" & Pad2(Minute(d)) & ":" & Pad2(Second(d))
End Function

' ================================================================
' 関数名: TodayStr
' 概要:   今日の日付を "yyyy-mm-dd" 形式で返す
' 引数:   なし
' 戻り値: String - "2026-04-19" 形式の文字列
' ================================================================
Public Function TodayStr() As String
    Dim d As Date
    d = Date
    TodayStr = Pad2(Year(d)) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d))
End Function

' ================================================================
' 関数名: IsDateInRange
' 概要:   target が fromDate <= target <= toDate の範囲内かを判定
' 引数:   target   - 判定対象の日付
'         fromDate - 範囲の開始（0の場合は下限なし）
'         toDate   - 範囲の終了（0の場合は上限なし）
' 戻り値: Boolean - 範囲内なら True
' ================================================================
Public Function IsDateInRange(ByVal target As Date, _
                                ByVal fromDate As Date, _
                                ByVal toDate As Date) As Boolean
    ' fromDate=0 は下限なし、toDate=0 は上限なし扱い
    If fromDate > 0 And target < fromDate Then
        IsDateInRange = False
        Exit Function
    End If
    If toDate > 0 And target > toDate Then
        IsDateInRange = False
        Exit Function
    End If
    IsDateInRange = True
End Function

' ================================================================
' 関数名: TryParseDate
' 概要:   文字列を日付に変換する。失敗時は False を返す
' 引数:   s        - 変換対象文字列（"2026-04-19" 等）
'         outDate  - 出力: 変換成功時の日付
' 戻り値: Boolean - 変換成功なら True
' 備考:   空文字列の場合は False（outDateは0のまま）
' ================================================================
Public Function TryParseDate(ByVal s As String, _
                               ByRef outDate As Date) As Boolean
    On Error GoTo ErrHandler
    
    If Trim(s) = "" Then
        outDate = 0
        TryParseDate = False
        Exit Function
    End If
    
    outDate = CDate(s)
    TryParseDate = True
    Exit Function

ErrHandler:
    outDate = 0
    TryParseDate = False
End Function

' ================================================================
' 関数名: Pad2
' 概要:   数値を2桁ゼロパディングした文字列に変換
'         例: 5 -> "05"、10 -> "10"
' 引数:   n - 変換対象の数値
' 戻り値: String - 2桁の文字列
' 備考:   100以上の値が来た場合はそのまま文字列化する
' ================================================================
Public Function Pad2(ByVal n As Long) As String
    If n < 10 Then
        Pad2 = "0" & CStr(n)
    Else
        Pad2 = CStr(n)
    End If
End Function

```
