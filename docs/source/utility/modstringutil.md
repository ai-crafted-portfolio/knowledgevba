---
title: modStringUtil.bas
---

# modStringUtil.bas

| 項目 | 値 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 役割 | 文字列処理の純粋関数群 (トリム / エスケープ / Split 拡張) |
| 行数 | 239 行 |

## 配置先

VBE で `挿入 > 標準モジュール`、F4 でプロパティ → `(オブジェクト名)` を `modStringUtil` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
Attribute VB_Name = "modStringUtil"
Option Explicit

' ================================================================
' モジュール: modStringUtil（ユーティリティ層）
' 概要:       文字列処理の純粋関数群
' 依存先:     modCommon
' ================================================================

' ================================================================
' 関数名: ContainsAllKeywords
' 概要:   対象文字列がキーワード（スペース区切り）を全て含むか判定（AND検索）
' 引数:   target   - 検索対象の文字列
'         keywords - スペース区切りのキーワード群（例: "メモリ サーバ"）
' 戻り値: Boolean - 全キーワードを含むなら True
' 備考:   大文字小文字を区別しない（LCase比較）
'         keywordsが空の場合は True を返す
' ================================================================
Public Function ContainsAllKeywords(ByVal target As String, _
                                      ByVal keywords As String) As Boolean
    Dim parts() As String
    Dim i As Long
    Dim targetLower As String
    
    If Trim(keywords) = "" Then
        ContainsAllKeywords = True
        Exit Function
    End If
    
    targetLower = LCase(target)
    parts = Split(Trim(keywords), " ")
    
    For i = LBound(parts) To UBound(parts)
        If Trim(parts(i)) <> "" Then
            If InStr(targetLower, LCase(parts(i))) = 0 Then
                ContainsAllKeywords = False
                Exit Function
            End If
        End If
    Next i
    
    ContainsAllKeywords = True
End Function

' ================================================================
' 関数名: ContainsAnyKeyword
' 概要:   対象文字列がキーワード（スペース区切り）のいずれかを含むか判定（OR検索）
' 引数:   target   - 検索対象の文字列
'         keywords - スペース区切りのキーワード群
' 戻り値: Boolean - いずれかのキーワードを含むなら True
' 備考:   大文字小文字を区別しない（LCase比較）
'         keywordsが空の場合は True を返す
' ================================================================
Public Function ContainsAnyKeyword(ByVal target As String, _
                                     ByVal keywords As String) As Boolean
    Dim parts() As String
    Dim i As Long
    Dim targetLower As String
    
    If Trim(keywords) = "" Then
        ContainsAnyKeyword = True
        Exit Function
    End If
    
    targetLower = LCase(target)
    parts = Split(Trim(keywords), " ")
    
    For i = LBound(parts) To UBound(parts)
        If Trim(parts(i)) <> "" Then
            If InStr(targetLower, LCase(parts(i))) > 0 Then
                ContainsAnyKeyword = True
                Exit Function
            End If
        End If
    Next i
    
    ContainsAnyKeyword = False
End Function

' ================================================================
' 関数名: FormatNumberByPattern
' 概要:   採番パターンと数値から、採番後の文字列を生成する
'         例: "INC-{0000}" + 5 -> "INC-0005"
' 引数:   pattern  - パターン文字列（{0000}等のプレースホルダ含む）
'         nextNum  - 採番する値
' 戻り値: String - パターンに数値を埋め込んだ文字列
' 備考:   {} で囲まれた 0 の個数がゼロパディング桁数になる
'         パターンに {} が含まれない場合は pattern をそのまま返す
' ================================================================
Public Function FormatNumberByPattern(ByVal pattern As String, _
                                        ByVal nextNum As Long) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim placeholder As String
    Dim digits As Long
    Dim numStr As String
    Dim i As Long
    
    startPos = InStr(pattern, "{")
    If startPos = 0 Then
        FormatNumberByPattern = pattern
        Exit Function
    End If
    
    endPos = InStr(startPos, pattern, "}")
    If endPos = 0 Then
        FormatNumberByPattern = pattern
        Exit Function
    End If
    
    placeholder = Mid(pattern, startPos + 1, endPos - startPos - 1)
    digits = Len(placeholder)
    numStr = CStr(nextNum)
    
    ' 指定桁数に満たない場合はゼロパディング
    For i = Len(numStr) + 1 To digits
        numStr = "0" & numStr
    Next i
    
    FormatNumberByPattern = Left(pattern, startPos - 1) & numStr & _
                             Mid(pattern, endPos + 1)
End Function

' ================================================================
' 関数名: ParseStanzaLine
' 概要:   スタンザ1行（"Key=Value" 形式）を分解する
' 引数:   line     - 入力行
'         outKey   - 出力: キー
'         outValue - 出力: 値
' 戻り値: Boolean - 解析成功なら True
' 備考:   "=" が含まれない行は False を返す
'         値に "=" が含まれる場合、最初の "=" で分割する
' ================================================================
Public Function ParseStanzaLine(ByVal line As String, _
                                  ByRef outKey As String, _
                                  ByRef outValue As String) As Boolean
    Dim pos As Long
    
    outKey = ""
    outValue = ""
    
    pos = InStr(line, "=")
    If pos = 0 Then
        ParseStanzaLine = False
        Exit Function
    End If
    
    outKey = Trim(Left(line, pos - 1))
    outValue = Mid(line, pos + 1)
    ParseStanzaLine = True
End Function

' ================================================================
' 関数名: TrimAll
' 概要:   前後の空白（半角・全角）を除去
' 引数:   s - 対象文字列
' 戻り値: String - 空白除去後の文字列
' ================================================================
Public Function TrimAll(ByVal s As String) As String
    Dim result As String
    result = s
    
    ' 先頭の全角スペースを除去
    Do While Len(result) > 0 And Left(result, 1) = "　"
        result = Mid(result, 2)
    Loop
    
    ' 末尾の全角スペースを除去
    Do While Len(result) > 0 And Right(result, 1) = "　"
        result = Left(result, Len(result) - 1)
    Loop
    
    ' 半角スペースのトリム
    TrimAll = Trim(result)
End Function

' ================================================================
' 関数名: SafeCellText
' 概要:   "=" 先頭文字列を数式解釈させないようアポストロフィを前置
'         Excel/LibreOffice 両方で有効
' 引数:   s - 対象文字列
' 戻り値: String - 数式化されない文字列
' 備考:   "=" 先頭以外はそのまま返す
'         呼び出し側ではセル代入前に NumberFormat = "@" も併用推奨
' ================================================================
Public Function SafeCellText(ByVal s As String) As String
    If Len(s) > 0 And Left(s, 1) = "=" Then
        SafeCellText = "'" & s
    Else
        SafeCellText = s
    End If
End Function

' ================================================================
' 関数名: IsValidKnowledgeId
' 概要:   ナレッジ番号が "KN-yyyy-NNNN" 形式かを検証する。
'         パストラバーサル / UNC リダイレクト / 予約デバイス名 / 危険文字
'         を MkDir / Kill / Open For Output 等の File system 操作直前で
'         reject するための security ガード (s-2)。
' 引数:   knwNo - 検証対象のナレッジ番号文字列
' 戻り値: Boolean - 形式準拠かつ危険文字を含まない場合のみ True
' 備考:   VBA 標準の RegExp はクロスプラットフォーム差があるため手動判定。
'         12 文字固定 / "KN-" 先頭 / 4 桁年 / "-" 区切り / 4 桁番号 で構成。
'         ".." "\" "/" ":" 等は二重防衛として明示 reject。
' 関連:   ADR 0016
' ================================================================
Public Function IsValidKnowledgeId(ByVal knwNo As String) As Boolean
    On Error GoTo ErrHandler
    IsValidKnowledgeId = False

    If Len(knwNo) <> 12 Then Exit Function
    If Left$(knwNo, 3) <> "KN-" Then Exit Function
    If Mid$(knwNo, 8, 1) <> "-" Then Exit Function

    Dim i As Long
    Dim ch As String
    For i = 4 To 7
        ch = Mid$(knwNo, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    For i = 9 To 12
        ch = Mid$(knwNo, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    ' 危険文字の二重防衛 reject (12 文字制約で本来は通らないが、明示)
    Dim danger As Variant
    danger = Array("..", "\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
    Dim d As Variant
    For Each d In danger
        If InStr(knwNo, CStr(d)) > 0 Then Exit Function
    Next d

    IsValidKnowledgeId = True
    Exit Function

ErrHandler:
    IsValidKnowledgeId = False
End Function

```
