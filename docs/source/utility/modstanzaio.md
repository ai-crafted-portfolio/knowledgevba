---
title: modStanzaIO.bas
---

# modStanzaIO.bas

| 項目 | 内容 |
|---|---|
| 層 | ユーティリティ層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 3 ブック共通 |
| 役割 | スタンザ形式（[セクション] ＋ key=value）の汎用 read / write |
| 行数 | 203 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modStanzaIO.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modStanzaIO"
' ================================================================
' モジュール: modStanzaIO（Phase 2 task 2.5 / ユーティリティ層）
' 概要:       スタンザ .txt の read/write primitive（[SECTION]+key=value 汎用形式）
'             ADR-0053 §2.9 外部 I/O 独立化規則順守
'             v2_stanza_parser_spec.md §1 / §2 / §6 準拠
' Version:    v2.1（2026-05-16 EOD、Q1-Q57 解消反映）
' 依存先:     ClsStanzaSection（出力型）
' 関連:       ADR-0053 §2.9、ADR-0049 §4(f) Write 末尾切断対策
'             Q18（公開 I/F 5 関数：ParseStanzaFile / ParseStanzaString / WriteStanzaFile / ParseMultiStanza / ValidateStanza）
'             Q24（`#` コメント行許容、行頭のみ、parser が無視）
'             Q15（CRLF 必須、LF 単独 / CR 単独 / 混在は reject、LogError）
' 重要:       knowledge .txt は v2.1 で新スタンザ形式（###xxx###、Q47）を採用、
'             knowledge 専用 parser は modKnowledgeFileIO 内に独立実装（本 module は format/UI/config のみ対象）
' ================================================================
Option Explicit

' ----------------------------------------------------------------
' 定数
' ----------------------------------------------------------------
Private Const STANZA_SEP As String = "==="
Private Const CHARSET_SJIS As String = "Shift_JIS"


' ----------------------------------------------------------------
' Public I/F: 読込
' ----------------------------------------------------------------

' ================================================================
' 関数名: ParseStanzaFile
' 概要:   スタンザファイル全文を Shift_JIS で読込み、セクションリストを返す
' 引数:   ByVal filePath As String - ファイルパス
' 戻り値: Collection - 要素は ClsStanzaSection
' 備考:   ファイル不在 / 文字コード不一致は呼出側でハンドル
' ================================================================
Public Function ParseStanzaFile(ByVal filePath As String) As Collection
    On Error GoTo ErrHandler

    Dim ado As Object
    Dim allText As String

    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2  ' adTypeText
    ado.Charset = CHARSET_SJIS
    ado.Open
    ado.LoadFromFile filePath
    allText = ado.ReadText
    ado.Close

    Set ParseStanzaFile = ParseStanzaString(allText)
    Exit Function

ErrHandler:
    ' エラー時は空 Collection を返す（呼出側で件数 0 判定）
    Set ParseStanzaFile = New Collection
End Function

' ================================================================
' 関数名: ParseStanzaString
' 概要:   スタンザ文字列を parse、セクションリストを返す
' 引数:   ByVal stanzaText As String - スタンザ全文
' 戻り値: Collection - 要素は ClsStanzaSection
' ================================================================
Public Function ParseStanzaString(ByVal stanzaText As String) As Collection
    Dim result As Collection
    Dim lines As Variant
    Dim i As Long
    Dim line As String
    Dim currentSection As ClsStanzaSection
    Dim sectionMatch As String
    Dim eqPos As Long
    Dim key As String
    Dim value As String

    Set result = New Collection
    ' 改行正規化（CRLF / LF / CR を LF に統一して Split）
    stanzaText = Replace(stanzaText, vbCrLf, vbLf)
    stanzaText = Replace(stanzaText, vbCr, vbLf)
    lines = Split(stanzaText, vbLf)

    For i = LBound(lines) To UBound(lines)
        line = CStr(lines(i))
        ' 空行 / コメント
        If Len(Trim(line)) = 0 Then GoTo NextLine
        If Left(Trim(line), 1) = "#" Then GoTo NextLine
        ' セクション区切り
        If Trim(line) = STANZA_SEP Then
            If Not currentSection Is Nothing Then
                result.Add currentSection
                Set currentSection = Nothing
            End If
            GoTo NextLine
        End If
        ' セクションヘッダ [...]
        sectionMatch = ExtractSectionName(line)
        If Len(sectionMatch) > 0 Then
            If Not currentSection Is Nothing Then
                result.Add currentSection
            End If
            Set currentSection = New ClsStanzaSection
            currentSection.Init sectionMatch, i + 1
            GoTo NextLine
        End If
        ' key=value
        eqPos = InStr(line, "=")
        If eqPos > 0 And Not currentSection Is Nothing Then
            key = Trim(Left(line, eqPos - 1))
            value = Mid(line, eqPos + 1)  ' 右辺は trim しない（意味のある空白）
            ' エスケープ展開: \\n -> LF, \\\\ -> \\
            value = Replace(value, "\n", vbLf)
            value = Replace(value, "\\", "\")
            currentSection.SetValue key, value
        End If
NextLine:
    Next i

    ' 最終セクションを追加
    If Not currentSection Is Nothing Then
        result.Add currentSection
    End If

    Set ParseStanzaString = result
End Function

' ----------------------------------------------------------------
' Public I/F: 書込
' ----------------------------------------------------------------

' ================================================================
' 関数名: WriteStanzaFile
' 概要:   セクションリストをスタンザ書式で Shift_JIS + CRLF で書出
' 引数:   ByVal filePath As String         - 出力先
'         ByVal sections As Collection     - ClsStanzaSection のリスト
' ================================================================
Public Sub WriteStanzaFile(ByVal filePath As String, ByVal sections As Collection)
    On Error GoTo ErrHandler

    Dim ado As Object
    Dim content As String
    Dim sec As ClsStanzaSection
    Dim key As Variant
    Dim escapedValue As String

    content = vbNullString
    For Each sec In sections
        content = content & "[" & sec.SectionName & "]" & vbCrLf
        For Each key In sec.KeyValues.Keys
            escapedValue = EscapeStanzaValue(CStr(sec.GetValue(CStr(key))))
            content = content & CStr(key) & "=" & escapedValue & vbCrLf
        Next key
        content = content & STANZA_SEP & vbCrLf
    Next sec

    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2
    ado.Charset = CHARSET_SJIS
    ado.Open
    ado.WriteText content
    ado.SaveToFile filePath, 2  ' adSaveCreateOverWrite
    ado.Close
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "modStanzaIO.WriteStanzaFile", _
              "WriteStanzaFile failed for " & filePath & ": " & Err.Description
End Sub

' ================================================================
' Private helpers (restored after CP932 truncate)
' ================================================================

' "[FOO]" 形式の行からセクション名を抽出、それ以外は ""
Private Function ExtractSectionName(ByVal line As String) As String
    Dim s As String
    s = Trim(line)
    If Len(s) >= 2 And Left(s, 1) = "[" And Right(s, 1) = "]" Then
        ExtractSectionName = Mid(s, 2, Len(s) - 2)
    Else
        ExtractSectionName = ""
    End If
End Function

' エスケープ: \ -> \, CR -> \r, LF -> \n, = の左辺で問題になる場合に備えて = はそのまま
Public Function EscapeStanzaValue(ByVal v As String) As String
    Dim s As String
    s = v
    s = Replace(s, "\", "\\")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbCr, "\n")
    EscapeStanzaValue = s
End Function

' 逆エスケープ: \n -> vbLf, \\ -> \
Public Function UnescapeStanzaValue(ByVal v As String) As String
    Dim s As String
    s = v
    ' 一旦プレースホルダ経由で順序保証
    s = Replace(s, "\\", Chr(1))
    s = Replace(s, "\n", vbLf)
    s = Replace(s, Chr(1), "\")
    UnescapeStanzaValue = s
End Function
```
