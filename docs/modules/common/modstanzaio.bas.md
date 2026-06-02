---
title: modStanzaIO.bas
description: modStanzaIO.bas のソースコード（コピペ用）
---

# modStanzaIO.bas

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modStanzaIO.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modStanzaIO"
' ================================================================
' モジュール: modStanzaIO（Phase 2 task 2.5 / ユーティリティ層）
' 概要:       スタンザ .txt の read/write primitive（[SECTION]+key=value 汎用形式）
'             ADR-0053 §2.9 外部 I/O 独立化規則順守
'             v2_stanza_parser_spec.md §1 / §2 / §6 準拠
' Version:    v2.2 (2026-05-23, task 2.11: UI スタンザ検証層 UI-SCHEMA-205-209 追加)
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

' --- v2.2 task 2.11: UI スタンザ検証 UI-SCHEMA-205-209 用の定数 ---
Private Const SEV_ERROR As String = "ERROR"
Private Const SEV_WARN As String = "WARN"
Private Const LOGID_INPUT_NAME_MISSING As String = "UI-SCHEMA-205"
Private Const LOGID_INPUT_NAME_DUPLICATE As String = "UI-SCHEMA-206"
Private Const LOGID_GRID_ENDCELL_MISSING As String = "UI-SCHEMA-207"
Private Const LOGID_GRID_COLUMNS_COLON As String = "UI-SCHEMA-208"
Private Const LOGID_GRID_HEADER_MISMATCH As String = "UI-SCHEMA-209"
Private Const SECTION_INPUT As String = "INPUT"
Private Const SECTION_GRID As String = "GRID"
Private Const KEY_NAME As String = "Name"
Private Const KEY_ENDCELL As String = "EndCell"
Private Const KEY_COLUMNS As String = "Columns"
Private Const KEY_COLUMN_HEADERS As String = "ColumnHeaders"


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

' ----------------------------------------------------------------
' Public I/F: UI スタンザ検証（v2.2 task 2.11、UI-SCHEMA-205-209）
' ----------------------------------------------------------------
' 検証層は v2_stanza_parser_spec.md §3 のとおり parse とは別 Function。
' 本 task 2.11 の範囲は UI スタンザ固有検証のうち v2.2 で新設された
' UI-SCHEMA-205-209 の 5 件のみ。共通検証(001-102)・既存 UI 固有検証
' (201-204)・フォーマット/config 固有検証は本 task の範囲外（未実装）。
' 結果は ClsStanzaValidationResult（Errors/Warnings に ClsStanzaValidationIssue）。
' ----------------------------------------------------------------

' ================================================================
' 関数名: ValidateUiStanza
' 概要:   parse 済み UI スタンザのセクション集合に v2.2 追加の UI スタンザ
'         固有検証 UI-SCHEMA-205-209 を実施する。
'         （v2_stanza_parser_spec.md §3.2 v2.2 / migration_plan §0.6.2 task 2.11）
' 引数:   ByVal sections As Collection - ParseStanzaFile/String の戻り値
'         ByVal filePath As String     - エラー報告用ファイルパス（空可）
' 戻り値: ClsStanzaValidationResult - ERROR/WARN を集約した検証結果
' 備考:   検出ルール一覧
'           UI-SCHEMA-205 ERROR [INPUT].Name 不在（キー無し or 空値）
'           UI-SCHEMA-206 ERROR 同一スタンザ内で [INPUT].Name 重複
'           UI-SCHEMA-207 ERROR [GRID].EndCell 不在（キー無し or 空値）
'           UI-SCHEMA-208 ERROR [GRID].Columns に ':' 残存（旧 v2.1 装飾記法）
'           UI-SCHEMA-209 WARN  [GRID].ColumnHeaders と Columns のトークン数不一致
' ================================================================
Public Function ValidateUiStanza( _
    ByVal sections As Collection, _
    ByVal filePath As String _
) As ClsStanzaValidationResult
    On Error GoTo ErrHandler

    Dim result As ClsStanzaValidationResult
    Dim sec As ClsStanzaSection

    Set result = New ClsStanzaValidationResult
    result.Init

    If sections Is Nothing Then
        Set ValidateUiStanza = result
        Exit Function
    End If

    ' UI-SCHEMA-205 / 206: [INPUT].Name の不在・重複
    Call ValidateInputNames(result, sections, filePath)

    ' UI-SCHEMA-207 / 208 / 209: [GRID] 各セクションの検証
    For Each sec In sections
        If sec.SectionName = SECTION_GRID Then
            Call ValidateGridSection(result, sec, filePath)
        End If
    Next sec

    Set ValidateUiStanza = result
    Exit Function

ErrHandler:
    ' 例外時もこれまでに収集した result を返す（result 未生成なら空を返す）
    If result Is Nothing Then
        Set result = New ClsStanzaValidationResult
        result.Init
    End If
    Set ValidateUiStanza = result
End Function

' ================================================================
' 関数名: ValidateInputNames
' 概要:   [INPUT] セクションを走査し UI-SCHEMA-205（Name 不在）と
'         UI-SCHEMA-206（同一スタンザ内 Name 重複）を検出する。
'         Name は入力欄の論理キーで 1 画面（1 スタンザ）内で一意（ギャップ A）。
' 引数:   ByRef outResult As ClsStanzaValidationResult - 検証結果（追記先）
'         ByVal sections As Collection                 - セクション集合
'         ByVal filePath As String                     - ファイルパス
' 戻り値: なし（outResult に issue を追記）
' 備考:   Name 空値も「不在」と同等に 205 扱い（canonical modEntryKnowledge は
'         strict 時 Name 空で Err.Raise 9001。空キーは入力欄を特定できない）。
' ================================================================
Private Sub ValidateInputNames( _
    ByRef outResult As ClsStanzaValidationResult, _
    ByVal sections As Collection, _
    ByVal filePath As String _
)
    Dim sec As ClsStanzaSection
    Dim nameValue As String
    Dim seenNames As Object

    Set seenNames = CreateObject("Scripting.Dictionary")

    For Each sec In sections
        If sec.SectionName = SECTION_INPUT Then
            nameValue = Trim(sec.GetValue(KEY_NAME))
            If (Not sec.HasKey(KEY_NAME)) Or Len(nameValue) = 0 Then
                Call AddUiIssue(outResult, SEV_ERROR, LOGID_INPUT_NAME_MISSING, _
                    filePath, sec.LineNumber, SECTION_INPUT, KEY_NAME, _
                    "[INPUT].Name が不在または空値です（入力欄の論理キー、画面内一意・必須）")
            ElseIf seenNames.Exists(nameValue) Then
                Call AddUiIssue(outResult, SEV_ERROR, LOGID_INPUT_NAME_DUPLICATE, _
                    filePath, sec.LineNumber, SECTION_INPUT, KEY_NAME, _
                    "[INPUT].Name '" & nameValue & "' が同一スタンザ内で重複しています")
            Else
                seenNames.Add nameValue, True
            End If
        End If
    Next sec
End Sub

' ================================================================
' 関数名: ValidateGridSection
' 概要:   [GRID] セクション 1 件に UI-SCHEMA-207/208/209 を検出する。
' 引数:   ByRef outResult As ClsStanzaValidationResult - 検証結果（追記先）
'         ByVal sec As ClsStanzaSection                - [GRID] セクション
'         ByVal filePath As String                     - ファイルパス
' 戻り値: なし（outResult に issue を追記）
' 備考:   208 の ':' 検出は [GRID] 限定。[FORM_FROM_FORMAT]/[GRID_FROM_FORMAT]
'         の Columns は '見出し:幅' 形式で ':' を含むが、本 Sub は呼出元
'         ValidateUiStanza で SectionName=GRID のみに適用されるため対象外。
' ================================================================
Private Sub ValidateGridSection( _
    ByRef outResult As ClsStanzaValidationResult, _
    ByVal sec As ClsStanzaSection, _
    ByVal filePath As String _
)
    Dim columnsValue As String
    Dim headersValue As String
    Dim columnsCount As Long
    Dim headersCount As Long

    ' UI-SCHEMA-207: EndCell 不在（キー無し or 空値）
    If (Not sec.HasKey(KEY_ENDCELL)) Or Len(Trim(sec.GetValue(KEY_ENDCELL))) = 0 Then
        Call AddUiIssue(outResult, SEV_ERROR, LOGID_GRID_ENDCELL_MISSING, _
            filePath, sec.LineNumber, SECTION_GRID, KEY_ENDCELL, _
            "[GRID].EndCell が不在または空値です（データ行容量算出に必須、ギャップ B）")
    End If

    columnsValue = sec.GetValue(KEY_COLUMNS)

    ' UI-SCHEMA-208: Columns に ':' 残存（v2.2 で論理列名のみの CSV と確定）
    If sec.HasKey(KEY_COLUMNS) Then
        If InStr(columnsValue, ":") > 0 Then
            Call AddUiIssue(outResult, SEV_ERROR, LOGID_GRID_COLUMNS_COLON, _
                filePath, sec.LineNumber, SECTION_GRID, KEY_COLUMNS, _
                "[GRID].Columns に ':' が残存（旧装飾記法。装飾は ColumnHeaders へ分離）")
        End If
    End If

    ' UI-SCHEMA-209: ColumnHeaders（任意キー）と Columns のトークン数不一致
    If sec.HasKey(KEY_COLUMN_HEADERS) And sec.HasKey(KEY_COLUMNS) Then
        headersValue = sec.GetValue(KEY_COLUMN_HEADERS)
        columnsCount = CountCsvTokens(columnsValue)
        headersCount = CountCsvTokens(headersValue)
        If columnsCount <> headersCount Then
            Call AddUiIssue(outResult, SEV_WARN, LOGID_GRID_HEADER_MISMATCH, _
                filePath, sec.LineNumber, SECTION_GRID, KEY_COLUMN_HEADERS, _
                "[GRID].ColumnHeaders(" & headersCount & ") と Columns(" & _
                columnsCount & ") のトークン数が不一致です")
        End If
    End If
End Sub

' ================================================================
' 関数名: AddUiIssue
' 概要:   ClsStanzaValidationIssue を生成し検証結果へ追記する内部ヘルパ。
' 引数:   ByRef outResult As ClsStanzaValidationResult - 追記先
'         ByVal severity As String    - SEV_ERROR / SEV_WARN
'         ByVal logId As String       - LogID（例 UI-SCHEMA-205）
'         ByVal filePath As String    - ファイルパス
'         ByVal lineNumber As Long    - セクション開始行番号
'         ByVal sectionName As String - セクション名
'         ByVal keyName As String     - key 名
'         ByVal message As String     - 人間可読メッセージ
' 戻り値: なし
' ================================================================
Private Sub AddUiIssue( _
    ByRef outResult As ClsStanzaValidationResult, _
    ByVal severity As String, _
    ByVal logId As String, _
    ByVal filePath As String, _
    ByVal lineNumber As Long, _
    ByVal sectionName As String, _
    ByVal keyName As String, _
    ByVal message As String _
)
    Dim issue As ClsStanzaValidationIssue
    Set issue = New ClsStanzaValidationIssue
    issue.Init severity, logId, filePath, lineNumber, sectionName, keyName, message
    outResult.AddIssue issue
End Sub

' ================================================================
' 関数名: CountCsvTokens
' 概要:   カンマ区切り文字列のトークン数を返す（空文字列は 0 件）。
' 引数:   ByVal csvValue As String - カンマ区切り文字列
' 戻り値: Long - トークン数
' ================================================================
Private Function CountCsvTokens(ByVal csvValue As String) As Long
    Dim parts As Variant
    If Len(csvValue) = 0 Then
        CountCsvTokens = 0
    Else
        parts = Split(csvValue, ",")
        CountCsvTokens = UBound(parts) - LBound(parts) + 1
    End If
End Function
```
