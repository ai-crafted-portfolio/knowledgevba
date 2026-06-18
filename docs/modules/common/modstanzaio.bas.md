---
title: modStanzaIO.bas
description: modStanzaIO.bas のソースコード（コピペ用）
---

# modStanzaIO.bas

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: 標準モジュール
**更新日**: 2026-06-05 01:27 JST

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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1384] modStanzaIO.ParseStanzaFile ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1385] modStanzaIO.ParseStanzaFile EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1386] modStanzaIO.ParseStanzaFile EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1387] modStanzaIO.ParseStanzaString ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1388] modStanzaIO.ParseStanzaString EXIT-OK"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1389] modStanzaIO.WriteStanzaFile ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1390] modStanzaIO.WriteStanzaFile EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1391] modStanzaIO.WriteStanzaFile EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    Err.Raise Err.Number, "modStanzaIO.WriteStanzaFile", _
              "WriteStanzaFile failed for " & filePath & ": " & Err.Description
End Sub

' ================================================================
' Private helpers (restored after CP932 truncate)
' ================================================================

' "[FOO]" 形式の行からセクション名を抽出、それ以外は ""
Private Function ExtractSectionName(ByVal line As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1392] modStanzaIO.ExtractSectionName ENTER"  ' [ADR-0100]
    Dim s As String
    s = Trim(line)
    If Len(s) >= 2 And Left(s, 1) = "[" And Right(s, 1) = "]" Then
        ExtractSectionName = Mid(s, 2, Len(s) - 2)
    Else
        ExtractSectionName = ""
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1393] modStanzaIO.ExtractSectionName EXIT-OK"  ' [ADR-0100]
End Function

' エスケープ: \ -> \, CR -> \r, LF -> \n, = の左辺で問題になる場合に備えて = はそのまま
Public Function EscapeStanzaValue(ByVal v As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1394] modStanzaIO.EscapeStanzaValue ENTER"  ' [ADR-0100]
    Dim s As String
    s = v
    s = Replace(s, "\", "\\")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbCr, "\n")
    EscapeStanzaValue = s
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1395] modStanzaIO.EscapeStanzaValue EXIT-OK"  ' [ADR-0100]
End Function

' 逆エスケープ: \n -> vbLf, \\ -> \
Public Function UnescapeStanzaValue(ByVal v As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1396] modStanzaIO.UnescapeStanzaValue ENTER"  ' [ADR-0100]
    Dim s As String
    s = v
    ' 一旦プレースホルダ経由で順序保証
    s = Replace(s, "\\", Chr(1))
    s = Replace(s, "\n", vbLf)
    s = Replace(s, Chr(1), "\")
    UnescapeStanzaValue = s
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1397] modStanzaIO.UnescapeStanzaValue EXIT-OK"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1398] modStanzaIO.ValidateUiStanza ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim result As ClsStanzaValidationResult
    Dim sec As ClsStanzaSection

    Set result = New ClsStanzaValidationResult
    result.Init

    If sections Is Nothing Then
        Set ValidateUiStanza = result
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1399] modStanzaIO.ValidateUiStanza EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    ' UI-SCHEMA-205 / 206: [INPUT].Name の不在・重複
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1402] modStanzaIO.ValidateUiStanza STEP-1 pre ValidateInputNames"  ' [ADR-0100]
    Call ValidateInputNames(result, sections, filePath)

    ' UI-SCHEMA-207 / 208 / 209: [GRID] 各セクションの検証
    For Each sec In sections
        If sec.SectionName = SECTION_GRID Then
            If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1403] modStanzaIO.ValidateUiStanza STEP-2 pre ValidateGridSection"  ' [ADR-0100]
            Call ValidateGridSection(result, sec, filePath)
        End If
    Next sec

    Set ValidateUiStanza = result
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1400] modStanzaIO.ValidateUiStanza EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-1401] modStanzaIO.ValidateUiStanza EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1404] modStanzaIO.ValidateInputNames ENTER"  ' [ADR-0100]
    Dim sec As ClsStanzaSection
    Dim nameValue As String
    Dim seenNames As Object

    Set seenNames = CreateObject("Scripting.Dictionary")

    For Each sec In sections
        If sec.SectionName = SECTION_INPUT Then
            nameValue = Trim(sec.GetValue(KEY_NAME))
            If (Not sec.HasKey(KEY_NAME)) Or Len(nameValue) = 0 Then
                If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1406] modStanzaIO.ValidateInputNames STEP-1 pre AddUiIssue"  ' [ADR-0100]
                Call AddUiIssue(outResult, SEV_ERROR, LOGID_INPUT_NAME_MISSING, _
                    filePath, sec.LineNumber, SECTION_INPUT, KEY_NAME, _
                    "[INPUT].Name " & ChrW(&H304C) & ChrW(&H4E0D) & ChrW(&H5728) & ChrW(&H307E) & ChrW(&H305F) & ChrW(&H306F) & ChrW(&H7A7A) & ChrW(&H5024) & ChrW(&H3067) & ChrW(&H3059) & ChrW(&HFF08) & ChrW(&H5165) & ChrW(&H529B) & ChrW(&H6B04) & ChrW(&H306E) & ChrW(&H8AD6) & ChrW(&H7406) & ChrW(&H30AD) & ChrW(&H30FC) & ChrW(&H3001) & ChrW(&H753B) & ChrW(&H9762) & ChrW(&H5185) & ChrW(&H4E00) & ChrW(&H610F) & ChrW(&H30FB) & ChrW(&H5FC5) & ChrW(&H9808) & ChrW(&HFF09))
            ElseIf seenNames.Exists(nameValue) Then
                If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1407] modStanzaIO.ValidateInputNames STEP-2 pre AddUiIssue"  ' [ADR-0100]
                Call AddUiIssue(outResult, SEV_ERROR, LOGID_INPUT_NAME_DUPLICATE, _
                    filePath, sec.LineNumber, SECTION_INPUT, KEY_NAME, _
                    "[INPUT].Name '" & nameValue & "' " & ChrW(&H304C) & ChrW(&H540C) & ChrW(&H4E00) & ChrW(&H30B9) & ChrW(&H30BF) & ChrW(&H30F3) & ChrW(&H30B6) & ChrW(&H5185) & ChrW(&H3067) & ChrW(&H91CD) & ChrW(&H8907) & ChrW(&H3057) & ChrW(&H3066) & ChrW(&H3044) & ChrW(&H307E) & ChrW(&H3059))
            Else
                seenNames.Add nameValue, True
            End If
        End If
    Next sec
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1405] modStanzaIO.ValidateInputNames EXIT-OK"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1408] modStanzaIO.ValidateGridSection ENTER"  ' [ADR-0100]
    Dim columnsValue As String
    Dim headersValue As String
    Dim columnsCount As Long
    Dim headersCount As Long

    ' UI-SCHEMA-207: EndCell 不在（キー無し or 空値）
    If (Not sec.HasKey(KEY_ENDCELL)) Or Len(Trim(sec.GetValue(KEY_ENDCELL))) = 0 Then
        If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1410] modStanzaIO.ValidateGridSection STEP-1 pre AddUiIssue"  ' [ADR-0100]
        Call AddUiIssue(outResult, SEV_ERROR, LOGID_GRID_ENDCELL_MISSING, _
            filePath, sec.LineNumber, SECTION_GRID, KEY_ENDCELL, _
            "[GRID].EndCell " & ChrW(&H304C) & ChrW(&H4E0D) & ChrW(&H5728) & ChrW(&H307E) & ChrW(&H305F) & ChrW(&H306F) & ChrW(&H7A7A) & ChrW(&H5024) & ChrW(&H3067) & ChrW(&H3059) & ChrW(&HFF08) & ChrW(&H30C7) & ChrW(&H30FC) & ChrW(&H30BF) & ChrW(&H884C) & ChrW(&H5BB9) & ChrW(&H91CF) & ChrW(&H7B97) & ChrW(&H51FA) & ChrW(&H306B) & ChrW(&H5FC5) & ChrW(&H9808) & ChrW(&H3001) & ChrW(&H30AE) & ChrW(&H30E3) & ChrW(&H30C3) & ChrW(&H30D7) & " B" & ChrW(&HFF09))
    End If

    columnsValue = sec.GetValue(KEY_COLUMNS)

    ' UI-SCHEMA-208: Columns に ':' 残存（v2.2 で論理列名のみの CSV と確定）
    If sec.HasKey(KEY_COLUMNS) Then
        If InStr(columnsValue, ":") > 0 Then
            If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1411] modStanzaIO.ValidateGridSection STEP-2 pre AddUiIssue"  ' [ADR-0100]
            Call AddUiIssue(outResult, SEV_ERROR, LOGID_GRID_COLUMNS_COLON, _
                filePath, sec.LineNumber, SECTION_GRID, KEY_COLUMNS, _
                "[GRID].Columns " & ChrW(&H306B) & " ':' " & ChrW(&H304C) & ChrW(&H6B8B) & ChrW(&H5B58) & ChrW(&HFF08) & ChrW(&H65E7) & ChrW(&H88C5) & ChrW(&H98FE) & ChrW(&H8A18) & ChrW(&H6CD5) & ChrW(&H3002) & ChrW(&H88C5) & ChrW(&H98FE) & ChrW(&H306F) & " ColumnHeaders " & ChrW(&H3078) & ChrW(&H5206) & ChrW(&H96E2) & ChrW(&HFF09))
        End If
    End If

    ' UI-SCHEMA-209: ColumnHeaders（任意キー）と Columns のトークン数不一致
    If sec.HasKey(KEY_COLUMN_HEADERS) And sec.HasKey(KEY_COLUMNS) Then
        headersValue = sec.GetValue(KEY_COLUMN_HEADERS)
        columnsCount = CountCsvTokens(columnsValue)
        headersCount = CountCsvTokens(headersValue)
        If columnsCount <> headersCount Then
            If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-1412] modStanzaIO.ValidateGridSection STEP-3 pre AddUiIssue"  ' [ADR-0100]
            Call AddUiIssue(outResult, SEV_WARN, LOGID_GRID_HEADER_MISMATCH, _
                filePath, sec.LineNumber, SECTION_GRID, KEY_COLUMN_HEADERS, _
                "[GRID].ColumnHeaders(" & headersCount & ") " & ChrW(&H3068) & " Columns(" & _
                columnsCount & ") " & ChrW(&H306E) & ChrW(&H30C8) & ChrW(&H30FC) & ChrW(&H30AF) & ChrW(&H30F3) & ChrW(&H6570) & ChrW(&H304C) & ChrW(&H4E0D) & ChrW(&H4E00) & ChrW(&H81F4) & ChrW(&H3067) & ChrW(&H3059))
        End If
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1409] modStanzaIO.ValidateGridSection EXIT-OK"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1413] modStanzaIO.AddUiIssue ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1414] modStanzaIO.CountCsvTokens ENTER"  ' [ADR-0100]
    Dim parts As Variant
    If Len(csvValue) = 0 Then
        CountCsvTokens = 0
    Else
        parts = Split(csvValue, ",")
        CountCsvTokens = UBound(parts) - LBound(parts) + 1
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-1415] modStanzaIO.CountCsvTokens EXIT-OK"  ' [ADR-0100]
End Function
```
