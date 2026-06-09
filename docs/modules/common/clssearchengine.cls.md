---
title: clsSearchEngine.cls
description: clsSearchEngine.cls のソースコード（コピペ用）
---

# clsSearchEngine.cls

**配置先**: 共通モジュール（3 ブック共通）
**種類**: クラスモジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsSearchEngine.cls`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSearchEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================================================
' クラス: clsSearchEngine（ビジネスロジック層）
' 概要:   ナレッジ検索を担当
'         番号ダイレクト検索、キーワード検索（AND/OR）、日付範囲フィルタ
'         スコアリング (出現回数 + タイトル/フィールドブースト)、
'         結果サムネ表示、詳細画像ペイン表示までを担う。
' 依存先: clsLogger, modFileIO, modStringUtil, modDateUtil, modCommon,
'         modImageRender (新規; サムネ/詳細画像 Shape 描画)
' ================================================================
'
' ====================================================================
' [真版 ChromaDB 連動 切替ポイント]
' 本クラスはモック実装として <dataFolder>/*.txt を ADODB で SJIS 読込し
' スタンザを line-split して match 判定する。真版では事前 ETL で
' "Data" シートに chunks を export しておき ScanAndMatch の txt ループを
' Range 走査に置き換える。VBA 子プロセス禁止 (Shell/Run/Exec/CreateObject
' Exec 系全部禁止) のため、本クラスから外部プロセスを起動しない。
' ====================================================================

' --- 内部固定値 (UI 化されない searchMode/targetField の既定) ---
Private Const DEFAULT_SEARCH_MODE As String = "AND"
Private Const DEFAULT_TARGET_FIELD As String = ""  ' 空 = 全フィールド対象

' --- 結果一覧の表示上限 ---
Private Const RESULT_MAX_ROWS As Long = 200

' --- スコアリングのブースト係数 ---
Private Const SCORE_BOOST_TITLE As Double = 3#
Private Const SCORE_BOOST_TARGET_FIELD As Double = 2#

' --- ImagePath スタンザのキー名 ---
Private Const IMAGE_PATH_KEY As String = "ImagePath"

' --- v2.2 検索エンジン v2 / M-09 表示描画で追加 (task 6.4 / 6.10) ---
Private Const EXCERPT_MAX_LEN As Long = 60
Private Const M09_FORM_CLEAR_RANGE As String = "A8:L39"

Private m_logger As clsLogger
Private m_dataFolder As String

' ================================================================
' 関数名: Init
' 概要:   初期化
' 引数:   logger     - ログ出力用
'         dataFolder - ナレッジデータフォルダ
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logger As clsLogger, ByVal dataFolder As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0617] clsSearchEngine.Init ENTER"  ' [ADR-0100]
    Set m_logger = logger
    m_dataFolder = dataFolder
End Sub

' ================================================================
' 関数名: FindByNumber
' 概要:   指定番号のナレッジが存在するか確認
' 引数:   knowledgeNo - ナレッジ番号
' 戻り値: Boolean - 存在すれば True
' ================================================================
Public Function FindByNumber(ByVal knowledgeNo As String) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0618] clsSearchEngine.FindByNumber ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")

    FindByNumber = FileExists(filePath)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0619] clsSearchEngine.FindByNumber EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0620] clsSearchEngine.FindByNumber EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    FindByNumber = False
End Function

' ================================================================
' 関数名: ExecuteSearch (v2.2 task 6.4)
' 概要:   検索条件 Dictionary を受け取り data\ 配下を全件 scan して
'         一致ナレッジを「行 Dictionary の Collection」で返す。
'         modEntrySearch.Search_Workflow が定める契約
'         (searchExecutor.ExecuteSearch(searchDict) As Collection、
'         stub clsTestSearchExecutor と同一 I/F) の実体実装。
'         v1 の固定セル SS_ROW_* 読み・自己結果書込 PopulateResults は
'         廃止し、scan/スコアリング内部ロジックのみを流用する。
' 引数:   searchDict - 検索条件。キーは OP-5 確定の 8 キー
'                      (keyword/targetFields/formatId/dateFrom/dateTo/
'                      category/assignee/directNo)。
' 戻り値: Collection - 1 行 1 Dictionary。各 Dictionary のキーは
'                      検索結果グリッドの論理列名 (No/KnowledgeNo/
'                      FormatName/CreatedDate/UpdatedDate/Title/
'                      Excerpt/Category/Assignee)。
' ================================================================
Public Function ExecuteSearch(ByVal searchDict As Object) As Collection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0621] clsSearchEngine.ExecuteSearch ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim results As Collection
    Set results = New Collection

    Dim keyword As String, formatId As String
    Dim dateFrom As String, dateTo As String
    Dim category As String, assignee As String, directNo As String
    keyword = GetDictStr(searchDict, "keyword")
    formatId = GetDictStr(searchDict, "formatId")
    dateFrom = GetDictStr(searchDict, "dateFrom")
    dateTo = GetDictStr(searchDict, "dateTo")
    category = GetDictStr(searchDict, "category")
    assignee = GetDictStr(searchDict, "assignee")
    directNo = GetDictStr(searchDict, "directNo")

    ' spec の placeholder 文字列を除去 (v1 SearchByKeywords と同じ前処理)
    keyword = StripPlaceholder(keyword)
    category = StripPlaceholder(category)
    assignee = StripPlaceholder(assignee)
    dateFrom = StripPlaceholder(dateFrom)
    dateTo = StripPlaceholder(dateTo)

    ' フォーマット dropdown の全件選択肢 "すべて" は絞込なし扱い
    If formatId = ALL_FORMATS_TOKEN Then formatId = ""

    ' 番号ダイレクト検索: directNo 指定時はその 1 件のみ返す
    If Trim(directNo) <> "" Then
        If FindByNumber(Trim(directNo)) Then
            results.Add BuildResultRow(1, Trim(directNo))
        End If
        Set ExecuteSearch = results
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0622] clsSearchEngine.ExecuteSearch EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    ' data\ 全件 scan (scan/スコアリング内部ロジックを流用)
    Dim matched() As String
    Dim scores() As Double
    Dim matchCount As Long
    matchCount = ScanAndMatch(formatId, keyword, _
                              DEFAULT_SEARCH_MODE, DEFAULT_TARGET_FIELD, _
                              category, assignee, _
                              dateFrom, dateTo, matched, scores)

    Dim i As Long
    For i = 0 To matchCount - 1
        results.Add BuildResultRow(i + 1, matched(i))
    Next i

    Set ExecuteSearch = results
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0623] clsSearchEngine.ExecuteSearch EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0624] clsSearchEngine.ExecuteSearch EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsSearchEngine", "ExecuteSearch", Err.Description, "", "LOG-M08-SEARCH-ERR"
    End If
    Set ExecuteSearch = New Collection
End Function

' ================================================================
' 関数名: GetDictStr
' 概要:   Dictionary から key の値を文字列で取得 (不在時は空文字列)
' 引数:   d   - 検索条件 Dictionary
'         key - 取得キー
' 戻り値: String - 値 (不在・Nothing なら "")
' ================================================================
Private Function GetDictStr(ByVal d As Object, ByVal key As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0625] clsSearchEngine.GetDictStr ENTER"  ' [ADR-0100]
    If d Is Nothing Then
        GetDictStr = ""
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0626] clsSearchEngine.GetDictStr EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    If d.Exists(key) Then
        GetDictStr = CStr(d(key))
    Else
        GetDictStr = ""
    End If
End Function

' ================================================================
' 関数名: BuildResultRow
' 概要:   1 件のナレッジから検索結果 1 行ぶんの Dictionary を組む。
'         キーは検索結果グリッドの論理列名と一致させる。
' 引数:   rowNo - 1 始まりの連番
'         knwNo - ナレッジ番号
' 戻り値: Object - 行 Dictionary
' ================================================================
Private Function BuildResultRow(ByVal rowNo As Long, ByVal knwNo As String) As Object
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0627] clsSearchEngine.BuildResultRow ENTER"  ' [ADR-0100]
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knwNo & ".txt")
    Dim content As String
    content = ReadShiftJisFile(filePath)

    d("No") = rowNo
    d("KnowledgeNo") = knwNo
    d("FormatName") = ExtractStanzaValue(content, "FormatName")
    d("CreatedDate") = ExtractStanzaValue(content, "CreatedDate")
    d("UpdatedDate") = ExtractStanzaValue(content, "UpdatedDate")
    d("Title") = ExtractKeyValueFromItems(content, ChrW(&H30BF) & ChrW(&H30A4) & ChrW(&H30C8) & ChrW(&H30EB))
    d("Excerpt") = BuildExcerpt(content)
    d("Category") = ExtractKeyValueFromItems(content, ChrW(&H30AB) & ChrW(&H30C6) & ChrW(&H30B4) & ChrW(&H30EA))
    d("Assignee") = ExtractKeyValueFromItems(content, ChrW(&H62C5) & ChrW(&H5F53) & ChrW(&H8005))

    Set BuildResultRow = d
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0628] clsSearchEngine.BuildResultRow EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' 関数名: BuildExcerpt
' 概要:   検索結果グリッドの「事象抜粋」列用に、事象フィールド値の
'         先頭 EXCERPT_MAX_LEN 文字を抜粋する。
' 引数:   content - ナレッジファイル内容
' 戻り値: String - 事象抜粋 (上限超過時は末尾に "...")
' ================================================================
Private Function BuildExcerpt(ByVal content As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0629] clsSearchEngine.BuildExcerpt ENTER"  ' [ADR-0100]
    Dim s As String
    s = ExtractKeyValueFromItems(content, ChrW(&H4E8B) & ChrW(&H8C61))
    If Len(s) > EXCERPT_MAX_LEN Then
        s = Left(s, EXCERPT_MAX_LEN) & "..."
    End If
    BuildExcerpt = s
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0630] clsSearchEngine.BuildExcerpt EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' 関数名: DisplayKnowledge (v2.2 task 6.10 / OP-3 確定)
' 概要:   指定ナレッジを M-09 ナレッジ表示画面へ描画する。
'         v1 の固定セル (KD_*) 直書きは廃止し、M-09 スタンザの
'         [FORM_FROM_FORMAT] (ReadOnly=TRUE 表示モード) を
'         modUILoader.ApplyFormFromFormat 経由で動的描画する経路に
'         一本化した (architecture_v2 §0.6.6.1)。本 Sub はメタ帯セル
'         (PopulateFromKnowledge=ナレッジ番号セル / SourceFormatRef=
'         FormatID セル) へ値を書き、ApplyFormFromFormat へ委譲する。
' 引数:   knowledgeNo - ナレッジ番号
' 戻り値: なし
' ================================================================
Public Sub DisplayKnowledge(ByVal knowledgeNo As String)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0631] clsSearchEngine.DisplayKnowledge ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")
    If Not FileExists(filePath) Then Exit Sub

    Dim content As String
    content = ReadShiftJisFile(filePath)
    If content = "" Then Exit Sub

    ' SourceFormatRef セルへ書く FormatID をナレッジ本文から取得
    Dim formatId As String
    formatId = ExtractStanzaValue(content, "FormatID")

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_DISPLAY)

    ' M-09 スタンザの [FORM_FROM_FORMAT] セクションを取得
    Dim formSec As ClsStanzaSection
    Set formSec = FindFormFromFormatSection()
    If formSec Is Nothing Then
        If Not m_logger Is Nothing Then
            m_logger.LogError "clsSearchEngine", "DisplayKnowledge", _
                "M-09 stanza has no FORM_FROM_FORMAT section", "", "LOG-M09-GOTOEDIT-ENTRY"
        End If
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0632] clsSearchEngine.DisplayKnowledge EXIT-OK"  ' [ADR-0100]
        Exit Sub
    End If

    ' メタ帯セルへ値を流し込む (ApplyFormFromFormat がこのセルから
    ' knowledgeNo / FormatID を読む。schema 3.12)
    Dim knwCell As String, fmtCell As String
    knwCell = LocalCellRef(formSec.GetValue("PopulateFromKnowledge"))
    fmtCell = LocalCellRef(formSec.GetValue("SourceFormatRef"))
    If Len(knwCell) > 0 Then ws.Range(knwCell).Value = knowledgeNo
    If Len(fmtCell) > 0 Then ws.Range(fmtCell).Value = formatId

    ' 旧描画行のクリア (再描画時の余分行除去、StartCell～ボタン行手前)
    On Error Resume Next
    ws.Range(M09_FORM_CLEAR_RANGE).ClearContents
    On Error GoTo ErrHandler

    ' [FORM_FROM_FORMAT] 描画を modUILoader へ委譲 (ReadOnly=TRUE 表示モード)
    modUILoader.ApplyFormFromFormat ws, formSec

    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsSearchEngine", "DisplayKnowledge", logId:="LOG-M09-GOTOEDIT-EXIT-OK", message:= _
                          ChrW(&H8868) & ChrW(&H793A) & ": " & knowledgeNo
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0633] clsSearchEngine.DisplayKnowledge EXIT-OK"  ' [ADR-0100]
    Exit Sub

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0634] clsSearchEngine.DisplayKnowledge EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsSearchEngine", "DisplayKnowledge", Err.Description, "", "LOG-M09-GOTOEDIT-ENTRY"
    End If
End Sub

' ================================================================
' 関数名: FindFormFromFormatSection
' 概要:   M-09 の UI スタンザを読み [FORM_FROM_FORMAT] セクションを返す。
' 戻り値: ClsStanzaSection - 見つからなければ Nothing
' ================================================================
Private Function FindFormFromFormatSection() As ClsStanzaSection
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0635] clsSearchEngine.FindFormFromFormatSection ENTER"  ' [ADR-0100]
    Dim sections As Collection
    Set sections = modUILoader.LoadUiDefinition(UiRoleKensaku(), "M-09")  ' R-3-b: UI-def id は M-NN(ファイル名)。シート名定数は流用しない
    Dim i As Long
    Dim sec As ClsStanzaSection
    For i = 1 To sections.Count
        Set sec = sections.Item(i)
        If sec.SectionName = "FORM_FROM_FORMAT" Then
            Set FindFormFromFormatSection = sec
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0636] clsSearchEngine.FindFormFromFormatSection EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next i
    Set FindFormFromFormatSection = Nothing
End Function

' ================================================================
' 関数名: LocalCellRef
' 概要:   A1 参照からシート修飾 (Sheet!) を除いたローカルセル番地を返す。
' 引数:   cellRef - A1 参照 (シート修飾可)
' 戻り値: String - ローカルセル番地
' ================================================================
Private Function LocalCellRef(ByVal cellRef As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0637] clsSearchEngine.LocalCellRef ENTER"  ' [ADR-0100]
    Dim p As Long
    p = InStr(cellRef, "!")
    If p > 0 Then
        LocalCellRef = Mid(cellRef, p + 1)
    Else
        LocalCellRef = cellRef
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0638] clsSearchEngine.LocalCellRef EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' 関数名: ScoreMatchPublic
' 概要:   ScoreMatch のテスト用 Public ラッパ。
'         BUG-001 修正 (2026-05-12): category / assignee 引数を追加。
'         既存テスト互換のため searchMode/targetField は引き続き受け取る。
' 戻り値: Double - 0=miss、1+=score
' 備考:   テスト層 (T4-006/007) から直接呼び出し可能にするための公開関数。
'         本番フローは ScanAndMatch 経由で内部呼び出しされるため通常は使わない。
' ================================================================
Public Function ScoreMatchPublic(ByVal content As String, _
                                   ByVal formatId As String, _
                                   ByVal keywords As String, _
                                   ByVal searchMode As String, _
                                   ByVal targetField As String, _
                                   ByVal fromDateStr As String, _
                                   ByVal toDateStr As String, _
                                   Optional ByVal category As String = "", _
                                   Optional ByVal assignee As String = "") As Double
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0639] clsSearchEngine.ScoreMatchPublic ENTER"  ' [ADR-0100]
    Dim fromDate As Date
    Dim toDate As Date
    Dim hasFromDate As Boolean
    Dim hasToDate As Boolean
    hasFromDate = TryParseDate(fromDateStr, fromDate)
    hasToDate = TryParseDate(toDateStr, toDate)

    ScoreMatchPublic = ScoreMatch(content, formatId, keywords, searchMode, _
                                    targetField, category, assignee, _
                                    fromDate, hasFromDate, _
                                    toDate, hasToDate)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0640] clsSearchEngine.ScoreMatchPublic EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' 関数名: ExtractImagePathPublic
' 概要:   ExtractImagePath のテスト用 Public ラッパ
' 引数:   content     - ナレッジファイル内容
'         knowledgeNo - ナレッジ番号 (既定パス生成用)
' 戻り値: String - ImagePath スタンザ値 or 既定値 (knowledgeNo & ".png")
' ================================================================
Public Function ExtractImagePathPublic(ByVal content As String, _
                                         ByVal knowledgeNo As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0641] clsSearchEngine.ExtractImagePathPublic ENTER"  ' [ADR-0100]
    ExtractImagePathPublic = ExtractImagePath(content, knowledgeNo)
End Function

' ================================================================
' 関数名: ResolveImageFolderPublic
' 概要:   ResolveImageFolder のテスト用 Public ラッパ
' 引数:   dataFolder - データフォルダのパス
' 戻り値: String - <dataFolder>/../kb_images
' ================================================================
Public Function ResolveImageFolderPublic(ByVal dataFolder As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0642] clsSearchEngine.ResolveImageFolderPublic ENTER"  ' [ADR-0100]
    ResolveImageFolderPublic = ResolveImageFolder(dataFolder)
End Function

' ================================================================
' 関数名: ScanAndMatch
' 概要:   データフォルダ内の全ナレッジファイルをスキャンして条件一致を抽出
'         + 各ファイルにスコアを付け、降順で並び替えた配列を返す
' 引数:   formatId     - フォーマットID (絞込、空なら全)
'         keywords     - キーワード
'         searchMode   - "AND" or "OR"
'         targetField  - 対象フィールド
'         fromDateStr  - 作成日From
'         toDateStr    - 作成日To
'         outMatched   - 出力: マッチしたナレッジ番号の配列 (score 降順)
'         outScores    - 出力: 対応するスコア配列
' 戻り値: Long - マッチ件数
' ================================================================
Private Function ScanAndMatch(ByVal formatId As String, _
                                ByVal keywords As String, _
                                ByVal searchMode As String, _
                                ByVal targetField As String, _
                                ByVal category As String, _
                                ByVal assignee As String, _
                                ByVal fromDateStr As String, _
                                ByVal toDateStr As String, _
                                ByRef outMatched() As String, _
                                ByRef outScores() As Double) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0643] clsSearchEngine.ScanAndMatch ENTER"  ' [ADR-0100]
    Dim files As Variant
    files = ListFilesInFolder(m_dataFolder, "txt")

    ReDim outMatched(0 To RESULT_MAX_ROWS - 1)
    ReDim outScores(0 To RESULT_MAX_ROWS - 1)
    Dim matchCount As Long
    matchCount = 0

    Dim fromDate As Date
    Dim toDate As Date
    Dim hasFromDate As Boolean
    Dim hasToDate As Boolean
    hasFromDate = TryParseDate(fromDateStr, fromDate)
    hasToDate = TryParseDate(toDateStr, toDate)

    Dim i As Long
    For i = LBound(files) To UBound(files)
        Dim fileName As String
        fileName = CStr(files(i))

        Dim filePath As String
        filePath = CombineFilePath(m_dataFolder, fileName)

        Dim content As String
        content = ReadShiftJisFile(filePath)
        If content <> "" Then
            Dim sc As Double
            sc = ScoreMatch(content, formatId, keywords, searchMode, _
                              targetField, category, assignee, _
                              fromDate, hasFromDate, _
                              toDate, hasToDate)
            If sc > 0# Then
                If matchCount < RESULT_MAX_ROWS Then
                    outMatched(matchCount) = Left(fileName, Len(fileName) - 4)
                    outScores(matchCount) = sc
                    matchCount = matchCount + 1
                End If
            End If
        End If
    Next i

    ' スコア降順ソート (バブルソート、件数 200 以下なので十分)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_DEBUG Then Debug.Print "[D-0645] clsSearchEngine.ScanAndMatch STEP-1 pre SortByScoreDesc"  ' [ADR-0100]
    Call SortByScoreDesc(outMatched, outScores, matchCount)

    ScanAndMatch = matchCount
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0644] clsSearchEngine.ScanAndMatch EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' 関数名: SortByScoreDesc
' 概要:   matched / scores の対を score 降順に並び替える (バブルソート)
' 備考:   同点の場合は入力順を維持 (stable)
' ================================================================
Private Sub SortByScoreDesc(ByRef matched() As String, _
                              ByRef scores() As Double, _
                              ByVal n As Long)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0646] clsSearchEngine.SortByScoreDesc ENTER"  ' [ADR-0100]
    Dim i As Long, j As Long
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If scores(j) < scores(j + 1) Then
                Dim ts As Double
                Dim tm As String
                ts = scores(j)
                scores(j) = scores(j + 1)
                scores(j + 1) = ts
                tm = matched(j)
                matched(j) = matched(j + 1)
                matched(j + 1) = tm
            End If
        Next j
    Next i
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0647] clsSearchEngine.SortByScoreDesc EXIT-OK"  ' [ADR-0100]
End Sub

' ================================================================
' 関数名: ScoreMatch
' 概要:   1 つのナレッジファイル内容にスコアを付与
'         0 = フィルタで弾かれた (該当なし)
'         1 以上 = キーワード出現回数 × ブースト係数
' 引数:   IsMatch とほぼ同じ + ブースト計算
' 戻り値: Double - スコア (0 = 該当なし)
' 備考:   既存 IsMatch (Boolean) はテスト互換性のため残してあり、
'         本関数で完全に上位互換できる。
' ================================================================
Private Function ScoreMatch(ByVal content As String, _
                              ByVal formatId As String, _
                              ByVal keywords As String, _
                              ByVal searchMode As String, _
                              ByVal targetField As String, _
                              ByVal category As String, _
                              ByVal assignee As String, _
                              ByVal fromDate As Date, _
                              ByVal hasFromDate As Boolean, _
                              ByVal toDate As Date, _
                              ByVal hasToDate As Boolean) As Double
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0648] clsSearchEngine.ScoreMatch ENTER"  ' [ADR-0100]
    ' フィルタ判定 (既存 IsMatch と同じ)
    If formatId <> "" Then
        Dim actualFmt As String
        actualFmt = ExtractStanzaValue(content, "FormatID")
        If actualFmt <> formatId Then
            ScoreMatch = 0#
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0649] clsSearchEngine.ScoreMatch EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    End If

    ' --- BUG-001 修正: カテゴリ絞込 (部分一致、大文字小文字無視) ---
    If Trim(category) <> "" Then
        Dim categoryActual As String
        categoryActual = ExtractKeyValueFromItems(content, ChrW(&H30AB) & ChrW(&H30C6) & ChrW(&H30B4) & ChrW(&H30EA))
        If InStr(1, LCase(categoryActual), LCase(Trim(category))) = 0 Then
            ScoreMatch = 0#
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0650] clsSearchEngine.ScoreMatch EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    End If

    ' --- BUG-001 修正: 担当者絞込 (部分一致、大文字小文字無視) ---
    If Trim(assignee) <> "" Then
        Dim assigneeActual As String
        assigneeActual = ExtractKeyValueFromItems(content, ChrW(&H62C5) & ChrW(&H5F53) & ChrW(&H8005))
        If InStr(1, LCase(assigneeActual), LCase(Trim(assignee))) = 0 Then
            ScoreMatch = 0#
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0651] clsSearchEngine.ScoreMatch EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    End If

    If hasFromDate Or hasToDate Then
        Dim createdStr As String
        createdStr = ExtractStanzaValue(content, "CreatedDate")
        Dim createdDate As Date
        If TryParseDate(createdStr, createdDate) Then
            If hasFromDate And createdDate < fromDate Then
                ScoreMatch = 0#
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0652] clsSearchEngine.ScoreMatch EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
            If hasToDate And createdDate > toDate Then
                ScoreMatch = 0#
                If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0653] clsSearchEngine.ScoreMatch EXIT-OK"  ' [ADR-0100]
                Exit Function
            End If
        Else
            ' 2026-06-06: date range specified but file CreatedDate is missing/unparseable.
            ' User explicitly filtered by date -> exclude this file from results.
            ScoreMatch = 0#
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0653b] clsSearchEngine.ScoreMatch EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    End If

    ' キーワード判定
    If Trim(keywords) = "" Then
        ' キーワードなし = 全件ヒット (score=1)
        ScoreMatch = 1#
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0654] clsSearchEngine.ScoreMatch EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    ' searchMode 判定 (空 or AND/OR でない場合は AND 扱い)
    Dim isOr As Boolean
    isOr = (UCase(searchMode) = "OR")

    ' AND/OR の bool 判定 (既存ロジック維持) で 0 確定をスキップ
    Dim allText As String
    allText = BuildSearchTarget(content, targetField)
    Dim hitAll As Boolean
    If isOr Then
        hitAll = ContainsAnyKeyword(allText, keywords)
    Else
        hitAll = ContainsAllKeywords(allText, keywords)
    End If
    If Not hitAll Then
        ScoreMatch = 0#
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0655] clsSearchEngine.ScoreMatch EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    ' --- ブースト計算 ---
    ' 全フィールド連結への出現回数
    Dim baseHits As Long
    baseHits = CountKeywordHits(allText, keywords)

    ' タイトルへの出現はブースト
    Dim titleText As String
    titleText = ExtractKeyValueFromItems(content, ChrW(&H30BF) & ChrW(&H30A4) & ChrW(&H30C8) & ChrW(&H30EB))
    Dim titleHits As Long
    titleHits = CountKeywordHits(titleText, keywords)

    ' targetField 指定時は通常ブーストの代わりに targetField ヒットを強調
    Dim fieldBoost As Double
    fieldBoost = 0#
    If targetField <> "" And targetField <> "(" & ChrW(&H5168) & ChrW(&H30D5) & ChrW(&H30A3) & ChrW(&H30FC) & ChrW(&H30EB) & ChrW(&H30C9) & ")" Then
        Dim fieldText As String
        fieldText = ExtractKeyValueFromItems(content, targetField)
        Dim fieldHits As Long
        fieldHits = CountKeywordHits(fieldText, keywords)
        fieldBoost = fieldHits * SCORE_BOOST_TARGET_FIELD
    End If

    ScoreMatch = CDbl(baseHits) + (CDbl(titleHits) * SCORE_BOOST_TITLE) + fieldBoost
End Function

' ================================================================
' 関数名: CountKeywordHits
' 概要:   target にスペース区切り keywords の各語が出現する合計回数を返す
'         (大文字小文字無視)
' ================================================================
Private Function CountKeywordHits(ByVal target As String, _
                                    ByVal keywords As String) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0656] clsSearchEngine.CountKeywordHits ENTER"  ' [ADR-0100]
    If Trim(target) = "" Or Trim(keywords) = "" Then
        CountKeywordHits = 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0657] clsSearchEngine.CountKeywordHits EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    Dim parts() As String
    parts = Split(Trim(keywords), " ")
    Dim total As Long
    total = 0
    Dim targetLower As String
    targetLower = LCase(target)

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim k As String
        k = Trim(parts(i))
        If k <> "" Then
            total = total + CountOccurrences(targetLower, LCase(k))
        End If
    Next i
    CountKeywordHits = total
End Function

' ================================================================
' 関数名: CountOccurrences
' 概要:   needle が haystack に何回出現するかカウント (重複なし)
' ================================================================
Private Function CountOccurrences(ByVal haystack As String, _
                                    ByVal needle As String) As Long
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0658] clsSearchEngine.CountOccurrences ENTER"  ' [ADR-0100]
    If needle = "" Then
        CountOccurrences = 0
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0659] clsSearchEngine.CountOccurrences EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    Dim cnt As Long
    Dim startPos As Long
    startPos = 1
    cnt = 0
    Do While startPos <= Len(haystack)
        Dim pos As Long
        pos = InStr(startPos, haystack, needle)
        If pos = 0 Then Exit Do
        cnt = cnt + 1
        startPos = pos + Len(needle)
    Loop
    CountOccurrences = cnt
End Function

' ================================================================
' 関数名: IsMatch (既存互換)
' 概要:   既存テスト (T4 系) との互換のため Bool 版を残す
'         内部的には ScoreMatch > 0 と等価
' ================================================================
Private Function IsMatch(ByVal content As String, _
                           ByVal formatId As String, _
                           ByVal keywords As String, _
                           ByVal searchMode As String, _
                           ByVal targetField As String, _
                           ByVal fromDate As Date, _
                           ByVal hasFromDate As Boolean, _
                           ByVal toDate As Date, _
                           ByVal hasToDate As Boolean) As Boolean
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0660] clsSearchEngine.IsMatch ENTER"  ' [ADR-0100]
    ' BUG-001 修正: ScoreMatch のシグニチャに category/assignee が追加された
    '              ため、IsMatch は空文字を渡して既存判定と等価動作させる。
    IsMatch = (ScoreMatch(content, formatId, keywords, searchMode, _
                           targetField, "", "", _
                           fromDate, hasFromDate, _
                           toDate, hasToDate) > 0#)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0661] clsSearchEngine.IsMatch EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' 関数名: BuildSearchTarget
' 概要:   検索対象文字列を構築 (対象フィールド指定時は特定のみ)
' ================================================================
Private Function BuildSearchTarget(ByVal content As String, _
                                     ByVal targetField As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0662] clsSearchEngine.BuildSearchTarget ENTER"  ' [ADR-0100]
    Dim lines() As String
    Dim result As String
    Dim i As Long

    lines = Split(content, vbCrLf)

    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "FieldName=") > 0 And _
           InStr(lines(i), "Value=") > 0 Then
            Dim fName As String
            Dim fValue As String
            fName = ExtractKeyFromItem(lines(i), "FieldName")
            fValue = ExtractKeyFromItem(lines(i), "Value")

            If targetField = "" Or _
               targetField = "(" & ChrW(&H5168) & ChrW(&H30D5) & ChrW(&H30A3) & ChrW(&H30FC) & ChrW(&H30EB) & ChrW(&H30C9) & ")" Or _
               targetField = fName Then
                result = result & " " & fValue
            End If
        End If
    Next i

    BuildSearchTarget = result
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0663] clsSearchEngine.BuildSearchTarget EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' 関数名: ExtractKeyFromItem
' 概要:   ITEM 行から指定キーの値を抽出
' ================================================================
Private Function ExtractKeyFromItem(ByVal lineStr As String, _
                                      ByVal keyName As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0664] clsSearchEngine.ExtractKeyFromItem ENTER"  ' [ADR-0100]
    Dim searchKey As String
    Dim startPos As Long
    Dim endPos As Long

    searchKey = keyName & "="
    startPos = InStr(lineStr, searchKey)
    If startPos = 0 Then
        ExtractKeyFromItem = ""
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0665] clsSearchEngine.ExtractKeyFromItem EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If

    startPos = startPos + Len(searchKey)
    endPos = InStr(startPos, lineStr, " / ")
    If endPos = 0 Then
        ExtractKeyFromItem = Mid(lineStr, startPos)
    Else
        ExtractKeyFromItem = Mid(lineStr, startPos, endPos - startPos)
    End If
End Function

' ================================================================
' 関数名: ExtractStanzaValue
' 概要:   スタンザ Key=Value 形式から単純抽出
' ================================================================
Private Function ExtractStanzaValue(ByVal content As String, _
                                      ByVal keyName As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0666] clsSearchEngine.ExtractStanzaValue ENTER"  ' [ADR-0100]
    Dim lines() As String
    Dim i As Long

    lines = Split(content, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), keyName & "=") = 1 Then
            ExtractStanzaValue = Mid(lines(i), Len(keyName) + 2)
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0667] clsSearchEngine.ExtractStanzaValue EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next i
    ExtractStanzaValue = ""
End Function

' ================================================================
' 関数名: ExtractImagePath
' 概要:   ImagePath スタンザを取得。なければ既定値 (knwNo & ".png") を返す
' 引数:   content     - ナレッジファイル内容
'         knowledgeNo - ナレッジ番号 (既定パス生成用)
' 戻り値: String - ImagePath 値 or 既定値
' 備考:   既存ファイル互換: ImagePath 行が無い場合も knwNo.png にフォールバック
' ================================================================
Private Function ExtractImagePath(ByVal content As String, _
                                    ByVal knowledgeNo As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0668] clsSearchEngine.ExtractImagePath ENTER"  ' [ADR-0100]
    Dim v As String
    v = ExtractStanzaValue(content, IMAGE_PATH_KEY)
    If v = "" Then
        ExtractImagePath = knowledgeNo & ".png"
    Else
        ExtractImagePath = v
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0669] clsSearchEngine.ExtractImagePath EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' 関数名: ResolveImageFolder
' 概要:   <dataFolder>/../kb_images/ のフルパスを返す
' 引数:   dataFolder - データフォルダパス
' 戻り値: String - 画像フォルダパス (末尾区切り無し)
' 備考:   path 区切りは Windows と互換 (\ 区切りを保持)
' ================================================================
Private Function ResolveImageFolder(ByVal dataFolder As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0670] clsSearchEngine.ResolveImageFolder ENTER"  ' [ADR-0100]
    Dim base As String
    base = dataFolder
    ' 末尾の \ / を除去
    Do While Len(base) > 0 And (Right(base, 1) = "\" Or Right(base, 1) = "/")
        base = Left(base, Len(base) - 1)
    Loop
    ' 親フォルダ
    Dim parentEnd As Long
    parentEnd = InStrRev(base, "\")
    If parentEnd = 0 Then parentEnd = InStrRev(base, "/")
    If parentEnd = 0 Then
        ResolveImageFolder = base & "\kb_images"
    Else
        ResolveImageFolder = Left(base, parentEnd - 1) & "\kb_images"
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0671] clsSearchEngine.ResolveImageFolder EXIT-OK"  ' [ADR-0100]
End Function

' ================================================================
' 関数名: ResolveImageFullPath
' 概要:   ImagePath (相対 or 絶対) を絶対パスに解決
' 引数:   imagePath  - スタンザの ImagePath 値 or 既定値
'         dataFolder - 解決基準フォルダ (<dataFolder>/../kb_images がベース)
' 戻り値: String - 絶対パス
' 備考:   imagePath が "X:\..." or "\..." なら絶対扱い、そうでなければ
'         <imageFolder>\<imagePath> として解決
' ================================================================
Private Function ResolveImageFullPath(ByVal imagePath As String, _
                                        ByVal dataFolder As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0672] clsSearchEngine.ResolveImageFullPath ENTER"  ' [ADR-0100]
    If Len(imagePath) >= 2 Then
        If Mid(imagePath, 2, 1) = ":" Or Left(imagePath, 2) = "\" Then
            ResolveImageFullPath = imagePath
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0673] clsSearchEngine.ResolveImageFullPath EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    End If
    ResolveImageFullPath = ResolveImageFolder(dataFolder) & "\" & imagePath
End Function

' ================================================================
' 関数名: ExtractKeyValueFromItems
' 概要:   [ITEM] 行から FieldName 指定の値を抽出
' ================================================================
Private Function ExtractKeyValueFromItems(ByVal content As String, _
                                            ByVal fieldName As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0674] clsSearchEngine.ExtractKeyValueFromItems ENTER"  ' [ADR-0100]
    Dim lines() As String
    Dim i As Long

    lines = Split(content, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "FieldName=" & fieldName) > 0 Then
            ExtractKeyValueFromItems = ExtractKeyFromItem(lines(i), "Value")
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0675] clsSearchEngine.ExtractKeyValueFromItems EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next i
    ExtractKeyValueFromItems = ""
End Function

' ================================================================
' 関数名: CombineFilePath
' 概要:   フォルダパスとファイル名を結合
' ================================================================
Private Function CombineFilePath(ByVal folder As String, _
                                   ByVal fileName As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0676] clsSearchEngine.CombineFilePath ENTER"  ' [ADR-0100]
    If Right(folder, 1) = "\" Or Right(folder, 1) = "/" Then
        CombineFilePath = folder & fileName
    Else
        CombineFilePath = folder & "\" & fileName
    End If
End Function

' ================================================================
' 関数名: StripPlaceholder (BUG-001 修正で追加)
' 概要:   検索画面で spec の HintText (例: "(キーワードを入力)") が
'         入力欄セルに残ったままだと engine が誤って検索文字列として
'         扱ってしまうため、placeholder 文字列を空文字に変換する。
' 引数:   raw - セルから読み取った文字列
' 戻り値: String - placeholder と判定した場合は ""、それ以外はそのまま
' 備考:   placeholder 判定条件: 先頭が "(" かつ末尾が ")" (全角丸括弧不要)。
'         ユーザが意図的に "(...)" を検索キーに使うケースは現状想定外。
'         必要になれば spec 側で sentinel 文字を埋め込む等で識別する。
' ================================================================
Private Function StripPlaceholder(ByVal raw As String) As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0677] clsSearchEngine.StripPlaceholder ENTER"  ' [ADR-0100]
    Dim t As String
    t = Trim(raw)
    If Len(t) >= 2 Then
        If Left(t, 1) = "(" And Right(t, 1) = ")" Then
            StripPlaceholder = ""
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0678] clsSearchEngine.StripPlaceholder EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    End If
    StripPlaceholder = raw
End Function

' ADR-0089: CJK ui_seed dir name (U+691C U+7D22)
' VBA Const cannot hold ChrW(); use Private Function instead.
Private Function UiRoleKensaku() As String
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0679] clsSearchEngine.UiRoleKensaku ENTER"  ' [ADR-0100]
    UiRoleKensaku = ChrW(&H691C) & ChrW(&H7D22)
End Function

' ADR-0006/0090/0094 JP literal removal:
Private Property Get ALL_FORMATS_TOKEN() As String
    ALL_FORMATS_TOKEN = ChrW(&H3059) & ChrW(&H3079) & ChrW(&H3066)
End Property
```
