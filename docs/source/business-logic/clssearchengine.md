---
title: clsSearchEngine.cls
---

# clsSearchEngine.cls

| 項目 | 値 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 役割 | スコアリング検索 / 結果描画 / サムネ画像描画フック |
| 行数 | 839 行 |

## 配置先

VBE で `挿入 > クラスモジュール`、F4 でプロパティ → `(オブジェクト名)` を `clsSearchEngine` に変更してから、コードペインに貼り付けます。

## ソースコード（コピペ可）

下のコードブロック右上にカーソルを当てるとコピーボタンが表示されます。

```vbnet linenums="1"
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

' --- 検索シートのセル位置 ---
Private Const SS_ROW_DIRECT_NO As Long = 3
Private Const SS_COL_DIRECT_NO As Long = 3
Private Const SS_ROW_FMT_ID As Long = 6
Private Const SS_ROW_KEYWORDS As Long = 7
Private Const SS_ROW_MODE As Long = 8
Private Const SS_ROW_TARGET_FIELD As Long = 9
Private Const SS_ROW_DATE_FROM As Long = 10
Private Const SS_ROW_DATE_TO As Long = 11
Private Const SS_COL_CONDITION_VALUE As Long = 3

' --- 検索結果一覧の位置 (rev22 + image_ext rev1) ---
Private Const SS_RESULT_START_ROW As Long = 15
Private Const SS_RESULT_COL_NO As Long = 1
Private Const SS_RESULT_COL_KNW_NO As Long = 2
Private Const SS_RESULT_COL_FMT_NAME As Long = 3
Private Const SS_RESULT_COL_TITLE As Long = 4
Private Const SS_RESULT_COL_CREATED As Long = 5
Private Const SS_RESULT_COL_UPDATED As Long = 6
Private Const SS_RESULT_COL_DETAIL As Long = 7
Private Const SS_RESULT_COL_THUMB As Long = 8     ' NEW: image_ext
Private Const SS_RESULT_COL_SCORE As Long = 9     ' NEW: image_ext

' --- 結果一覧の表示上限 ---
Private Const RESULT_MAX_ROWS As Long = 200

' --- スコアリングのブースト係数 ---
Private Const SCORE_BOOST_TITLE As Double = 3#
Private Const SCORE_BOOST_TARGET_FIELD As Double = 2#

' --- ImagePath スタンザのキー名 ---
Private Const IMAGE_PATH_KEY As String = "ImagePath"

' --- ナレッジ表示シートの位置 ---
Private Const KD_ROW_KNW_NO As Long = 1
Private Const KD_COL_KNW_NO_VAL As Long = 2
Private Const KD_COL_FMT_INFO As Long = 3
Private Const KD_FORM_START_ROW As Long = 4
Private Const KD_COL_FIELD_NO As Long = 1
Private Const KD_COL_FIELD_NAME As Long = 2
Private Const KD_COL_FIELD_VALUE As Long = 3
Private Const KD_COL_FIELD_LINK As Long = 4

' --- 詳細画像ペインの位置 (M-09 ナレッジ表示シート) ---
Private Const KD_DETAIL_IMG_TOP_ROW As Long = 4
Private Const KD_DETAIL_IMG_LEFT_COL As Long = 10   ' J
Private Const KD_DETAIL_IMG_BOT_ROW As Long = 20
Private Const KD_DETAIL_IMG_RIGHT_COL As Long = 14  ' N

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
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")

    FindByNumber = FileExists(filePath)
    Exit Function

ErrHandler:
    FindByNumber = False
End Function

' ================================================================
' 関数名: SearchByKeywords
' 概要:   検索シートの条件を読み取り、キーワード検索を実行
'         結果を検索シートの結果一覧に出力 (スコア降順)
' 引数:   なし
' 戻り値: Long - ヒット件数
' ================================================================
Public Function SearchByKeywords() As Long
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SEARCH)

    Dim formatId As String
    Dim keywords As String
    Dim searchMode As String
    Dim targetField As String
    Dim fromDateStr As String
    Dim toDateStr As String

    formatId = CStr(ws.Cells(SS_ROW_FMT_ID, SS_COL_CONDITION_VALUE).Value)
    keywords = CStr(ws.Cells(SS_ROW_KEYWORDS, SS_COL_CONDITION_VALUE).Value)
    searchMode = CStr(ws.Cells(SS_ROW_MODE, SS_COL_CONDITION_VALUE).Value)
    targetField = CStr(ws.Cells(SS_ROW_TARGET_FIELD, SS_COL_CONDITION_VALUE).Value)
    fromDateStr = CStr(ws.Cells(SS_ROW_DATE_FROM, SS_COL_CONDITION_VALUE).Value)
    toDateStr = CStr(ws.Cells(SS_ROW_DATE_TO, SS_COL_CONDITION_VALUE).Value)

    Dim matched() As String
    Dim scores() As Double
    Dim matchCount As Long
    matchCount = ScanAndMatch(formatId, keywords, searchMode, targetField, _
                                fromDateStr, toDateStr, matched, scores)

    Call ClearResults(ws)
    Call ClearAllThumbs(ws)
    Call PopulateResults(ws, matched, scores, matchCount)

    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsSearchEngine", "SearchByKeywords", _
                          "検索実行: " & CStr(matchCount) & "件ヒット"
    End If

    SearchByKeywords = matchCount
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsSearchEngine", "SearchByKeywords", Err.Description
    End If
    SearchByKeywords = 0
End Function

' ================================================================
' 関数名: DisplayKnowledge
' 概要:   指定ナレッジ番号のファイルを読み込んでナレッジ表示シートに展開
'         + 画像ペイン (J4:N20) に LoadPicture 相当の Shape を貼付
' 引数:   knowledgeNo - ナレッジ番号
' 戻り値: なし
' ================================================================
Public Sub DisplayKnowledge(ByVal knowledgeNo As String)
    On Error GoTo ErrHandler

    Dim filePath As String
    filePath = CombineFilePath(m_dataFolder, knowledgeNo & ".txt")

    If Not FileExists(filePath) Then
        Exit Sub
    End If

    Dim content As String
    content = ReadShiftJisFile(filePath)
    If content = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_KNW_DISPLAY)

    Call ClearDisplaySheet(ws)
    Call ClearDetailImage(ws)
    Call RenderKnowledgeToDisplay(ws, knowledgeNo, content)
    Call RenderDetailImagePane(ws, knowledgeNo, content)

    If Not m_logger Is Nothing Then
        m_logger.LogInfo "clsSearchEngine", "DisplayKnowledge", _
                          "表示: " & knowledgeNo
    End If
    Exit Sub

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsSearchEngine", "DisplayKnowledge", Err.Description
    End If
End Sub

' ================================================================
' 関数名: ScoreMatchPublic
' 概要:   ScoreMatch のテスト用 Public ラッパ。引数は ScoreMatch と同じ。
' 戻り値: Long - 0=miss、1+=score
' 備考:   テスト層 (T4-006/007) から直接呼び出し可能にするための公開関数。
'         本番フローは ScanAndMatch 経由で内部呼び出しされるため通常は使わない。
' ================================================================
Public Function ScoreMatchPublic(ByVal content As String, _
                                   ByVal formatId As String, _
                                   ByVal keywords As String, _
                                   ByVal searchMode As String, _
                                   ByVal targetField As String, _
                                   ByVal fromDateStr As String, _
                                   ByVal toDateStr As String) As Double
    Dim fromDate As Date
    Dim toDate As Date
    Dim hasFromDate As Boolean
    Dim hasToDate As Boolean
    hasFromDate = TryParseDate(fromDateStr, fromDate)
    hasToDate = TryParseDate(toDateStr, toDate)

    ScoreMatchPublic = ScoreMatch(content, formatId, keywords, searchMode, _
                                    targetField, fromDate, hasFromDate, _
                                    toDate, hasToDate)
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
    ExtractImagePathPublic = ExtractImagePath(content, knowledgeNo)
End Function

' ================================================================
' 関数名: ResolveImageFolderPublic
' 概要:   ResolveImageFolder のテスト用 Public ラッパ
' 引数:   dataFolder - データフォルダのパス
' 戻り値: String - <dataFolder>/../kb_images
' ================================================================
Public Function ResolveImageFolderPublic(ByVal dataFolder As String) As String
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
                                ByVal fromDateStr As String, _
                                ByVal toDateStr As String, _
                                ByRef outMatched() As String, _
                                ByRef outScores() As Double) As Long
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
                              targetField, fromDate, hasFromDate, _
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
    Call SortByScoreDesc(outMatched, outScores, matchCount)

    ScanAndMatch = matchCount
End Function

' ================================================================
' 関数名: SortByScoreDesc
' 概要:   matched / scores の対を score 降順に並び替える (バブルソート)
' 備考:   同点の場合は入力順を維持 (stable)
' ================================================================
Private Sub SortByScoreDesc(ByRef matched() As String, _
                              ByRef scores() As Double, _
                              ByVal n As Long)
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
                              ByVal fromDate As Date, _
                              ByVal hasFromDate As Boolean, _
                              ByVal toDate As Date, _
                              ByVal hasToDate As Boolean) As Double
    ' フィルタ判定 (既存 IsMatch と同じ)
    If formatId <> "" Then
        Dim actualFmt As String
        actualFmt = ExtractStanzaValue(content, "FormatID")
        If actualFmt <> formatId Then
            ScoreMatch = 0#
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
                Exit Function
            End If
            If hasToDate And createdDate > toDate Then
                ScoreMatch = 0#
                Exit Function
            End If
        End If
    End If

    ' キーワード判定
    If Trim(keywords) = "" Then
        ' キーワードなし = 全件ヒット (score=1)
        ScoreMatch = 1#
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
        Exit Function
    End If

    ' --- ブースト計算 ---
    ' 全フィールド連結への出現回数
    Dim baseHits As Long
    baseHits = CountKeywordHits(allText, keywords)

    ' タイトルへの出現はブースト
    Dim titleText As String
    titleText = ExtractKeyValueFromItems(content, "タイトル")
    Dim titleHits As Long
    titleHits = CountKeywordHits(titleText, keywords)

    ' targetField 指定時は通常ブーストの代わりに targetField ヒットを強調
    Dim fieldBoost As Double
    fieldBoost = 0#
    If targetField <> "" And targetField <> "(全フィールド)" Then
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
    If Trim(target) = "" Or Trim(keywords) = "" Then
        CountKeywordHits = 0
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
    If needle = "" Then
        CountOccurrences = 0
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
    IsMatch = (ScoreMatch(content, formatId, keywords, searchMode, _
                           targetField, fromDate, hasFromDate, _
                           toDate, hasToDate) > 0#)
End Function

' ================================================================
' 関数名: BuildSearchTarget
' 概要:   検索対象文字列を構築 (対象フィールド指定時は特定のみ)
' ================================================================
Private Function BuildSearchTarget(ByVal content As String, _
                                     ByVal targetField As String) As String
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
               targetField = "(全フィールド)" Or _
               targetField = fName Then
                result = result & " " & fValue
            End If
        End If
    Next i

    BuildSearchTarget = result
End Function

' ================================================================
' 関数名: ExtractKeyFromItem
' 概要:   ITEM 行から指定キーの値を抽出
' ================================================================
Private Function ExtractKeyFromItem(ByVal line As String, _
                                      ByVal keyName As String) As String
    Dim searchKey As String
    Dim startPos As Long
    Dim endPos As Long

    searchKey = keyName & "="
    startPos = InStr(line, searchKey)
    If startPos = 0 Then
        ExtractKeyFromItem = ""
        Exit Function
    End If

    startPos = startPos + Len(searchKey)
    endPos = InStr(startPos, line, " / ")
    If endPos = 0 Then
        ExtractKeyFromItem = Mid(line, startPos)
    Else
        ExtractKeyFromItem = Mid(line, startPos, endPos - startPos)
    End If
End Function

' ================================================================
' 関数名: ExtractStanzaValue
' 概要:   スタンザ Key=Value 形式から単純抽出
' ================================================================
Private Function ExtractStanzaValue(ByVal content As String, _
                                      ByVal keyName As String) As String
    Dim lines() As String
    Dim i As Long

    lines = Split(content, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), keyName & "=") = 1 Then
            ExtractStanzaValue = Mid(lines(i), Len(keyName) + 2)
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
    Dim v As String
    v = ExtractStanzaValue(content, IMAGE_PATH_KEY)
    If v = "" Then
        ExtractImagePath = knowledgeNo & ".png"
    Else
        ExtractImagePath = v
    End If
End Function

' ================================================================
' 関数名: ResolveImageFolder
' 概要:   <dataFolder>/../kb_images/ のフルパスを返す
' 引数:   dataFolder - データフォルダパス
' 戻り値: String - 画像フォルダパス (末尾区切り無し)
' 備考:   path 区切りは Windows と互換 (\ 区切りを保持)
' ================================================================
Private Function ResolveImageFolder(ByVal dataFolder As String) As String
    Dim base As String
    base = dataFolder
    ' 末尾の \ / を除去
    Do While Len(base) > 0 And (Right(base, 1) = "" Or Right(base, 1) = "/")
        base = Left(base, Len(base) - 1)
    Loop
    ' 親フォルダ
    Dim parentEnd As Long
    parentEnd = InStrRev(base, "")
    If parentEnd = 0 Then parentEnd = InStrRev(base, "/")
    If parentEnd = 0 Then
        ResolveImageFolder = base & "\kb_images"
    Else
        ResolveImageFolder = Left(base, parentEnd - 1) & "\kb_images"
    End If
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
    If Len(imagePath) >= 2 Then
        If Mid(imagePath, 2, 1) = ":" Or Left(imagePath, 2) = "\" Then
            ResolveImageFullPath = imagePath
            Exit Function
        End If
    End If
    ResolveImageFullPath = ResolveImageFolder(dataFolder) & "" & imagePath
End Function

' ================================================================
' 関数名: ClearResults
' 概要:   検索結果一覧の値部分のみクリア (結合セル対策、行ごと Resume Next)
' ================================================================
Private Sub ClearResults(ByVal ws As Worksheet)
    Dim i As Long
    For i = SS_RESULT_START_ROW To SS_RESULT_START_ROW + RESULT_MAX_ROWS
        On Error Resume Next
        ws.Range(ws.Cells(i, 1), ws.Cells(i, SS_RESULT_COL_SCORE)).ClearContents
        Err.Clear
        On Error GoTo 0
    Next i
End Sub

' ================================================================
' 関数名: ClearAllThumbs
' 概要:   検索シート上の旧サムネ Shape (kbThumb_*) を全削除
' 備考:   modImageRender.ClearShapesByPrefix を呼ぶ
' ================================================================
Private Sub ClearAllThumbs(ByVal ws As Worksheet)
    Call ClearShapesByPrefix(ws, "kbThumb_")
End Sub

' ================================================================
' 関数名: ClearDetailImage
' 概要:   ナレッジ表示シート上の旧詳細画像 Shape (kbDetailImg_*) を全削除
' ================================================================
Private Sub ClearDetailImage(ByVal ws As Worksheet)
    Call ClearShapesByPrefix(ws, "kbDetailImg_")
End Sub

' ================================================================
' 関数名: PopulateResults
' 概要:   検索結果を 9 列で書き出す。サムネは modImageRender 経由で
'         Shape として配置。
' 引数:   ws       - 検索シート
'         matched  - 番号配列 (score 降順)
'         scores   - 対応スコア配列
'         count    - 件数
' ================================================================
Private Sub PopulateResults(ByVal ws As Worksheet, _
                              ByRef matched() As String, _
                              ByRef scores() As Double, _
                              ByVal count As Long)
    Dim i As Long
    For i = 0 To count - 1
        Dim knwNo As String
        knwNo = matched(i)

        Dim filePath As String
        filePath = CombineFilePath(m_dataFolder, knwNo & ".txt")

        Dim content As String
        content = ReadShiftJisFile(filePath)

        Dim targetRow As Long
        targetRow = SS_RESULT_START_ROW + i

        ws.Cells(targetRow, SS_RESULT_COL_NO).Value = i + 1
        ws.Cells(targetRow, SS_RESULT_COL_KNW_NO).Value = knwNo
        ws.Cells(targetRow, SS_RESULT_COL_FMT_NAME).Value = _
            ExtractStanzaValue(content, "FormatName")
        ws.Cells(targetRow, SS_RESULT_COL_TITLE).Value = _
            ExtractKeyValueFromItems(content, "タイトル")
        ws.Cells(targetRow, SS_RESULT_COL_CREATED).Value = _
            ExtractStanzaValue(content, "CreatedDate")
        ws.Cells(targetRow, SS_RESULT_COL_UPDATED).Value = _
            ExtractStanzaValue(content, "UpdatedDate")
        ' "▶ 詳細" は CP932 で ▶ が encode 失敗するため ChrW で動的構築する
        ws.Cells(targetRow, SS_RESULT_COL_DETAIL).Value = ChrW(&H25B6) & " 詳細"

        ' --- サムネ Shape 配置 (列 H) ---
        Dim imgRel As String
        imgRel = ExtractImagePath(content, knwNo)
        Dim imgFull As String
        imgFull = ResolveImageFullPath(imgRel, m_dataFolder)
        Call RenderThumb(ws, targetRow, SS_RESULT_COL_THUMB, imgFull, knwNo)

        ' --- スコア (列 I) ---
        ws.Cells(targetRow, SS_RESULT_COL_SCORE).Value = _
            Format(scores(i), "0.00")
    Next i
End Sub

' ================================================================
' 関数名: ExtractKeyValueFromItems
' 概要:   [ITEM] 行から FieldName 指定の値を抽出
' ================================================================
Private Function ExtractKeyValueFromItems(ByVal content As String, _
                                            ByVal fieldName As String) As String
    Dim lines() As String
    Dim i As Long

    lines = Split(content, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "FieldName=" & fieldName) > 0 Then
            ExtractKeyValueFromItems = ExtractKeyFromItem(lines(i), "Value")
            Exit Function
        End If
    Next i
    ExtractKeyValueFromItems = ""
End Function

' ================================================================
' 関数名: ClearDisplaySheet
' 概要:   ナレッジ表示シートの値部分をクリア (Shape は別関数でクリア)
' ================================================================
Private Sub ClearDisplaySheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Cells.ClearContents
    On Error GoTo 0
End Sub

' ================================================================
' 関数名: RenderKnowledgeToDisplay
' 概要:   ナレッジ表示シートに ITEM 行を展開
' ================================================================
Private Sub RenderKnowledgeToDisplay(ByVal ws As Worksheet, _
                                       ByVal knowledgeNo As String, _
                                       ByVal content As String)
    Dim formatId As String
    Dim formatName As String
    formatId = ExtractStanzaValue(content, "FormatID")
    formatName = ExtractStanzaValue(content, "FormatName")

    ws.Cells(KD_ROW_KNW_NO, 1).Value = "ナレッジ番号:"
    ws.Cells(KD_ROW_KNW_NO, KD_COL_KNW_NO_VAL).Value = knowledgeNo
    ws.Cells(KD_ROW_KNW_NO, KD_COL_FMT_INFO).Value = _
        formatName & " (" & formatId & ")"

    Dim lines() As String
    lines = Split(content, vbCrLf)

    Dim targetRow As Long
    targetRow = KD_FORM_START_ROW

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), "FieldNo=") > 0 And _
           InStr(lines(i), "FieldName=") > 0 Then
            ws.Cells(targetRow, KD_COL_FIELD_NO).Value = _
                ExtractKeyFromItem(lines(i), "FieldNo")
            ws.Cells(targetRow, KD_COL_FIELD_NAME).Value = _
                ExtractKeyFromItem(lines(i), "FieldName")
            ws.Cells(targetRow, KD_COL_FIELD_VALUE).NumberFormat = "@"
            ws.Cells(targetRow, KD_COL_FIELD_VALUE).Value = _
                SafeCellText(ExtractKeyFromItem(lines(i), "Value"))
            targetRow = targetRow + 1
        End If
    Next i
End Sub

' ================================================================
' 関数名: RenderDetailImagePane
' 概要:   ナレッジ表示シートの J4:N20 領域に詳細画像を貼付
'         画像が無い場合は薄字で「[画像未配置: ...]」を表示
' ================================================================
Private Sub RenderDetailImagePane(ByVal ws As Worksheet, _
                                    ByVal knowledgeNo As String, _
                                    ByVal content As String)
    On Error Resume Next
    Dim imgRel As String
    imgRel = ExtractImagePath(content, knowledgeNo)
    Dim imgFull As String
    imgFull = ResolveImageFullPath(imgRel, m_dataFolder)

    Call RenderDetailImage(ws, _
                            KD_DETAIL_IMG_TOP_ROW, KD_DETAIL_IMG_LEFT_COL, _
                            KD_DETAIL_IMG_BOT_ROW, KD_DETAIL_IMG_RIGHT_COL, _
                            imgFull, knowledgeNo)
End Sub

' ================================================================
' 関数名: CombineFilePath
' 概要:   フォルダパスとファイル名を結合
' ================================================================
Private Function CombineFilePath(ByVal folder As String, _
                                   ByVal fileName As String) As String
    If Right(folder, 1) = "" Or Right(folder, 1) = "/" Then
        CombineFilePath = folder & fileName
    Else
        CombineFilePath = folder & "" & fileName
    End If
End Function
```
