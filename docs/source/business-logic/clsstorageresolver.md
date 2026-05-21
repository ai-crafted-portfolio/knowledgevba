---
title: clsStorageResolver.cls
---

# clsStorageResolver.cls

| 項目 | 内容 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | クラスモジュール (.cls) |
| 配置ブック | 3 ブック共通 |
| 役割 | 格納先設定に基づきファイル参照リンクを解決する |
| 行数 | 171 行 |

## 取り込み先

クラスモジュール（.cls）です。下記コードをコピーし、`clsStorageResolver.cls` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールのファイル形式の一部なので、削らずにそのまま保存してください。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStorageResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Phase 6 レビュー: GetStorageConfig は O(N) だが N ? MAX_STORAGE_SCAN_ROWS=1000。
' 現状想定 (docType 数十件) では十分。指摘なし。

' ================================================================
' クラス: clsStorageResolver（ビジネスロジック層）
' 概要:   格納先設定に基づき、ファイル参照フィールドからリンクを生成
'         ドキュメント種類ごとに「共有フォルダ/BOX」「ファイル指定/フォルダ指定」を
'         解釈してクリック可能なパスを組み立てる
' 依存先: clsLogger, modCommon
' ================================================================

' 格納先設定シートの列番号
Private Const STORAGE_COL_NO As Long = 1
Private Const STORAGE_COL_DOC_TYPE As Long = 2
Private Const STORAGE_COL_STORAGE_TYPE As Long = 3
Private Const STORAGE_COL_PATH_TYPE As Long = 4
Private Const STORAGE_COL_BASE_PATH As Long = 5
Private Const STORAGE_COL_NOTE As Long = 6

' 格納先設定の開始行（ヘッダーの次）
Private Const STORAGE_START_ROW As Long = 2

' 格納種別
Private Const MAX_STORAGE_SCAN_ROWS As Long = 1000  ' m-11: マジック数 Const 化
Private Const STORAGE_TYPE_SHARED As String = "共有フォルダ"
Private Const STORAGE_TYPE_BOX As String = "BOX"

' パス種別
Private Const PATH_TYPE_FILE As String = "ファイル指定"
Private Const PATH_TYPE_FOLDER As String = "フォルダ指定"

Private m_logger As clsLogger

' ================================================================
' 関数名: Init
' 概要:   初期化（ロガー参照を保持）
' 引数:   logger - ログ出力用
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logger As clsLogger)
    Set m_logger = logger
End Sub

' ================================================================
' 関数名: ResolveLink
' 概要:   ドキュメント種類とファイル名からクリック可能なリンク文字列を生成
' 引数:   docType  - ドキュメント種類（例: "障害報告書"）
'         fileName - ファイル名（例: "report.xlsx"）
' 戻り値: String - 生成されたリンク文字列
'                  （例: "\\filesvr\incident\reports\report.xlsx"）
' 備考:   設定が見つからない/パス種別がフォルダの場合はベースパスのみ返す
'         BOXの場合はBOX URL
' ================================================================
Public Function ResolveLink(ByVal docType As String, _
                              ByVal fileName As String) As String
    On Error GoTo ErrHandler
    
    Dim storageType As String
    Dim pathType As String
    Dim basePath As String
    
    If Not GetStorageConfig(docType, storageType, pathType, basePath) Then
        If Not m_logger Is Nothing Then
            m_logger.LogWarn "clsStorageResolver", "ResolveLink", _
                              "格納先設定なし: docType=" & docType, , LOG_STORAGE_OPEN_ENTRY
        End If
        ResolveLink = ""
        Exit Function
    End If
    
    ' BOXの場合はBOX URLをそのまま返す（ファイル直指定不可）
    If storageType = STORAGE_TYPE_BOX Then
        ResolveLink = basePath
        Exit Function
    End If
    
    ' 共有フォルダの場合
    If pathType = PATH_TYPE_FILE Then
        ' ファイル指定: basePath + fileName
        ResolveLink = CombinePath(basePath, fileName)
    Else
        ' フォルダ指定: basePath のみ
        ResolveLink = basePath
    End If
    Exit Function

ErrHandler:
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsStorageResolver", "ResolveLink", _
                           "リンク生成失敗: " & Err.Description, , LOG_STORAGE_OPEN_EXIT_OK
    End If
    ResolveLink = ""
End Function

' ================================================================
' 関数名: GetStorageConfig
' 概要:   格納先設定シートから指定ドキュメント種類の設定を取得
' 引数:   docType        - ドキュメント種類
'         outStorageType - 出力: 格納種別
'         outPathType    - 出力: パス種別
'         outBasePath    - 出力: ベースパス
' 戻り値: Boolean - 設定が見つかったなら True
' 備考:   ドキュメント種類名の完全一致で検索
' ================================================================
Private Function GetStorageConfig(ByVal docType As String, _
                                    ByRef outStorageType As String, _
                                    ByRef outPathType As String, _
                                    ByRef outBasePath As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_STORAGE)
    
    Dim i As Long
    Dim maxRow As Long
    maxRow = MAX_STORAGE_SCAN_ROWS
    
    For i = STORAGE_START_ROW To maxRow
        Dim currentType As String
        currentType = CStr(ws.Cells(i, STORAGE_COL_DOC_TYPE).Value)
        
        If currentType = "" Then
            Exit For
        End If
        
        If currentType = docType Then
            outStorageType = CStr(ws.Cells(i, STORAGE_COL_STORAGE_TYPE).Value)
            outPathType = CStr(ws.Cells(i, STORAGE_COL_PATH_TYPE).Value)
            outBasePath = CStr(ws.Cells(i, STORAGE_COL_BASE_PATH).Value)
            GetStorageConfig = True
            Exit Function
        End If
    Next i
    
    GetStorageConfig = False
    Exit Function

ErrHandler:
    GetStorageConfig = False
End Function

' ================================================================
' 関数名: CombinePath
' 概要:   ベースパスとファイル名を結合する
' 引数:   basePath - ベースパス
'         fileName - ファイル名
' 戻り値: String - 結合されたパス
' 備考:   basePath の末尾にセパレータがなければ追加
' ================================================================
Private Function CombinePath(ByVal basePath As String, _
                               ByVal fileName As String) As String
    Dim sep As String
    sep = "\"
    
    If Right(basePath, 1) = "\" Or Right(basePath, 1) = "/" Then
        CombinePath = basePath & fileName
    Else
        CombinePath = basePath & sep & fileName
    End If
End Function
```
