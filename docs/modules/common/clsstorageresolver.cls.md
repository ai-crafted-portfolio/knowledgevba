---
title: clsStorageResolver.cls
description: clsStorageResolver.cls のソースコード（コピペ用）
---

# clsStorageResolver.cls

**配置先**: 共通モジュール（検索.xlsm / 管理.xlsm 共通）
**種類**: クラスモジュール
**更新日**: 2026-06-05 01:27 JST

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `clsStorageResolver.cls`
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
Private Const STORAGE_TYPE_BOX As String = "BOX"

' パス種別

Private m_logger As clsLogger

' ================================================================
' 関数名: Init
' 概要:   初期化（ロガー参照を保持）
' 引数:   logger - ログ出力用
' 戻り値: なし
' ================================================================
Public Sub Init(ByVal logger As clsLogger)
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0760] clsStorageResolver.Init ENTER"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0761] clsStorageResolver.ResolveLink ENTER"  ' [ADR-0100]
    On Error GoTo ErrHandler
    
    Dim storageType As String
    Dim pathType As String
    Dim basePath As String
    
    If Not GetStorageConfig(docType, storageType, pathType, basePath) Then
        If Not m_logger Is Nothing Then
            m_logger.LogWarn "clsStorageResolver", "ResolveLink", _
                              ChrW(&H683C) & ChrW(&H7D0D) & ChrW(&H5148) & ChrW(&H8A2D) & ChrW(&H5B9A) & ChrW(&H306A) & ChrW(&H3057) & ": docType=" & docType, , LOG_STORAGE_OPEN_ENTRY
        End If
        ResolveLink = ""
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0762] clsStorageResolver.ResolveLink EXIT-OK"  ' [ADR-0100]
        Exit Function
    End If
    
    ' BOXの場合はBOX URLをそのまま返す（ファイル直指定不可）
    If storageType = STORAGE_TYPE_BOX Then
        ResolveLink = basePath
        If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0763] clsStorageResolver.ResolveLink EXIT-OK"  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0764] clsStorageResolver.ResolveLink EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0765] clsStorageResolver.ResolveLink EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
    If Not m_logger Is Nothing Then
        m_logger.LogError "clsStorageResolver", "ResolveLink", _
                           ChrW(&H30EA) & ChrW(&H30F3) & ChrW(&H30AF) & ChrW(&H751F) & ChrW(&H6210) & ChrW(&H5931) & ChrW(&H6557) & ": " & Err.Description, , LOG_STORAGE_OPEN_EXIT_OK
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0766] clsStorageResolver.GetStorageConfig ENTER"  ' [ADR-0100]
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
            If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0767] clsStorageResolver.GetStorageConfig EXIT-OK"  ' [ADR-0100]
            Exit Function
        End If
    Next i
    
    GetStorageConfig = False
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0768] clsStorageResolver.GetStorageConfig EXIT-OK"  ' [ADR-0100]
    Exit Function

ErrHandler:
    If modCommon.gDebugLevel >= DEBUG_LEVEL_ERROR Then Debug.Print "[D-0769] clsStorageResolver.GetStorageConfig EXIT-ERR " & "errNum=" & Err.Number & " desc=" & Err.Description  ' [ADR-0100]
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
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0770] clsStorageResolver.CombinePath ENTER"  ' [ADR-0100]
    Dim sep As String
    sep = "\"
    
    If Right(basePath, 1) = "\" Or Right(basePath, 1) = "/" Then
        CombinePath = basePath & fileName
    Else
        CombinePath = basePath & sep & fileName
    End If
    If modCommon.gDebugLevel >= DEBUG_LEVEL_TRACE Then Debug.Print "[D-0771] clsStorageResolver.CombinePath EXIT-OK"  ' [ADR-0100]
End Function

' ADR-0006/0090/0094 JP literal removal:
Private Property Get STORAGE_TYPE_SHARED() As String
    STORAGE_TYPE_SHARED = ChrW(&H5171) & ChrW(&H6709) & ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30EB) & ChrW(&H30C0)
End Property

Private Property Get PATH_TYPE_FILE() As String
    PATH_TYPE_FILE = ChrW(&H30D5) & ChrW(&H30A1) & ChrW(&H30A4) & ChrW(&H30EB) & ChrW(&H6307) & ChrW(&H5B9A)
End Property

Private Property Get PATH_TYPE_FOLDER() As String
    PATH_TYPE_FOLDER = ChrW(&H30D5) & ChrW(&H30A9) & ChrW(&H30EB) & ChrW(&H30C0) & ChrW(&H6307) & ChrW(&H5B9A)
End Property
```
