---
title: modFileIO.bas
description: modFileIO.bas のソースコード（コピペ用）
---

# modFileIO.bas

**配置先**: 共通モジュール（3 ブック共通）  
**種類**: 標準モジュール

---

## 保存方法

下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。

- 場所: `C:\KnowledgeMgr\installer\vba_modules\common\\`
- ファイル名: `modFileIO.bas`
- ファイルの種類: **すべてのファイル**
- 文字コード: **ANSI**（Shift-JIS）

> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。
> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。
> 改行コードは CRLF（Windows 標準）のままで OK です。

---

## ソースコード

```vb
Attribute VB_Name = "modFileIO"
Option Explicit

' ================================================================
' モジュール: modFileIO（ユーティリティ層）
' 概要:       Shift_JIS ファイルI/O、ファイル/フォルダ操作の純粋関数群
'             3段フォールバック方式（ADODB.Stream → UNO → Open For Output）
'             でExcel/LibreOffice 両環境に対応
' 依存先:     modCommon
' ================================================================

' ================================================================
' 関数名: WriteShiftJisFile
' 概要:   Shift_JIS で文字列をファイルに書き込む
'         ADODB.Stream（Windows/Excel） → UNO TextOutputStream（LibreOffice）
'         → Open For Output（最終フォールバック）の3段階で試行
' 引数:   filePath - 出力先ファイルパス
'         content  - 書き込む文字列
' 戻り値: Boolean - 成功なら True
' 備考:   既存ファイルは上書き
' ================================================================
Public Function WriteShiftJisFile(ByVal filePath As String, _
                                    ByVal content As String) As Boolean
    ' 第1段: ADODB.Stream
    If TryWriteWithADODB(filePath, content) Then
        WriteShiftJisFile = True
        Exit Function
    End If
    
    ' 第2段: UNO TextOutputStream（LibreOffice環境）
    If TryWriteWithUNO(filePath, content) Then
        WriteShiftJisFile = True
        Exit Function
    End If
    
    ' 第3段: 古典的な Open For Output
    If TryWriteWithOpenStatement(filePath, content) Then
        WriteShiftJisFile = True
        Exit Function
    End If
    
    WriteShiftJisFile = False
End Function

' 第1段: ADODB.Stream による書き込み
Private Function TryWriteWithADODB(ByVal filePath As String, _
                                     ByVal content As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2  ' adTypeText
    stream.Charset = CHARSET_SJIS
    stream.Open
    stream.WriteText content
    stream.SaveToFile filePath, 2  ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing
    
    TryWriteWithADODB = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    On Error GoTo 0
    TryWriteWithADODB = False
End Function

' 第2段: UNO TextOutputStream による書き込み（旧 LibreOffice 環境用）
' rev14 注意:
'   LibreOffice Basic 固有の組み込み createUnoService() は Excel VBA
'   （日本語 Windows 版含む）では未定義シンボルとなり、Option Explicit
'   による静的解析でプロジェクト全体がコンパイル失敗する。On Error GoTo
'   で実行時エラーをトラップしても、**コンパイル段階**で落ちるため
'   Application.Run が「マクロが使用できない」系のメッセージを出す。
'
'   現行ビルドは Excel を唯一の対象ランタイムとするため、UNO パスを
'   ソース上から撤去し、本関数は常に False を返す no-op に固定する。
'   LibreOffice headless での自動テストが必要になった場合は、別モジュール
'   として UNO 対応版を差し替える方式とする（本モジュールには混入させない）。
'
'   ADODB.Stream パス（第1段）が Windows Excel では常に通るため、
'   実運用で第2段に落ちるケースは無い。最終フォールバックとして
'   Open For Output（第3段）も健在。
Private Function TryWriteWithUNO(ByVal filePath As String, _
                                   ByVal content As String) As Boolean
    ' 未使用引数の参照抑止（WARN 回避）
    If Len(filePath) = 0 And Len(content) = 0 Then
        TryWriteWithUNO = False
    End If
    TryWriteWithUNO = False
End Function

' 第3段: Open For Output による書き込み（最終フォールバック）
Private Function TryWriteWithOpenStatement(ByVal filePath As String, _
                                             ByVal content As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Output As #fNum
    Print #fNum, content
    Close #fNum
    
    TryWriteWithOpenStatement = True
    Exit Function

ErrHandler:
    On Error Resume Next
    Close #fNum
    On Error GoTo 0
    TryWriteWithOpenStatement = False
End Function

' ================================================================
' 関数名: ReadShiftJisFile
' 概要:   Shift_JIS ファイルを文字列として読み込む
'         3段フォールバック方式
' 引数:   filePath - 読み込み元ファイルパス
' 戻り値: String - ファイル内容（失敗時は空文字列）
' ================================================================
Public Function ReadShiftJisFile(ByVal filePath As String) As String
    Dim result As String
    
    ' 第1段: ADODB.Stream
    If TryReadWithADODB(filePath, result) Then
        ReadShiftJisFile = result
        Exit Function
    End If
    
    ' 第2段: UNO
    If TryReadWithUNO(filePath, result) Then
        ReadShiftJisFile = result
        Exit Function
    End If
    
    ' 第3段: Open For Input
    If TryReadWithOpenStatement(filePath, result) Then
        ReadShiftJisFile = result
        Exit Function
    End If
    
    ReadShiftJisFile = ""
End Function

' 第1段: ADODB.Stream
Private Function TryReadWithADODB(ByVal filePath As String, _
                                    ByRef outContent As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2  ' adTypeText
    stream.Charset = CHARSET_SJIS
    stream.Open
    stream.LoadFromFile filePath
    outContent = stream.ReadText
    stream.Close
    Set stream = Nothing
    
    TryReadWithADODB = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    On Error GoTo 0
    outContent = ""
    TryReadWithADODB = False
End Function

' 第2段: UNO（旧 LibreOffice 環境用、Excel では no-op）
' rev14 注意: TryWriteWithUNO と同じ理由で createUnoService を撤去し、
'             Excel VBA でコンパイル可能な no-op に固定した。詳細は
'             TryWriteWithUNO の冒頭コメントを参照。
Private Function TryReadWithUNO(ByVal filePath As String, _
                                  ByRef outContent As String) As Boolean
    ' 未使用引数の参照抑止（WARN 回避）
    If Len(filePath) = 0 Then
        outContent = ""
    End If
    outContent = ""
    TryReadWithUNO = False
End Function

' 第3段: Open For Input
Private Function TryReadWithOpenStatement(ByVal filePath As String, _
                                            ByRef outContent As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim fNum As Integer
    Dim buffer As String
    Dim line As String
    
    fNum = FreeFile
    Open filePath For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, line
        If buffer = "" Then
            buffer = line
        Else
            buffer = buffer & vbCrLf & line
        End If
    Loop
    Close #fNum
    
    outContent = buffer
    TryReadWithOpenStatement = True
    Exit Function

ErrHandler:
    On Error Resume Next
    Close #fNum
    On Error GoTo 0
    outContent = ""
    TryReadWithOpenStatement = False
End Function

' ================================================================
' 関数名: FileExists
' 概要:   ファイルが存在するか確認
' 引数:   filePath - ファイルパス
' 戻り値: Boolean - 存在すれば True
' 備考:   Dir関数を使用（LibreOffice互換）
' ================================================================
Public Function FileExists(ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler
    FileExists = (Dir(filePath) <> "")
    Exit Function
ErrHandler:
    FileExists = False
End Function

' ================================================================
' 関数名: FolderExists
' 概要:   フォルダが存在するか確認
' 引数:   folderPath - フォルダパス
' 戻り値: Boolean - 存在すれば True
' ================================================================
Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo ErrHandler
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    Exit Function
ErrHandler:
    FolderExists = False
End Function

' ================================================================
' 関数名: EnsureFolder
' 概要:   フォルダが存在しなければ作成する
' 引数:   folderPath - フォルダパス
' 戻り値: Boolean - 成功（既存含む）なら True
' ================================================================
Public Function EnsureFolder(ByVal folderPath As String) As Boolean
    On Error GoTo ErrHandler
    
    If FolderExists(folderPath) Then
        EnsureFolder = True
        Exit Function
    End If
    
    MkDir folderPath
    EnsureFolder = True
    Exit Function

ErrHandler:
    EnsureFolder = False
End Function

' ================================================================
' 関数名: ListFilesInFolder
' 概要:   フォルダ内の指定拡張子ファイル一覧を取得する
' 引数:   folderPath - 対象フォルダパス
'         extension  - 拡張子（"txt", "bas" 等、ドット不要）
' 戻り値: Variant - ファイル名の配列（拡張子込み）。0件の場合は空配列
' 備考:   サブフォルダは辿らない
' ================================================================
' M-4: 空配列ガード ? 呼出側は IsEmpty / UBound < LBound チェック必須
'        本関数は ReDim 済み配列を返すが、Dir 失敗時は ReDim Preserve なしで返る場合がある
' s-2 contract: 本モジュール内の Kill / MkDir / Open For Output に渡されるパスは
'                呼出側で IsValidKnowledgeId / 自前パス検証済みであること。
'                本モジュールはパス検証責任を負わない (低レイヤ I/O 専念)。
' v14 D-4: UNO 段は LibreOffice 対応用 (本プロジェクトでは未使用)。
'          将来 LibreOffice 互換が必要なら別 modFileIO_UNO に切出し、本モジュールは
'          Excel 専用 ADODB + Open For Output の 2 段フォールバックに専念する。
' v14 s-3: ADODB fallback (Open For Output) で SJIS が壊れる可能性。
'          modFileIO は循環依存回避のため Logger 非依存だが、呼出側で
'          ADODB 失敗 → fallback 経路選択時に Logger.LogWarn で通知すること。
Public Function ListFilesInFolder(ByVal folderPath As String, _
                                    ByVal extension As String) As Variant
    On Error GoTo ErrHandler
    
    Dim results() As String
    Dim count As Long
    Dim fileName As String
    Dim searchPattern As String
    Dim sep As String
    
    count = 0
    ReDim results(0 To 99)  ' 初期100件
    
    ' パス末尾のセパレータ調整
    sep = "\"
    If Right(folderPath, 1) <> "\" And Right(folderPath, 1) <> "/" Then
        searchPattern = folderPath & sep & "*." & extension
    Else
        searchPattern = folderPath & "*." & extension
    End If
    
    fileName = Dir(searchPattern)
    Do While fileName <> ""
        If count > UBound(results) Then
            ReDim Preserve results(0 To count + 99)
        End If
        results(count) = fileName
        count = count + 1
        fileName = Dir()
    Loop
    
    If count = 0 Then
        ListFilesInFolder = Array()
    Else
        ReDim Preserve results(0 To count - 1)
        ListFilesInFolder = results
    End If
    Exit Function

ErrHandler:
    ListFilesInFolder = Array()
End Function

' ================================================================
' 関数名: DeleteFile
' 概要:   ファイルを削除する（物理削除）
' 引数:   filePath - ファイルパス
' 戻り値: Boolean - 成功なら True
' ================================================================
Public Function DeleteFile(ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler
    
    If Not FileExists(filePath) Then
        DeleteFile = False
        Exit Function
    End If
    
    Kill filePath
    DeleteFile = True
    Exit Function

ErrHandler:
    DeleteFile = False
End Function

' ================================================================
' 関数名: ConvertLocalPathToURL
' 概要:   ローカルファイルパスを UNO で使う URL 形式に変換
' 引数:   localPath - ローカルパス（例: "C:\work\file.txt"）
' 戻り値: String - URL 形式（例: "file:///C:/work/file.txt"）
' 備考:   UNO TextOutputStream 等で使用する内部ヘルパ
' ================================================================
Private Function ConvertLocalPathToURL(ByVal localPath As String) As String
    Dim result As String
    result = localPath
    
    ' バックスラッシュをスラッシュに
    result = Replace(result, "\", "/")
    
    ' Windows 形式（例: "C:/..."）の場合は "file:///" を前置
    If InStr(result, ":") > 0 Then
        ConvertLocalPathToURL = "file:///" & result
    Else
        ' Unix 形式（"/..."）の場合
        If Left(result, 1) = "/" Then
            ConvertLocalPathToURL = "file://" & result
        Else
            ConvertLocalPathToURL = "file:///" & result
        End If
    End If
End Function
```
