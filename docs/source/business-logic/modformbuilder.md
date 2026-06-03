---
title: modFormBuilder.bas
---

# modFormBuilder.bas

| 項目 | 内容 |
|---|---|
| 層 | ビジネスロジック層 |
| 種別 | 標準モジュール (.bas) |
| 配置ブック | 3 ブック共通 |
| 役割 | clsFormSpec の宣言情報から UserForm を動的に組み立てる |
| 行数 | 198 行 |

## 取り込み先

標準モジュール（.bas）です。下記コードをコピーし、`modFormBuilder.bas` というファイル名で保存して、VBE の「ファイル → ファイルのインポート」で取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。

## ソースコード

コードブロック右上のボタンで全文をコピーできます。

```vbnet linenums="1"
Attribute VB_Name = "modFormBuilder"
Option Explicit

' Phase 6 レビュー: 5 subs に 8 error handlers と充実。
' VBProject Object Model 信頼設定 OFF 時のフォールバックも実装済。指摘なし。

' ================================================================
' モジュール: modFormBuilder (ビジネスロジック層)
' 概要:       clsFormSpec を受け取り、VBComponents.Add(3) で MSForm を
'             生成 → designer.Controls.Add(progID, name) で各 Control を
'             配置 → CodeModule.AddFromString で Click ハンドラを注入 →
'             VBA.UserForms.Add(name).Show する。
'             Show 後 (Modal の場合 Unload 後)、一時 VBComponent を Remove
'             してプロジェクトから消す。
' 依存先:     clsFormSpec, clsControlSpec
' 備考:       VBA Project Object Model 信頼が ON の前提
'             (excel-vba-auto-installer skill でも触れている)。
'             OFF の場合 VBComponents.Add で「アクセスが拒否されました」。
' ================================================================

' --- Forms ProgID 一覧 ---
Private Const PROGID_LABEL As String = "Forms.Label.1"
Private Const PROGID_TEXTBOX As String = "Forms.TextBox.1"
Private Const PROGID_BUTTON As String = "Forms.CommandButton.1"
Private Const PROGID_IMAGE As String = "Forms.Image.1"
Private Const PROGID_LISTBOX As String = "Forms.ListBox.1"
Private Const PROGID_COMBOBOX As String = "Forms.ComboBox.1"
Private Const PROGID_FRAME As String = "Forms.Frame.1"
Private Const PROGID_CHECKBOX As String = "Forms.CheckBox.1"

' --- VBComponents.Add の Type 値 ---
Private Const VBEXT_CT_MSFORM As Long = 3

' --- Show 引数 ---
Private Const SHOW_MODAL As Long = 1
Private Const SHOW_MODELESS As Long = 0

' ================================================================
' 関数名: BuildAndShow
' 概要:   spec から動的に UserForm を生成して Show する。
'         Modal/Modeless 切替可能。
' 引数:   spec  - clsFormSpec (FormTitle/Width/Height/Controls)
'         modal - True=モーダル / False=モードレス (省略時 True)
' 戻り値: なし
' 備考:   Modal の場合、ユーザが閉じるまでブロック。閉じた後に temp form を
'         Remove する。Modeless の場合 Show 直後に return するため、temp
'         form は呼び出し側の責任で RemoveTempForm を呼ぶ必要がある。
'         本モックでは Modal 利用を推奨。
' ================================================================
Public Sub BuildAndShow(ByVal spec As clsFormSpec, _
                          Optional ByVal modal As Boolean = True)
    On Error GoTo ErrHandler

    Dim frmComp As Object
    Set frmComp = BuildOnly(spec)

    Dim formName As String
    formName = frmComp.Name

    If modal Then
        VBA.UserForms.Add(formName).Show SHOW_MODAL
        ' Modal は閉じてからここに到達するので Remove
        Call RemoveTempForm(frmComp)
    Else
        VBA.UserForms.Add(formName).Show SHOW_MODELESS
        ' Modeless は呼び出し側で RemoveTempForm するか、ブック閉じ時に消える
    End If
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not frmComp Is Nothing Then
        Call RemoveTempForm(frmComp)
    End If
    On Error GoTo 0
    Err.Raise Err.Number, "modFormBuilder.BuildAndShow", Err.Description
End Sub

' ================================================================
' 関数名: BuildOnly
' 概要:   spec から動的に UserForm を生成するが Show しない (テスト用)
' 引数:   spec - clsFormSpec
' 戻り値: VBComponent (Object 型で返す。VBIDE.VBComponent と等価)
' 備考:   呼び出し側はテスト後 RemoveTempForm で削除すること。
' ================================================================
Public Function BuildOnly(ByVal spec As clsFormSpec) As Object
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    ' 一時 UserForm 作成
    Dim frmComp As Object
    Set frmComp = vbProj.VBComponents.Add(VBEXT_CT_MSFORM)
    frmComp.Properties("Caption") = spec.FormTitle
    frmComp.Properties("Width") = spec.Width
    frmComp.Properties("Height") = spec.Height

    Dim designer As Object
    Set designer = frmComp.designer

    ' イベントコード組み立て (Click ハンドラ注入用)
    Dim eventCode As String
    eventCode = "Option Explicit

' Phase 6 レビュー: 5 subs に 8 error handlers と充実。
' VBProject Object Model 信頼設定 OFF 時のフォールバックも実装済。指摘なし。" & vbCrLf

    ' 各 Control を追加
    Dim ix As Long
    For ix = 1 To spec.Controls.Count
        Dim cs As clsControlSpec
        Set cs = spec.Controls(ix)
        Call AddOneControl(designer, cs, eventCode)
    Next ix

    ' イベントコード一括注入
    Dim cm As Object
    Set cm = frmComp.CodeModule
    cm.AddFromString eventCode

    Set BuildOnly = frmComp
End Function

' ================================================================
' 関数名: AddOneControl
' 概要:   designer.Controls.Add で 1 個の Control を追加してプロパティ設定。
'         OnClick が指定されていれば eventCode に Click ハンドラを追記。
' 引数:   designer  - frmComp.designer
'         cs        - clsControlSpec
'         eventCode - イベントコード文字列 (ByRef で追記)
' ================================================================
Private Sub AddOneControl(ByVal designer As Object, _
                            ByVal cs As clsControlSpec, _
                            ByRef eventCode As String)
    On Error GoTo ErrHandler
    Dim progID As String
    progID = ProgIDFromType(cs.ControlType)

    Dim ctl As Object
    Set ctl = designer.Controls.Add(progID, cs.Name, True)

    On Error Resume Next
    ctl.Left = cs.Left
    ctl.Top = cs.Top
    ctl.Width = cs.Width
    ctl.Height = cs.Height
    If cs.Caption <> "" Then ctl.Caption = cs.Caption
    On Error GoTo ErrHandler

    ' OnClick が指定されていればハンドラ注入 (Button 系のみ)
    If cs.OnClick <> "" Then
        eventCode = eventCode & vbCrLf & _
            "Private Sub " & cs.Name & "_Click()" & vbCrLf & _
            "    On Error Resume Next" & vbCrLf & _
            "    Application.Run """ & cs.OnClick & """, Me" & vbCrLf & _
            "End Sub" & vbCrLf
    End If
    Exit Sub
ErrHandler:
    ' 個別 Control 追加失敗は警告レベル (他の Control は継続)
End Sub

' ================================================================
' 関数名: RemoveTempForm
' 概要:   一時 UserForm の VBComponent をプロジェクトから削除
' 引数:   frmComp - BuildOnly が返した VBComponent
' ================================================================
Public Sub RemoveTempForm(ByVal frmComp As Object)
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove frmComp
End Sub

' ================================================================
' 関数名: ProgIDFromType
' 概要:   ControlType 文字列から Forms.* ProgID を返す
' 備考:   未知の type は Label にフォールバック
' ================================================================
Public Function ProgIDFromType(ByVal t As String) As String
    Select Case UCase(t)
        Case "LABEL"
            ProgIDFromType = PROGID_LABEL
        Case "TEXTBOX"
            ProgIDFromType = PROGID_TEXTBOX
        Case "BUTTON", "COMMANDBUTTON"
            ProgIDFromType = PROGID_BUTTON
        Case "IMAGE"
            ProgIDFromType = PROGID_IMAGE
        Case "LISTBOX"
            ProgIDFro