---
title: 設定・カスタマイズ（全項目リファレンス）
description: 保存先・画面の見た目・入力フォーム・フォーマットを、設定シートや外部ファイルの編集で調整するための完全な一覧と記入サンプル
---

# 設定・カスタマイズ（全項目リファレンス）

ナレッジ管理ツールは、画面の見た目・保存先・入力フォームの大きさ・入力項目の構成などを、**ブックの中のコードを書き換えずに**、設定シートや外部のテキストファイルを編集するだけで調整できます。このページは、利用者が編集できる設定を**1項目ずつ漏れなく**まとめた完全リファレンスです。各設定について「どのファイル（シート）の・どの項目を・どう変えるか」「いつ反映されるか」を、**コピーして使える記入サンプル付き**で説明します。各章末の表に続けて記入例を、最後の章（11章）には実際に動く**まるごとのサンプルファイル**を載せています。

## 1. この章でできること

- 保存先フォルダ（ナレッジ本体・フォーマット・画面定義・バックアップ）の置き場所を変える
- 画面のタブ名・見出し・フォント・色・列幅・枠固定・ボタンの文言や大きさ・一覧の見た目を変える
- 登録・検索で開く入力フォームの幅や高さ、項目の行の高さを変える
- 入力項目（フォーマット）の名前・型・必須/任意・選択肢・並び順を変える
- ボタンや入力項目を増やす

設定の入り口は3つです。管理ブックの **「格納先設定」シート**（保存先と動作）、**画面定義ファイル**（`M-xx.txt`：画面の見た目と入力フォームの寸法）、**フォーマット定義ファイル**（`formats` フォルダ内の `<フォーマットID>.txt`：入力項目）です。

## 2. 編集前の共通の注意

テキストファイル（`M-xx.txt` や フォーマット定義）を直接編集するときは、次を必ず守ってください。

!!! warning "文字コードと改行"
    メモ帳で開く場合は、保存時の文字コードを **「ANSI」（日本語環境では Shift_JIS）**、改行を **CRLF** のままにしてください。**UTF-8 で保存すると日本語の項目名や見出しが文字化けし、画面が読み込めなくなることがあります。**

- 編集する前に、対象ファイルの**コピーを取って**おいてください（うまくいかなかったら戻せます）。
- フォルダの区切りは**半角の `\`（バックスラッシュ）**を使います。全角の「￥」や全角スペースが混じるとパスが壊れます。
- 値は「`キー=値`」の形で書きます。`=` の前後に余計な空白を入れない方が確実です。
- 各セクションは `===` の行で区切ります（区切りは画面定義・フォーマット定義に共通の書き方です）。
- 行頭が `[`〜`]` の行は「セクション（区切り）」です。セクション名そのものは変えないでください。
- 色は **`#RRGGBB` の16進**（例 `#1F3864`）でも、**日本語の色名**（例「濃いグレー」「白」）でも指定できます。色を付けたくないときは「―（色を付けない）」と書けます。

## 3. 保存先と動作の設定（「格納先設定」シート）

保存先フォルダやログの設定は、管理ブックの **「格納先設定」シート** で行います。各行の項目名の右どなりのセルに値を入れ、**Ctrl+S** で上書き保存します（専用の保存ボタンはありません。セルに直接入力 → 通常の保存、という流れです）。設定はブックと一緒に保存されます。

設定は次の **7項目**です。

| 項目（行の見出し） | 意味 | 既定の値（例） | 取りうる値・注意 |
| --- | --- | --- | --- |
| データフォルダ | ナレッジ本体（1件1ファイル）の保存先 | `C:\KnowledgeMgr\data\` | 末尾の `\` を必ず付ける |
| フォーマットフォルダ | フォーマット定義ファイルの置き場 | `C:\KnowledgeMgr\formats\` | 末尾の `\` を必ず付ける |
| UI 定義フォルダ | 画面定義ファイル（`M-xx.txt`）の置き場 | `C:\KnowledgeMgr\ui\管理\`（検索ブックは `…\ui\検索\`） | 末尾の `\` を必ず付ける。**役割（管理／検索）のサブフォルダまで**指す |
| バックアップフォルダ | 自動バックアップの保存先 | `C:\KnowledgeMgr\backup\` | 末尾の `\` を必ず付ける |
| デバッグレベル | ログに残す記録の詳しさ | `INFO` | 下表の6段階から選ぶ（ドロップダウン） |
| ログ回転行数 | ログとして保持する行数の上限。超えると古い行から自動削除 | `10000` | 数値 |
| UI スキーマ失敗時 | 画面定義に不備があったときの動き | `safeDefault` | `safeDefault` / `warn` / `abort` から選ぶ（ドロップダウン） |

!!! warning "フォルダのパスは末尾の `\` が必須です"
    パスはそのままファイル名とつなげて使われます。末尾に `\` が無いと、フォルダ名とファイル名がくっついて「ファイルが見つかりません」になります。例えば画面定義フォルダは `\\サーバ名\共有名\KnowledgeMgr\ui\管理\` のように**必ず `\` で終える**ようにしてください。

ネットワーク共有を使う場合は、**`\\サーバ名\共有名\…\管理\` の形（UNC）で直接**入れてください。`Z:\` のような割り当てドライブは、Excel を「管理者として実行」していると見えなくなることがあり、「見つからない」の原因になります。割り当てドライブよりも UNC を直接指定する方が確実です。

**デバッグレベルの6段階**（上にいくほど記録が少なく、下にいくほど詳しくなります）：

| 値 | 記録される内容 |
| --- | --- |
| `OFF` | 記録しない |
| `ERROR` | エラーのみ |
| `WARN` | エラーと警告 |
| `INFO` | 上記に加えて通常の操作（既定） |
| `DEBUG` | 上記に加えて調査向けの詳細 |
| `TRACE` | 最も詳しい記録 |

通常は `INFO` のままで十分です。不具合を調べたいときだけ一時的に `DEBUG` や `TRACE` に上げ、調査が終わったら戻すと、ログが余計に大きくなりません。

**「UI スキーマ失敗時」**は、画面定義に不備があったときの動きです。`safeDefault` はその画面だけを安全な既定表示にして他の画面は通常どおり起動、`warn` は知らせたうえで続行、`abort` は1件でも不備があれば起動を中止します。

!!! note "初回起動時の保存先の確認"
    インストール後にブックを初めて開いたとき、保存先がすべて空であれば、**保存先の親フォルダをたずねるダイアログ**が表示されます（既定 `C:\KnowledgeMgr`）。そのまま入力すると、その下に上記のフォルダが自動で設定され、ブックに保存されます。設定し忘れても既定値で起動します。

**反映のタイミング:** 保存先や動作の設定は**起動時に読み込まれる**ため、変更後は**ブックを開き直す（実質的に再起動）**と反映されます。

## 4. 画面（シート）の見た目を変える

各画面の見た目は、`UI 定義フォルダ` 内の **`M-xx.txt`**（例：検索ブックの検索画面は `ui\検索\M-08.txt`、登録画面は `ui\検索\M-05.txt`、管理ブックのフォーマット一覧は `ui\管理\M-02.txt`）で決まります。ファイルは `[セクション名]` の行に続けて `キー=値` を並べ、`===` で区切る形式です。書く順番は自由で、読み込み時に適切な順序で画面へ反映されます。

以下、セクションごとに、コードが実際に読むキーと**コピーして使える記入例**を挙げます。**「見た目」**の列は自由に変えてよいもの、**「構造」**と記したキーは位置や仕組みに関わるため原則そのままにしてください（→ 9章）。記入例の値は実際の画面定義から採っています。

### 4-1. シート全体・全体書式

**[SHEET]**（シート全体）

| キー | 意味 | 値の例 | 区分 |
| --- | --- | --- | --- |
| `SheetTitle` | 画面の見出し | `ナレッジ検索` | 見た目 |
| `Visible` | 表示/非表示 | `visible` / `hidden` | 見た目 |
| `TabColor` | シートタブの色 | `#FFCC00` / 色名 | 見た目 |
| `HideColumnsFrom` | 指定した列以降を隠す | `G` | 見た目 |
| `SheetName` | シートの内部名 | （変更しない） | 構造 |

```
[SHEET]
SheetName=ナレッジ検索
SheetTitle=ナレッジ検索
Visible=visible
TabColor=―（色を付けない）
HideColumnsFrom=G
```

**[STYLE]**（画面全体の既定書式）

| キー | 意味 | 値の例 |
| --- | --- | --- |
| `BaseFont` | 既定フォント名 | `メイリオ` |
| `BaseFontSize` | 既定文字サイズ | `10` |
| `BaseFontColor` | 既定文字色 | `濃いグレー` / `#333333` |
| `SheetBackColor` | 背景色 | `白` / `#FFFFFF` |
| `BaseAlignH` | 横位置 | `left` / `center` / `right` |
| `BaseAlignV` | 縦位置 | `center` |
| `BaseWrap` | 折り返し | `on` / `off` |
| `BaseNumberFormat` | 既定の表示形式 | `standard` |

```
[STYLE]
BaseFont=メイリオ
BaseFontSize=10
BaseFontColor=濃いグレー
SheetBackColor=白
BaseAlignH=left
BaseAlignV=center
BaseWrap=off
BaseNumberFormat=standard
```

### 4-2. 見出し・帯・列幅・枠固定

**[SCREEN]**（上部の見出し帯）：`Title`（タイトル）、`Width`（列数）、`BackColor`／`FontColor`（帯と文字の色）、`EndCol`（帯の右端列）、`BarHeight`（帯の高さ・ポイント）。

```
[SCREEN]
Width=7（列数）
Height=auto
Title=ナレッジ検索
BackColor=#1F3864
FontColor=#FFFFFF
EndCol=G
BarHeight=30
```

**[SUBHEADER]**（小見出し帯）：`Title`、`BackColor`／`FontColor`、`EndCol`、`Height`（高さ・ポイント）、`Cell`（位置＝構造）。

```
[SUBHEADER]
Title=条件で探す
Cell=A3
BackColor=#2F5597
FontColor=#FFFFFF
EndCol=G
Height=22
```

**[COLUMN]**（列の幅）：`ColumnWidth_A`、`ColumnWidth_B`… のように **`ColumnWidth_<列の文字>`** で列ごとに指定します。値は Excel の列幅（おおよその文字数）。

```
[COLUMN]
ColumnWidth_A=6
ColumnWidth_B=20
ColumnWidth_C=24
ColumnWidth_D=24
ColumnWidth_E=24
ColumnWidth_F=17
```

**[ROW]**（行の高さ）：`RowHeight_<行番号>`（値はポイント）。標準の画面では使っていませんが、行の高さを固定したいときに使えます。

```
[ROW]
RowHeight_3=24
RowHeight_4=60
```

**[FREEZE]**（ウィンドウ枠の固定）：`Enabled`（`on` / `off`）、`Rows`／`Cols`（固定する位置）。

```
[FREEZE]
Enabled=on
Rows=13
Cols=0
```

### 4-3. ラベル・注釈・入力欄

**[LABEL]**（項目名などの固定文字）：`Text`（表示する文字）、`ForeColor`（文字色）、`FontSize`、`Cell`（位置＝構造）。

```
[LABEL]
Cell=B8
Text=キーワード
```

**[NOTE]**（説明文のボックス）：`Text`（本文。`\n` で改行）、`BackColor`、`Border`（`thin` など）、`Cell`（範囲＝構造）。

```
[NOTE]
Cell=A4:G7
Text=キーワードを入れて「検索」を押すと、データの全ナレッジから一致するものを下の一覧に表示します。\nフォーマットで絞り込むこともできます。「条件クリア」で入力を空に戻します。
BackColor=#DEEBF7
Border=thin
```

**[INPUT]**（入力欄）：`MaxLength`（最大文字数）、`InputColor`（背景色）、`Border`、`Value`（初期値）、`InputType`（`text` / `dropdown` など＝構造寄り）、`DropdownSource`（選択肢の取得元・構造）、`Cell`／`inputDataKey`（位置と内部キー＝構造）。

文字入力欄の例：

```
[INPUT]
Cell=C8:E8
inputDataKey=keyword
InputType=text
MaxLength=200
InputColor=#FFFFFF
Border=thin
```

ドロップダウン欄の例：

```
[INPUT]
Cell=C10:E10
inputDataKey=formatId
InputType=dropdown
DropdownSource=format:list
InputColor=#FFFFFF
Border=thin
```

### 4-4. ボタン・一覧表

**[BUTTON]**（ボタン）

| キー | 意味 | 値の例 | 区分 |
| --- | --- | --- | --- |
| `Text` | ボタンの文言 | `検索` | 見た目 |
| `Width` | 幅 | `110` | 見た目 |
| `Height` | 高さ | `22` | 見た目 |
| `BackColor` | 背景色 | `#4472C4` | 見た目 |
| `ForeColor` | 文字色 | `#FFFFFF` | 見た目 |
| `OnClick` | 押したときに動く処理の名前 | `Btn_…` | 構造 |
| `Cell` | 配置位置 | `B11` | 構造 |

```
[BUTTON]
Cell=B11
Text=検索
OnClick=Btn_SearchV23
Width=110
Height=22
BackColor=#4472C4
ForeColor=#FFFFFF
```

**[GRID]**（一覧表）

| キー | 意味 | 値の例 | 区分 |
| --- | --- | --- | --- |
| `InitialRows` | 最初に表示する行数 | `10` | 見た目 |
| `ColumnHeaders` | 列の見出し（表示文字） | `No,フォーマット,検索1,…` | 見た目 |
| `HeaderBackColor` | 見出し行の色 | `#D9E1F2` | 見た目 |
| `FreezeHeader` | 見出し行を固定するか | `TRUE` / `FALSE` | 見た目 |
| `StartCell` / `HeaderRow` | 表の開始位置・見出し行 | `A13` / `13` | 構造 |
| `Columns` | 各列の内部キー | `No,formatId,…` | 構造 |
| `Name` | 表の内部名 | `gridResult` | 構造 |

```
[GRID]
Name=gridResult
StartCell=A13
HeaderRow=13
InitialRows=10
Columns=No,formatId,searchField1,searchField2,searchField3,updatedAt,knowledgeNo
ColumnHeaders=No,フォーマット,検索1,検索2,検索3,更新日,ナレッジ番号
HeaderBackColor=#D9E1F2
FreezeHeader=FALSE
```

### 4-5. 入力部品・上級者向けのセクション

次のセクションも使えます。標準の画面では使っていない（または構造側で使う）ものが中心です。位置や内部名（`Cell` / `Name` / `OnClick` / `LinkedCell` など）は構造に関わるため、文言・色以外は原則そのままにしてください。記入例は最小構成です。

**[CHECK]**（チェック欄）：`Cell`、`DefaultChecked`

```
[CHECK]
Cell=B5
DefaultChecked=FALSE
```

**[CHECKBOX]**（チェックボックス部品）：`Caption`（文言）、`Cell`／`OnClick`／`LinkedCell`／`Name`（構造）

```
[CHECKBOX]
Cell=B5
Caption=対象に含める
OnClick=Btn_ToggleInclude
LinkedCell=H5
Name=chkInclude
```

**[LISTBOX]**（リスト部品）：`List`／`Choices`（選択肢）、`Cell`／`LinkedCell`／`Name`（構造）

```
[LISTBOX]
Cell=B6:C9
List=選択肢1|選択肢2|選択肢3
Choices=選択肢1|選択肢2|選択肢3
LinkedCell=H6
Name=lstSample
```

**[VALIDATION]**（入力規則の選択肢）：`List`／`Choices`／`DefaultValue`、`Cell`（構造）

```
[VALIDATION]
Cell=C8
List=低|中|高
Choices=低|中|高
DefaultValue=中
```

**[BUTTON_TEMPLATE]**（一覧の行ごとに自動で増減するボタンのひな型）：`GridRef`／`TemplateId`／`Name`（すべて構造）。実際のボタンは表示するデータが決まったときに自動で作られます。

```
[BUTTON_TEMPLATE]
GridRef=gridResult
TemplateId=rowOpen
Name=btnRowOpen
```

**[FORM_FROM_FORMAT] / [GRID_FROM_FORMAT]**（フォーマットの内容から入力欄や一覧を自動生成する仕組み）：内部利用のためのセクションです。中身は触らず、項目を増やしたいときは6章のフォーマット編集で行ってください。

**[FONT] / [HEADER]**：以前の様式のセクションです。現行の画面定義では [STYLE] / [SCREEN] を使います。

**反映のタイミング:** 画面の `M-xx.txt` を変更したら、画面上の **「画面更新」ボタン**を押すか、**ブックを開き直す**と反映されます（Excel 全体の再起動は不要です）。

## 5. 入力フォームの寸法・文言を変える

ナレッジを登録・閲覧・編集するときに開く**入力フォーム**の大きさや文言は、対応する `M-xx.txt` の **`[USERFORM]`** セクションで調整します。検索ブックでは、**登録フォームが `ui\検索\M-05.txt`** です。

寸法の単位はポイントで、数値を大きくすると広く・高くなります。`[USERFORM]` で指定できる項目は次のとおりです。

| キー | 意味 | 既定値 | 区分 |
| --- | --- | --- | --- |
| `formWidth` | フォームの幅 | `486` | 見た目 |
| `formHeight` | フォームの高さ（`0` で自動計算） | `0` | 見た目 |
| `headerHeight` | 見出し部の高さ | `48` | 見た目 |
| `labelWidth` | 項目名ラベルの幅 | `107` | 見た目 |
| `margin` | 左右の余白 | `18` | 見た目 |
| `rowHeightSingle` | 単一行項目の行の高さ | `48` | 見た目 |
| `rowHeightMulti` | 複数行項目の行の高さ | `114` | 見た目 |
| `rowHeightMultiLong` | 長い複数行項目の行の高さ | `129` | 見た目 |
| `subheaderHeight` | 小見出しの高さ | `28` | 見た目 |
| `buttonBarHeight` | ボタン帯の高さ | `48` | 見た目 |
| `buttonWidth` | ボタンの幅 | `113` | 見た目 |
| `buttonHeight` | ボタンの高さ | `26` | 見た目 |
| `buttonGap` | ボタンの間隔 | `11` | 見た目 |
| `bottomMargin` | 下の余白 | `16` | 見た目 |
| `badgeWidth` | 必須バッジの幅 | `29` | 見た目 |
| `badgeHeight` | 必須バッジの高さ | `16` | 見た目 |
| `fontSize` | 文字サイズ | `10` | 見た目 |
| `subheaderFontSize` | 小見出しの文字サイズ | `12` | 見た目 |
| `fontName` | フォント名 | （フォーム既定） | 見た目 |
| `multiLineScrollBars` | 複数行欄のスクロールバー | `vertical` | 見た目 |
| `caption` | フォームの見出し文言 | （画面ごと） | 見た目 |
| `backColor` | フォームの背景色 | `FFFFFF` | 見た目 |
| `formatHelp` | フォーマット選択のヘルプ文 | （任意） | 見た目 |
| `knowledgeNoLabel` | ナレッジ番号欄のラベル | （任意） | 見た目 |
| `knowledgeNoHelp` | ナレッジ番号欄のヘルプ文 | （任意） | 見た目 |
| `loadButtonLabel` | 読み込みボタンの文言 | （任意） | 見た目 |
| `headerHelp_<項目>` | 各項目に付けるヘルプ文（項目名ごと） | （任意） | 見た目 |
| `headerFields` | 上部に出す項目の指定 | （画面ごと） | 構造 |
| `formatSelectorType` | フォーマット選択欄の種類 | `dropdown` | 構造 |
| `formatRowEnabled` | フォーマット選択行の有無 | （画面ごと） | 構造 |
| `knowledgeNoRow` | ナレッジ番号行の有無 | （画面ごと） | 構造 |
| `buttons` | 並べるボタンの構成 | （画面ごと） | 構造 |

記入例（登録フォームの実際の `[USERFORM]`）：

```
[USERFORM]
formWidth=486
formHeight=0
fontName=メイリオ
fontSize=10
labelWidth=107
rowHeightSingle=48
rowHeightMulti=114
rowHeightMultiLong=129
margin=18
buttonBarHeight=48
caption=ナレッジ登録
multiLineScrollBars=vertical
formatSelectorType=dropdown
headerFields=knowledgeId
formatHelp=登録するナレッジのフォーマットを選びます。選ぶと下に項目が並びます。
headerHelp_knowledgeId=登録時に自動で割り当てられます。空き番号は手で入れ直すこともできます。
```

例えばフォームを少し広げたいときは、`formWidth=486` を `formWidth=620` のように書き換えます。閲覧フォーム・編集フォームの寸法も変えたいときは、それぞれの画面定義ファイルを用意し、同じ `[USERFORM]` セクションに値を書けば反映されます（管理ブックの設計画面からも調整できます）。専用ファイルが無い場合は上の既定値で表示されます。

**反映のタイミング:** 入力フォームの設定は、**フォームを開き直すだけ**で反映されます（Excel の再起動は不要です）。

## 6. フォーマット（入力項目）の定義

入力項目の構成は、**フォーマット定義ファイル**（`フォーマットフォルダ` 内の `<フォーマットID>.txt`、例 `formats\FAULT.txt`）で決まります。通常は**管理ブックの設計画面から**作成・編集できますが、中身はテキストです。先頭の `[FORMAT]` で型そのものを、続く `[FIELD]` で項目を1つずつ定義し、項目どうしは `===` で区切ります。**書いた順がそのまま画面の項目の並び順**になります。

**[FORMAT]（型そのもの）**

| キー | 意味 | 値の例 | 区分 |
| --- | --- | --- | --- |
| `FormatName` | 画面に表示される型の名前 | `障害対応` | 見た目 |
| `Status` | 状態（任意） | （任意） | 任意 |
| `FormatId` | 型のID（ファイル名・フォルダ名にもなる） | `FAULT` | 構造（変更非推奨） |

**[FIELD]（項目ごと）**

| キー | 意味 | 取りうる値 | 区分 |
| --- | --- | --- | --- |
| `FieldName` | 項目名（入力欄に表示される名前） | 文字列 | 見た目 |
| `FieldType` | 型 | `単一行` / `複数行` / `日付` / `選択` | 見た目 |
| `Required` | 入力を必須にするか | `TRUE` / `FALSE` | 見た目 |
| `MaxLength` | 文字数の上限（単一行で使用） | 数字 | 見た目 |
| `Rows` | 入力欄の高さ（複数行で使用） | 数字 | 見た目 |
| `Scroll` | 複数行欄でスクロールバーを出すか | `TRUE` / `FALSE` | 見た目 |
| `DropdownOptions` | 選択肢（`|` 区切り） | `低|中|高|緊急` | 見た目 |
| `fieldPlaceholder` | 入力欄に薄く出す見本 | 文字列 | 見た目 |
| `searchTarget` | キーワード検索の対象にするか | `TRUE` / `FALSE` | 見た目 |

型ごとの記入例（単一行・日付・選択・複数行）：

```
[FIELD]
FieldName=件名
FieldType=単一行
Required=TRUE
MaxLength=120
fieldPlaceholder=（例）経理システムへログインできない
searchTarget=TRUE
===
[FIELD]
FieldName=発生日時
FieldType=日付
Required=TRUE
fieldPlaceholder=（例）2026-06-09 10:30
searchTarget=TRUE
===
[FIELD]
FieldName=カテゴリ
FieldType=選択
Required=FALSE
DropdownOptions=アプリ障害|インフラ障害|ネットワーク障害|セキュリティ
fieldPlaceholder=（例）アプリ障害
searchTarget=FALSE
===
[FIELD]
FieldName=概要
FieldType=複数行
Required=TRUE
Rows=4
fieldPlaceholder=（例）応答時間が通常の10倍以上に増加
searchTarget=FALSE
```

!!! note "フォーマットIDの変更について"
    `FormatId` は、ファイル名であると同時に、そのフォーマットのナレッジを入れるフォルダ名にもなります。`FormatId` を変えると別のフォーマット扱いになり、元の `FormatId` で保存していたナレッジのフォルダが取り残されます。IDは原則そのままにしてください。

## 7. ボタンや入力項目を増やすには

### ボタンを1つ増やす

ボタンを増やすときは、**次の3つの作業をもれなく**行います。1つでも欠けると「ボタンはあるが押しても何も起きない」「処理はあるがボタンが現れない」状態になります。

1. **画面定義に `[BUTTON]` セクションを追加する。** `OnClick` には、押したときに実行する処理の名前（`Btn_` で始まる名前を推奨）を書きます。

    ```
    [BUTTON]
    Cell=N10
    Text=お気に入り
    OnClick=Btn_Favorite
    BackColor=#3366CC
    ForeColor=#FFFFFF
    ```

2. **`OnClick` に書いた名前と同じ処理（手続き）をブックの VBA に用意する。** 名前は `[BUTTON]` の `OnClick` と完全に一致させます。

3. **Excel の画面シートに実機ボタンを配置し、そのマクロを割り当てる。** 「開発 → 挿入 → フォームコントロール」のボタンを `[BUTTON]` の `Cell` で指定した位置に置き、右クリック →「マクロの登録」で手順2の処理を選びます。

一覧表の行ごとに増減するボタン（各行の「編集」「削除」など）は、`[BUTTON]` ではなく `[BUTTON_TEMPLATE]` でひな型を宣言します。実際のボタンは表示するデータが決まったときに自動で作られるため、実機での手配置は不要です。

### 入力項目を増やす

登録・修正の入力項目は、画面定義ではなく**フォーマット定義**で決まります。項目を増やすときは、管理ブックの**フォーマット設計**画面でフォーマットを編集し、項目（フィールド）を追加します。フォーマットを保存すると、そのフォーマットを使う既存ナレッジへ変更が反映されます（反映の前にバックアップが取られ、確認のメッセージが表示されます）。検索画面の検索条件のような、画面に固定された入力欄は、画面定義の `[INPUT]` で管理します。

## 8. 変更が反映されるタイミング（早見表）

| 変更したもの | 反映の方法 |
| --- | --- |
| 保存先のパス（「格納先設定」のフォルダ各種） | **ブックを開き直す**（実質的に再起動） |
| 画面（シート）の見た目（`M-xx.txt`） | **「画面更新」ボタン**、または開き直し |
| 入力フォームの寸法・文言（`[USERFORM]`） | **フォームを開き直す**だけ |
| フォーマット（入力項目） | **次にそのフォーマットを使う**（登録フォームや検索を開く）とき |

画面定義ファイルやフォーマット定義の**中身**は、その都度読み直されます。ただし、保存先フォルダの**場所そのもの**を変えたときだけは、開き直し（再起動）が必要です。

## 9. 触らない方がよい項目

見た目や寸法の調整と違い、次の項目は**動作と直接つながっている**ため、値を変えると画面やボタンが正しく動かなくなります。原則そのままにしてください。

- 各セクションの**セルの位置を表す値**（`Cell` / `StartCell` / `EndCell` / `HeaderRow` など）
- ボタンに割り当てられた**動作の指定**（`OnClick` の値）と、それに対応する処理
- 入力欄や一覧の**内部キー**（`inputDataKey`、`[GRID]` の `Columns`、`GridRef`、`Name`、`LinkedCell` など）
- シートの**内部名**（`SheetName`）と、`[SHEET]` セクションそのものの削除（必須のため）
- フォーマットの **`FormatId`**
- `[FORM_FROM_FORMAT]` / `[GRID_FROM_FORMAT]` / `[BUTTON_TEMPLATE]` といった、自動生成のためのセクション
- `[HEADER]` と `[INPUT]` などで**同じセル範囲を重ねて**指定すること（範囲が衝突し画面が崩れます）
- 定義ファイルを **Shift_JIS 以外で保存**すること（文字化け・読み込み失敗の原因）
- **バックアップを取らずに**編集すること（元に戻せなくなります）

これらを変えたい場合は、見た目の調整とは性質が異なるため、慎重に行うか、管理ブックの設計画面からの操作をおすすめします。

## 10. 使われない過去の名残

一部の画面定義ファイルには、過去の様式の名残として **`[VALUE]` セクション**や **`Editable` / `Source`** といったキーが残っていることがあります。現在のバージョンではこれらは読み込まれず、**書いても画面には反映されません**。そのまま残しておいても害はありませんが、新しく書き加える必要はありません。

## 11. まるごとサンプル（実際のファイル）

ここでは、実際に使われているファイルを**1枚まるごと**載せます。コピーして自分の環境に合わせて値を変えれば、そのまま動くサンプルとして使えます（保存は Shift_JIS・改行 CRLF）。

### 11-1. 検索画面（`ui\検索\M-08.txt`）

キーワード欄・フォーマット絞り込み欄・検索/条件クリア/画面更新ボタン・結果一覧（[GRID]）を持つ、検索ブックの検索画面の全体です。

```
[SHEET]
SheetName=ナレッジ検索
SheetTitle=ナレッジ検索
Visible=visible
TabColor=―（色を付けない）
HideColumnsFrom=G
===
[STYLE]
BaseFont=メイリオ
BaseFontSize=10
BaseFontColor=濃いグレー
SheetBackColor=白
BaseAlignH=left
BaseAlignV=center
BaseWrap=off
BaseNumberFormat=standard
===
[SCREEN]
Width=7（列数）
Height=auto
Title=ナレッジ検索
BackColor=#1F3864
FontColor=#FFFFFF
EndCol=G
BarHeight=30
===
[FREEZE]
Enabled=on
Rows=13
Cols=0
===
[COLUMN]
ColumnWidth_A=6
ColumnWidth_B=20
ColumnWidth_C=24
ColumnWidth_D=24
ColumnWidth_E=24
ColumnWidth_F=17
===
[SUBHEADER]
Title=条件で探す
Cell=A3
BackColor=#2F5597
FontColor=#FFFFFF
EndCol=G
Height=22
===
[NOTE]
Cell=A4:G7
Text=キーワードを入れて「検索」を押すと、データの全ナレッジから一致するものを下の一覧に表示します。\nフォーマットで絞り込むこともできます。「条件クリア」で入力を空に戻します。\n検索結果の行のどこでもダブルクリックすると、その内容がナレッジ表示画面で開きます。
BackColor=#DEEBF7
Border=thin
===
[LABEL]
Cell=B8
Text=キーワード
===
[INPUT]
Cell=C8:E8
inputDataKey=keyword
InputType=text
MaxLength=200
InputColor=#FFFFFF
Border=thin
===
[LABEL]
Cell=B10
Text=フォーマット
===
[INPUT]
Cell=C10:E10
inputDataKey=formatId
InputType=dropdown
DropdownSource=format:list
InputColor=#FFFFFF
Border=thin
===
[BUTTON]
Cell=B11
Text=検索
OnClick=Btn_SearchV23
Width=110
Height=22
BackColor=#4472C4
ForeColor=#FFFFFF
===
[BUTTON]
Cell=D11
Text=条件クリア
OnClick=Btn_SearchClearV23
Width=110
Height=22
BackColor=#4472C4
ForeColor=#FFFFFF
===
[BUTTON]
Cell=F1
Text=画面更新
OnClick=Btn_RefreshAllSheets
Width=120
Height=22
BackColor=#D9E1F2
ForeColor=#000000
===
[SUBHEADER]
Title=検索結果
Cell=A12
BackColor=#2F5597
FontColor=#FFFFFF
EndCol=G
Height=22
===
[GRID]
Name=gridResult
StartCell=A13
HeaderRow=13
InitialRows=10
Columns=No,formatId,searchField1,searchField2,searchField3,updatedAt,knowledgeNo
ColumnHeaders=No,フォーマット,検索1,検索2,検索3,更新日,ナレッジ番号
HeaderBackColor=#D9E1F2
FreezeHeader=FALSE
===
```

### 11-2. フォーマット一覧画面（`ui\管理\M-02.txt`）

操作ボタン（編集・プレビュー・無効化・削除・リロード）と説明、登録済みフォーマットの一覧（[GRID]）を持つ、管理ブックの画面です。

```
[SHEET]
SheetName=フォーマット一覧
SheetTitle=フォーマット一覧
Visible=visible
TabColor=―（色を付けない）
HideColumnsFrom=I
===
[STYLE]
BaseFont=メイリオ
BaseFontSize=10
BaseFontColor=濃いグレー
SheetBackColor=白
BaseAlignH=left
BaseAlignV=center
BaseWrap=off
BaseNumberFormat=standard
===
[SCREEN]
Width=10（列数）
Height=auto
Title=フォーマット一覧
BackColor=#1F3864
FontColor=#FFFFFF
EndCol=J
BarHeight=30
===
[FREEZE]
Enabled=on
Rows=16
Cols=0
===
[COLUMN]
ColumnWidth_A=6
ColumnWidth_B=5
ColumnWidth_C=14
ColumnWidth_D=22
ColumnWidth_E=12
ColumnWidth_F=12
ColumnWidth_G=18
ColumnWidth_H=12
ColumnWidth_I=2
ColumnWidth_J=14
===
[SUBHEADER]
Title=操作
Cell=A3
BackColor=#2F5597
FontColor=#FFFFFF
EndCol=J
Height=22
===
[BUTTON]
Cell=B4
Text=編集
OnClick=Btn_EditFormat
Width=90
Height=22
BackColor=#4472C4
ForeColor=#FFFFFF
===
[BUTTON]
Cell=D4
Text=プレビュー
OnClick=Btn_PreviewFormat
Width=90
Height=22
BackColor=#4472C4
ForeColor=#FFFFFF
===
[BUTTON]
Cell=F4
Text=無効化
OnClick=Btn_DisableFormat
Width=90
Height=22
BackColor=#4472C4
ForeColor=#FFFFFF
===
[BUTTON]
Cell=H4
Text=削除
OnClick=Btn_DeleteFormat
Width=90
Height=22
BackColor=#4472C4
ForeColor=#FFFFFF
===
[SUBHEADER]
Title=操作の説明
Cell=A6
BackColor=#2F5597
FontColor=#FFFFFF
EndCol=J
Height=22
===
[NOTE]
Cell=A7:K12
Text=各ボタンは、下の一覧でチェックを入れた行のフォーマットに働きます。\n編集 … そのフォーマットの内容を変更します。\nプレビュー … そのフォーマットの見た目を確認します。\n無効化 … 押すたびに「無効」と「有効」を切り替えます（無効にすると一覧で灰色になり、新しい登録に使えなくなります）。\n削除 … そのフォーマットを完全に消します（そのフォーマットを使うナレッジが残っていると消せません）。\nボタンを押す前に、一覧で対象の行にチェックを入れてください。
BackColor=#DEEBF7
Border=thin
===
[SUBHEADER]
Title=登録済みフォーマット
Cell=A13
BackColor=#2F5597
FontColor=#FFFFFF
EndCol=J
Height=22
===
[BUTTON]
Cell=B14
Text=リロード
OnClick=Btn_ReloadFormats
Width=90
Height=22
BackColor=#4472C4
ForeColor=#FFFFFF
===
[BUTTON]
Cell=I1
Text=画面更新
OnClick=Btn_RefreshAllSheets
Width=120
Height=22
BackColor=#D9E1F2
ForeColor=#000000
===
[GRID]
Name=gridFormatList
StartCell=A16
HeaderRow=16
Columns=Selected,No,FormatID,FormatName,FieldCount,KnowledgeCount,UpdatedAt,Status
ColumnHeaders=選択,No,FormatID,フォーマット名,フィールド数,ナレッジ数,更新日,状態
FreezeHeader=FALSE
HeaderBackColor=#D9E1F2
InitialRows=10
===
```

### 11-3. 登録フォームの定義（`ui\検索\M-05.txt`）

入力フォームの寸法・文言を決める `[USERFORM]` を含む、検索ブックの登録画面の全体です。

```
[SHEET]
SheetName=ナレッジ登録
SheetTitle=ナレッジ登録
Visible=visible
TabColor=―（色を付けない）
HideColumnsFrom=H
===
[STYLE]
BaseFont=メイリオ
BaseFontSize=10
BaseFontColor=濃いグレー
SheetBackColor=白
BaseAlignH=left
BaseAlignV=center
BaseWrap=off
BaseNumberFormat=standard
===
[SCREEN]
Width=7（列数）
Height=auto
Title=ナレッジ登録
BackColor=#1F3864
FontColor=#FFFFFF
EndCol=G
BarHeight=30
===
[USERFORM]
formWidth=486
formHeight=0
fontName=メイリオ
fontSize=10
labelWidth=107
rowHeightSingle=48
rowHeightMulti=114
rowHeightMultiLong=129
margin=18
buttonBarHeight=48
caption=ナレッジ登録
multiLineScrollBars=vertical
formatSelectorType=dropdown
headerFields=knowledgeId
formatHelp=登録するナレッジのフォーマットを選びます。選ぶと下に項目が並びます。
headerHelp_knowledgeId=登録時に自動で割り当てられます。空き番号は手で入れ直すこともできます。
===
[BUTTON]
Cell=F1
Text=画面更新
OnClick=Btn_RefreshAllSheets
Width=120
Height=22
BackColor=#D9E1F2
ForeColor=#000000
===
```

### 11-4. フォーマット定義（`formats\FAULT.txt`）

「障害対応」型の完全な定義です。単一行・日付・選択・複数行のすべての型を含みます。

```
[FORMAT]
FormatId=FAULT
FormatName=障害対応
===
[FIELD]
FieldName=件名
FieldType=単一行
Required=TRUE
MaxLength=120
fieldPlaceholder=（例）経理システムへログインできない
searchTarget=TRUE
===
[FIELD]
FieldName=発生日時
FieldType=日付
Required=TRUE
fieldPlaceholder=（例）2026-06-09 10:30
searchTarget=TRUE
===
[FIELD]
FieldName=担当者
FieldType=単一行
Required=FALSE
MaxLength=40
fieldPlaceholder=（例）山田 太郎
searchTarget=TRUE
===
[FIELD]
FieldName=カテゴリ
FieldType=選択
Required=FALSE
DropdownOptions=アプリ障害|インフラ障害|ネットワーク障害|セキュリティ
fieldPlaceholder=（例）アプリ障害
searchTarget=FALSE
===
[FIELD]
FieldName=優先度
FieldType=選択
Required=FALSE
DropdownOptions=低|中|高|緊急
fieldPlaceholder=（例）高
searchTarget=FALSE
===
[FIELD]
FieldName=概要
FieldType=複数行
Required=TRUE
Rows=4
fieldPlaceholder=（例）応答時間が通常の10倍以上に増加
searchTarget=FALSE
===
[FIELD]
FieldName=原因
FieldType=複数行
Required=FALSE
Rows=4
fieldPlaceholder=（例）長時間トランザクションの累積
searchTarget=FALSE
===
```
