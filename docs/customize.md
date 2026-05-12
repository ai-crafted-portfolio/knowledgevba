# 外観のカスタマイズ

ボタンの大きさやフォントサイズ、タイトル帯の高さなどの「見た目」を、
すべて 1 つのモジュール (`modUIConfig`) で集中管理しています。
ここを書き換えるだけで全画面 (メイン / フォーマット / ナレッジ / 検索 など 14 画面) に
一斉に反映されるので、シート毎の手作業は不要です。

---

## カスタマイズの考え方

- `modUIConfig` が **唯一の設定の正** です。他のモジュールはこの値を参照するだけです。
- 値を書き換えた後、Excel ブックを **保存 → 閉じる → もう一度開く** と、
  次回画面表示時に新しい設定で再描画されます。
  (起動中の画面はその場では変わりません。各画面のシートを開き直すか、
  メインの「ナレッジ登録 / 検索 / ナレッジ一覧」などボタンで遷移するとそのシートが再描画されます。)
- 値を間違えてしまっても、もとの値に戻して保存 → 再オープンするだけで元通りです。

---

## 変数一覧

| 変数名 | 役割 | 既定値 | 変更例 |
|---|---|---|---|
| `UI_BTN_WIDTH_PT` | ボタンの最小幅 (ポイント) | 110 | キャプションが長い場合は 130 などに増やす |
| `UI_BTN_HEIGHT_PT` | ボタンの最小高さ (ポイント) | 26 | フォントを大きくしたら 30〜32 に増やす |
| `UI_BTN_CELL_SPAN` | ボタンが横方向に占めるセル数 | 2 | 1 にすると単独セル幅のみ。3 にするとさらに横に長いボタン |
| `UI_BTN_FONT_SIZE` | ボタン文字の大きさ (ポイント) | 10 | 視認性を上げるなら 11〜12 |
| `UI_BTN_CAPTION_PREFIX` | 全ボタン名の頭につく記号 | `""` (なし) | `ChrW(&H25B6) & " "` で「▶ 」、`"> "` で「> 」 |
| `UI_BTN_ROW_HEIGHT_PT` | ボタンが置かれる行の高さ | 28 | ボタンを大きくしたら同じだけ増やす |
| `UI_HINT_ROW_HEIGHT_PT` | ボタン直下の説明文の行の高さ | 32 | 説明文が長くて見切れる場合は 40 などに |
| `UI_TITLE_FONT_SIZE` | 画面タイトル帯の文字サイズ | 14 | 16〜18 にすると太く強調 |
| `UI_TITLE_ROW_HEIGHT` | 画面タイトル帯の行の高さ | 28 | フォントを上げたら 32〜36 に |
| `UI_SECTION_FONT_SIZE` | セクション帯 (`■ ラベル` 等) の文字サイズ | 11 | 12〜13 で見出しを強める |
| `UI_SECTION_ROW_HEIGHT` | セクション帯の行の高さ | 20 | フォントを上げたら 24 等に |
| `UI_LABEL_FONT_SIZE` | 入力欄ラベル (項目名) の文字サイズ | 10 | 11 で読みやすく |
| `UI_HINT_FONT_SIZE` | 説明文・型表示の文字サイズ | 9 | 10 で読みやすく |
| `UI_HINT_CELL_WRAP` | 説明文セルで折り返すか | `True` | `False` にすると 1 行のみ表示 |

---

## 手順

1. Excel ブックを開く。
2. キーボードで **Alt + F11** を押す (Visual Basic Editor が開く)。
3. 左側のプロジェクトツリーから
   **VBAProject → 標準モジュール → modUIConfig** をダブルクリックする。
4. 変更したい変数の右辺の値を書き換える。
   例: ボタンを少し大きくしたい場合 →
   ```vb
   Public Const UI_BTN_WIDTH_PT As Double = 130#
   Public Const UI_BTN_HEIGHT_PT As Double = 30#
   Public Const UI_BTN_ROW_HEIGHT_PT As Double = 32#
   ```
5. **Ctrl + S** で保存。
6. VBE を閉じる (右上 ×)。Excel に戻る。
7. Excel ブック自体を **上書き保存 → 一度閉じる → もう一度開く**。
8. メイン画面で各タスクボタンをクリックすると、対応シートが新しい設定で再描画される。

---

## よくあるカスタマイズ例

### 1. ボタン名の頭に「▶」を付けたい

`modUIConfig` の `UI_BTN_CAPTION_PREFIX` を次のように変更します。

```vb
Public Const UI_BTN_CAPTION_PREFIX As String = ChrW(&H25B6) & " "
```

`ChrW(&H25B6)` は「▶」の文字コードです。直接 `"▶ "` と書いてもよいですが、
VBE の保存時にエンコーディングで化けることがあるため `ChrW` 表記が安全です。

### 2. ボタン名が長くてはみ出す場合

ボタン幅を増やすか、セル span を増やします。

```vb
Public Const UI_BTN_WIDTH_PT As Double = 140#
Public Const UI_BTN_CELL_SPAN As Long = 3
```

### 3. 老眼に優しく全体的に大きくしたい

文字サイズと行高を一段階ずつ上げます。

```vb
Public Const UI_BTN_FONT_SIZE As Double = 12#
Public Const UI_BTN_HEIGHT_PT As Double = 30#
Public Const UI_BTN_ROW_HEIGHT_PT As Double = 34#
Public Const UI_HINT_ROW_HEIGHT_PT As Double = 38#
Public Const UI_TITLE_FONT_SIZE As Double = 16#
Public Const UI_TITLE_ROW_HEIGHT As Double = 32#
Public Const UI_SECTION_FONT_SIZE As Double = 12#
Public Const UI_LABEL_FONT_SIZE As Double = 11#
Public Const UI_HINT_FONT_SIZE As Double = 10#
```

### 4. 印刷を意識して全体を控えめにしたい

逆に小さめに揃えます。

```vb
Public Const UI_BTN_WIDTH_PT As Double = 90#
Public Const UI_BTN_HEIGHT_PT As Double = 22#
Public Const UI_BTN_ROW_HEIGHT_PT As Double = 24#
Public Const UI_HINT_ROW_HEIGHT_PT As Double = 28#
Public Const UI_BTN_FONT_SIZE As Double = 9#
```

---

## 元に戻したい場合

各変数の既定値は、本ページの「変数一覧」の **既定値** 列に載っています。
書き換えた値を既定値に戻してブックを再オープンすれば、初期状態に戻ります。
