---
title: ナレッジ管理ツール (VBA 製)
description: Excel VBA で実装したナレッジ管理ツールの仕様・操作・アーキテクチャまとめ
tags:
  - excel
  - vba
  - knowledge-base
---

# ナレッジ管理ツール (VBA 製)

PC 上にノウハウ・トラブル対応メモ・手順書をテキストファイル（`.txt`、Shift_JIS + CRLF）で蓄積し、Excel から横断検索・サムネ画像つきプレビューできるようにするツールです。現行リリースは **release v2 + image_ext rev1**（2026-05-04）。

## このツールで解決したいこと

- 過去の障害メモや手順を「あのとき書いたはずだ」で終わらせない
- 子プロセス起動禁止のセキュリティポリシーで動かない外部ツールを諦め、**Excel プロセス内で完結**する

## 想定読者

- 同ツールを PC に展開するメンバー
- VBA モジュール構成・検索スコアリング設計をレビューする人

## 全体像

```mermaid
flowchart LR
    User([利用者]) -->|キーワード| MainSheet[メインシート]
    MainSheet --> SearchSheet[検索シート]
    SearchSheet -->|検索実行| Engine[clsSearchEngine]
    Engine -->|.txt 走査| Data[(dataFolder /<br/>*.txt SJIS+CRLF)]
    Engine -->|サムネ取得| Images[(kb_images /<br/>KnwNo.png)]
    Engine -->|Score 降順| Result[結果一覧<br/>サムネ + Score 列]
    Result -->|行選択 + Macro| Form[spec 駆動 UserForm<br/>詳細プレビュー]
```

## 主要機能

ナレッジ一覧でメタデータを管理し、検索画面でキーワード一致をスコアリング（タイトル ×3 + 対象フィールド ×2 + 出現回数）して降順ソート、結果一覧の H 列にサムネ（60×40px）、I 列に Score を表示します。結果行を選択して `Macro_ShowSearchResultPreview` を実行すると、`clsFormSpec` で宣言した DSL から `modFormBuilder` が UserForm を実行時生成して詳細プレビューを表示します。設定画面で `dataFolder` を切り替えれば複数ナレッジ集を運用でき、`<dataFolder の親>/kb_images/<KnwNo>.png` の規約で画像が解決されます。`ImagePath=` スタンザを `.txt` 末尾に書けばパス上書きも可能です。

## 設計方針のハイライト

- **モジュール配布 + セルフセットアップ** — `.xlsm` 本体は配布せず、`.bas`/`.cls` 24 個 + セットアップマクロ 1 回実行という形で各 PC に展開
- **VBA 子プロセス全面禁止** — `Shell` / `WScript.Shell.Run` 等は使わず、すべて Excel プロセス内で完結
- **層分離** — エントリポイント / ビジネスロジック / ユーティリティ / インストーラ / 特殊 の 5 層で責務を分離。詳細は [アーキテクチャ](architecture.md) 参照

## ページ構成

- [概要](index.md)（このページ）
- [仕様](spec.md) — モジュール 24 個、検索スコアリング、画像解決、`clsFormSpec` DSL、テスト構成
- [操作マニュアル](operations.md) — 各シートの操作手順、ボタン名、画面遷移
- [アーキテクチャ](architecture.md) — 層分離図、依存関係、配布パターン

---

## TODO・制約・既知の限界

### 制約

- VBA 子プロセス禁止（Shell/Run/WScript.Shell/Exec）— 職場 PC ポリシー (ADR-0002)
- クラスモジュール (.cls) 内の `Public Const/Type/Declare/Static` 禁止 (ADR-0027)
- aspose-cells-python 単独では VBA binary stub のみ生成、real Excel COM 必須 (ADR-0026)
- mkdocs Material のモバイル UX が完璧ではない（ナビ collapse は OK、図解の細部は要確認）

### 既知の限界

- Excel 単体動作のため Web/モバイルでは利用不可
- 同時編集なし（ファイルベース、Git 等での並行編集が必要）
- 検索性能はモジュール数 O(n) 線形（数千ナレッジまでは実用）

### TODO（v15 以降のロードマップ）

- M-5: modImageRender の RowHeight 副作用排除
- D-3: clsKnowledgeManager / clsSearchEngine / clsTaskController に Worksheet DI 追加
- D-5: 責務固有の Const をクラス側へ移動
- Minor m-1: SHEET_* / FIELD_TYPE_* の Enum 化
- ベンチマーク取得（A+ 到達条件）
- 単体テストカバレッジ整備（A+ 到達条件）

詳細は ADR ([0001-0033](https://github.com/ai-crafted-portfolio/knowledgevba)) を参照。
