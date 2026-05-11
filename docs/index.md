---
title: ナレッジ管理ツール (VBA 製)
description: Excel VBA で実装したナレッジ管理ツールの仕様・操作・アーキテクチャまとめ
tags:
  - excel
  - vba
  - knowledge-base
---

# ナレッジ管理ツール (VBA 製)

ノウハウ・トラブル対応メモ・手順書をテキストファイル（`.txt`、Shift_JIS + CRLF 改行）で蓄積し、Excel から横断検索・サムネ画像つきプレビューできるようにする内製ツールです。

!!! tip "本ドキュメントで使う用語"
    - **VBA** — Excel に標準搭載されているプログラミング言語（Visual Basic for Applications）
    - **VBE** — VBA のコードを書くエディタ（Visual Basic Editor）。Excel 上で `Alt+F11` で開きます
    - **`mod` / `cls` プレフィックス** — VBA 慣習で標準モジュール（手続き集）と クラスモジュール（オブジェクト指向）を区別
    - **ThisWorkbook** — Excel が `.xlsm` ごとに自動で持つ特殊モジュール。`Workbook_Open` 等の Excel イベントを受け取ります
    - **dataFolder** — ナレッジ `.txt` ファイルを置くフォルダ。設定シートで切り替え可能

## このツールで解決したいこと

- 過去の障害メモや手順を「あのとき書いたはずだ」で終わらせない
- 単純なキーワードヒット数では拾いきれない、タイトル一致や対象フィールド一致のブーストを反映した順位付け
- 子プロセス起動を伴わない、Excel プロセス内で完結する検索 UI を提供する

## 想定読者

- 同ツールを利用するメンバー
- VBA モジュール構成・検索スコアリング設計を確認したい人
- 利用者向け操作手順を確認したい人 → [操作手順](operations.md) へ

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

- **モジュール配布 + セルフセットアップ** — `.xlsm` 本体は配布せず、`.bas`/`.cls` 48 個 + セットアップマクロ 1 回実行という形で各 PC に展開
- **層分離** — エントリポイント / ビジネスロジック / ユーティリティ / 画面 / インストーラ / 特殊 の 6 層で責務を分離。詳細は [アーキテクチャ](architecture.md) 参照
- **画面 spec 駆動描画** — 各業務画面（M-01〜M-14）は `clsScreenSpec` の宣言情報を `clsSheetRenderer` が物理配置することで、画面追加時の差分を局所化
- **検索バックエンド差替境界** — `clsSearchEngine` 1 ファイルだけ差し替えれば txt 走査 → 別データソースへの切替が可能な構造

## ページ構成

- [概要](index.md)（このページ）
- [仕様](spec.md) — モジュール 48 本 + ThisWorkbook の構成、検索スコアリング、画像解決、`clsFormSpec` DSL、テスト構成
- [セットアップ](setup.md) — マクロ有効化、新規 `.xlsm` 作成、モジュール手動 import、初回セットアップマクロ実行、トラブルシューティング
- [操作手順](operations.md) — 検索 / 詳細表示 / ナレッジ登録 / 画像添付 / ログ確認 等、日常操作の手順
- [アーキテクチャ](architecture.md) — 層分離図、依存関係、配布パターン
- [ソースコード一覧](source/index.md) — 全 48 モジュール + ThisWorkbook の本文をコピペ可能な形で公開

## 既知の制約

- **VBA 子プロセス起動は使わない** — `Shell` / `WScript.Shell.Run` 等は使わず、外部ツール連携が必要な機能はすべて Excel プロセス内で代替実装しています
- **モバイル / Excel for Mac 不可** — `Forms.*` ProgID が無いため `Macro_ShowSearchResultPreview` は Windows 版 Excel でのみ動作します（[セットアップ §0](setup.md#0-前提環境)）
- **クラスモジュール内 `Public Const` 禁止** — 実 Excel が reject する仕様への対応として、共通定数は `modCommon.bas` に集約しています
