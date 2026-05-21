# knowledgevba

ナレッジ管理ツール (VBA 製) v2.1 の公開ドキュメントサイト（MkDocs Material 製）。

- 想定公開 URL: `https://ai-crafted-portfolio.github.io/knowledgevba/`
- 全ページに `noindex,nofollow` を付与

## ローカル確認

```bash
pip install mkdocs mkdocs-material
mkdocs serve   # http://localhost:8000
mkdocs build   # ./site/ にビルド
```

## 構成

- `docs/` — ドキュメント原稿（Markdown）
- `docs/source/` — VBA モジュールのソースページ（`tools/gen_source_pages.py` が生成）
- `mkdocs.yml` — サイト設定とナビゲーション
- `overrides/` — テーマの上書き
- `tools/gen_source_pages.py` — ソースページ生成スクリプト

## ページ

- `docs/index.md` — 概要
- `docs/architecture.md` — 全体アーキテクチャ
- `docs