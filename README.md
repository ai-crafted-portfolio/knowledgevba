# knowledgevba

ナレッジ管理ツール (VBA 製) のドキュメントサイト（mkdocs Material）。

- 想定公開 URL: `https://<github-user>.github.io/knowledgevba/`
- 全ページ `noindex,nofollow` で検索エンジン indexing 抑止（職場利用ツールのため）

## ローカル確認

```bash
pip install mkdocs-material
mkdocs serve   # http://localhost:8000
mkdocs build   # ./site/ にビルド
```

## ページ

- `docs/index.md` — 概要
- `docs/spec.md` — 仕様
- `docs/operations.md` — 操作手順
- `docs/architecture.md` — アーキテクチャ

## デプロイ

`main` への push で GitHub Actions（`.github/workflows/deploy.yml`）が `mkdocs gh-deploy --force` を走らせて `gh-pages` ブランチに反映。GitHub repo 側で Pages を `gh-pages` branch に設定済であること前提。
