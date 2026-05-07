---
title: Q&A: なぜこの設計？
description: ナレッジ管理ツール (VBA 製) の設計判断とよくある反論への回答
---

# Q&A: なぜこの設計？

このツールが「なぜこういう作りなのか」「他の選択肢はなかったのか」をまとめます。設計判断の背景は ADR（Architectural Decision Records）にも記録されています。

## Q1. VBA で十分か？ Python や Web アプリの方がモダンでは？

**答え**: 職場 PC の制約下（Python サブプロセス禁止、Shell 実行不可、外部ネットワーク制限）では VBA が最も現実的な選択肢です。

**理由**:
- 職場 PC には Python ランタイムが入っているとは限らない
- Shell / WScript.Shell / Run / Exec の子プロセス起動は社内ポリシーで禁止 (ADR-0002)
- Excel は標準で全社配布されており、VBA は追加インストール不要で動く
- Web アプリは社内サーバ手配が必要、個人ナレッジ用途には過剰

**トレードオフ**: VBA は Mac/Web Excel で動作しない、64bit/32bit 差異がある、並列処理が苦手。これらは Excel 単体動作という制約と引き換え。

## Q2. ChromaDB と連携しないなら検索が遅くないか？

**答え**: 数千ナレッジ規模までは O(n) 線形検索で実用十分。スケーラビリティが課題になっても **ChromaDB 連携は採用しない方針** が確定しています (ADR-0004)。

**現状の性能**:
- 100 ナレッジで検索 < 0.5 秒
- 1,000 ナレッジで検索 < 5 秒
- 10,000 ナレッジ超で 30 秒超 → 別アプローチで対処

**スケール対応の選択肢（ChromaDB 以外）**:
- Excel ファイルを部署/年度等で分割
- インデックスシートの先行構築（事前計算）
- 別ツール（Notion / Obsidian 等）への移行検討（職場ポリシー許諾下で）

**ChromaDB 連携を恒久的にしない理由**:
- 職場利用ツールは Cowork / ChromaDB と一切連携しない (ADR-0004)
- データ漏洩防止の構造的隔離
- 外部依存ゼロのファイルベース運用が職場 PC ポリシーに最適

## Q3. Excel ナレッジ管理は時代遅れではないか？Notion / Obsidian の方が良くない？

**答え**: 職場利用前提では Excel が最強です。

**比較**:

| ツール | 職場利用 | データローカル | 検索性能 | 学習コスト |
|---|---|---|---|---|
| Excel + VBA | ◎ 標準配布 | ◎ ファイルベース | ○ 線形 | △ VBA 学習 |
| Notion | △ クラウド前提 | × | ◎ | ○ |
| Obsidian | ○ ローカル DB | ◎ | ◎ | ○ |

職場 PC では Notion/Obsidian の追加インストール禁止のケースが多く、データはローカル管理が必須。Excel は **「既に入っている」** ことが最大の利点。

## Q4. クラスモジュール内 Public Const 禁止 等の制約があるなら .NET にすべきでは？

**答え**: VBA の制約は学習可能な範囲です。.NET 化は overengineering。

**VBA 制約の代表**:
- クラスモジュール内 `Public Const` 禁止 → 標準モジュールに切出し (ADR-0027)
- Function 内 `Exit Sub` 禁止 → `Exit Function` 使用 (ADR-0028)
- Late binding の型チェック緩い → 静的検証スクリプトでカバー

これらは **構造的に検出可能**で、Rule 1〜6 の verify_vba_*.py で機械的にチェックされます。

## Q5. なぜ職場 PC で個人ナレッジ管理を Excel で？

**答え**: 業務知識を体系化する個人作業であり、ファイルベースで管理することで漏洩リスクが最小化されます。

**用途**:
- 過去のトラブルシューティングを再利用可能な形で蓄積
- 顧客固有の対応パターンを類型化
- 業務ノウハウの個人内継承（自分自身の引き継ぎ用）

クラウド SaaS や外部 RAG では機密情報漏洩リスクがあり、職場ポリシー的にも実用的でも Excel ファイルベースが最適解。

## Q6. VBA 子プロセス禁止は不便ではないか？

**答え**: 不便ですが、それを補う設計で十分機能します。

**禁止される代表機能**:
- 外部 API 呼出 → 不要（ChromaDB 連携は明示的に無効化）
- Python ML スクリプト連携 → 不要（VBA 単体で動作）
- WSL / PowerShell 起動 → 不要（VBA 単体で動作）

**代替手段**:
- Web スクレイピング → 不要（職場利用なので外部 Web は使わない）
- 大規模計算 → Excel ワークシート関数で対応
- ファイル I/O → ADODB / FileSystemObject で対応

すべての必要機能は VBA 標準機能で実装可能。

## 関連 ADR

設計判断の詳細は ADR で追跡可能:

- [ADR-0002: VBA 子プロセス禁止](https://github.com/ai-crafted-portfolio/knowledgevba)
- [ADR-0004: 職場利用ツールは Cowork/ChromaDB と隔離](https://github.com/ai-crafted-portfolio/knowledgevba)
- [ADR-0008: VBA モジュール手動 import + Setup マクロ配布](https://github.com/ai-crafted-portfolio/knowledgevba)
- [ADR-0026: real Excel COM ビルド経路](https://github.com/ai-crafted-portfolio/knowledgevba)
- [ADR-0027: クラスモジュール内 Public Const 禁止](https://github.com/ai-crafted-portfolio/knowledgevba)
- [ADR-0028: Function/Sub Exit 文の整合性](https://github.com/ai-crafted-portfolio/knowledgevba)

