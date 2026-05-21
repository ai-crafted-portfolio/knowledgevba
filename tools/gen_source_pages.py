#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""knowledgevba v2.1 ソースコードページ生成スクリプト。

現行 v2.1 の canonical VBA モジュール一式から、ドキュメントサイトの
``docs/source/`` 配下のページを一括生成する。

- 入力: ``<repo>/.tmp_upload/_canon_v21/vba/`` の .bas / .cls（canonical 実装）
- 出力: ``<repo>/docs/source/<層slug>/<モジュール名>.md`` + ``docs/source/index.md``

仕様:
- 配布対象（非テスト）モジュールだけをページ化する。名前に Test を含む
  モジュール（modTestHelpers / clsTestSearchExecutor / mod*TestRunner 等）は
  テスト用のため除外する。
- VBA ソースは CP932（Shift_JIS）で読み込む。CP932 で読めない場合のみ
  UTF-8 を試し、どちらも失敗したら CP932 + 置換で読み込む。
- 各ページにはソース全文をコピー可能なコードブロックで掲載する。
  ダウンロードリンクは一切貼らない。

使い方:
    python tools/gen_source_pages.py
スクリプトは自身の位置からリポジトリルートを解決するため、引数は不要。
"""
from __future__ import annotations

import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_ROOT = os.path.join(REPO_ROOT, ".tmp_upload", "_canon_v21", "vba")
OUT_ROOT = os.path.join(REPO_ROOT, "docs", "source")

# ------------------------------------------------------------------
# 層の定義（表示順）。slug = 出力サブフォルダ名。
# ------------------------------------------------------------------
LAYERS = [
    ("infrastructure", "インストーラ層",
     "空のブックに必要なシートを用意するセットアップ入口です。"),
    ("entrypoint", "エントリポイント層",
     "シート上のボタン操作や起動イベントを受け取り、業務ロジックへ橋渡しする"
     "薄い受け口です。"),
    ("screen", "画面層",
     "M-02〜M-14 の各画面の組み立てと再描画を担うクラス群です。"
     "レイアウトの実体は UI 定義ファイルにあり、ここでは描画の指揮だけを行います。"),
    ("business-logic", "ビジネスロジック層",
     "ナレッジ・フォーマット・検索・ログなど、ツールの中心的な業務処理を担う層です。"),
    ("utility", "ユーティリティ層",
     "外部ファイルの入出力、スタンザ解析、文字列・日付処理など、"
     "上位の層から共通利用される下支えの層です。"),
    ("special", "特殊モジュール",
     "Excel がブックごとに自動で持つ ThisWorkbook モジュールです。"
     "起動・終了イベントを受け取ります。"),
]
LAYER_JP = {slug: jp for slug, jp, _ in LAYERS}

# ------------------------------------------------------------------
# モジュール → 層 slug の割り当て（配布対象 67 モジュール）。
# ------------------------------------------------------------------
LAYER_MEMBERS = {
    "infrastructure": [
        "modSetup",
    ],
    "entrypoint": [
        "modEntryKnowledge", "modEntrySearch", "modEntryFormat",
        "modEntrySettings", "modSpecExamples",
    ],
    "screen": [
        "clsFormatListScreen", "clsFormatDesignScreen", "clsFormatPreviewScreen",
        "clsKnowledgeRegisterScreen", "clsKnowledgeEditScreen",
        "clsKnowledgeListScreen", "clsSearchScreen", "clsKnowledgeViewScreen",
        "clsStorageScreen", "clsSystemSettingsScreen", "clsMigrationScreen",
        "clsFileFormatScreen", "clsLogScreen",
        "clsErrorBandScreen", "clsBackupMgmtScreen",
    ],
    "business-logic": [
        "clsKnowledgeManager", "clsFormatManager", "clsFieldMigrator",
        "clsSearchEngine", "clsLogger", "clsStorageResolver",
        "clsSetupOrchestrator", "clsSheetRenderer", "clsUserFormRenderer",
        "IScreenRenderer", "modScreenRender", "modFactory", "modFormBuilder",
        "clsButtonSpec", "clsControlSpec", "clsFieldSpec", "clsFormSpec",
        "clsScreenSpec", "clsSectionSpec",
    ],
    "utility": [
        "modConfigLoader", "modConfigHolder", "modStanzaIO", "modUILoader",
        "modFormatLoader", "modKnowledgeFileIO", "modFileIO", "modCommon",
        "modStringUtil", "modDateUtil", "modImageRender", "modLogIds",
        "modSheetMap", "modUIConfig", "modFormatColumns", "modUILayoutExtractor",
        "clsLogEntry", "ClsStanzaSection", "ClsStanzaValidationIssue",
        "ClsStanzaValidationResult",
        "clsCellBinding", "clsCellAddrHelper", "clsCellIO", "clsGridIO",
    ],
    "special": [
        "ThisWorkbook_管理", "ThisWorkbook_検索", "ThisWorkbook_登録修正",
    ],
}
MODULE_LAYER = {m: slug for slug, mods in LAYER_MEMBERS.items() for m in mods}

# ------------------------------------------------------------------
# モジュール → 配置ブック。未登録は「3 ブック共通」。
# ------------------------------------------------------------------
BOOK_MAP = {
    # 登録修正.xlsm 専用
    "modEntryKnowledge": "登録修正.xlsm",
    "clsKnowledgeRegisterScreen": "登録修正.xlsm",
    "clsKnowledgeEditScreen": "登録修正.xlsm",
    "ThisWorkbook_登録修正": "登録修正.xlsm",
    # 検索.xlsm 専用
    "modEntrySearch": "検索.xlsm",
    "clsKnowledgeListScreen": "検索.xlsm",
    "clsSearchScreen": "検索.xlsm",
    "clsKnowledgeViewScreen": "検索.xlsm",
    "ThisWorkbook_検索": "検索.xlsm",
    # 管理.xlsm 専用
    "modEntryFormat": "管理.xlsm",
    "modEntrySettings": "管理.xlsm",
    "clsFormatListScreen": "管理.xlsm",
    "clsFormatDesignScreen": "管理.xlsm",
    "clsFormatPreviewScreen": "管理.xlsm",
    "clsStorageScreen": "管理.xlsm",
    "clsSystemSettingsScreen": "管理.xlsm",
    "clsMigrationScreen": "管理.xlsm",
    "clsFileFormatScreen": "管理.xlsm",
    "clsLogScreen": "管理.xlsm",
    "clsErrorBandScreen": "管理.xlsm",
    "clsBackupMgmtScreen": "管理.xlsm",
    "ThisWorkbook_管理": "管理.xlsm",
}
DEFAULT_BOOK = "3 ブック共通"

# ------------------------------------------------------------------
# モジュール → 役割（1 行説明）。canonical 実装のヘッダコメントに基づく。
# ------------------------------------------------------------------
ROLE_MAP = {
    "modSetup": "3 ブック共通のセットアップ入口。各ブックに必要なシート（LOG・各画面）を確認し、無ければ作成する",
    "modEntryKnowledge": "ナレッジ登録・修正画面のボタン処理。入力セルとナレッジデータの相互変換、保存・読込ワークフロー",
    "modEntrySearch": "ナレッジ検索画面のボタン処理。検索条件セルの読み取りと結果グリッドへの書き出し",
    "modEntryFormat": "フォーマット設計画面のボタン処理。フォーマット定義セルとデータの相互変換、保存・読込ワークフロー",
    "modEntrySettings": "格納先設定（M-10）・設定（M-11）画面の入口。設定値のシート反映・再読込・破棄",
    "modSpecExamples": "clsFormSpec を使った UserForm 組み立てのデモ用コード",
    "clsFormatListScreen": "M-02 フォーマット一覧画面の構築・再描画",
    "clsFormatDesignScreen": "M-03 フォーマット設計画面の構築・再描画",
    "clsFormatPreviewScreen": "M-04 フォーマットプレビュー画面の構築・再描画",
    "clsKnowledgeRegisterScreen": "M-05 ナレッジ登録画面の構築・再描画",
    "clsKnowledgeEditScreen": "M-06 ナレッジ修正画面の構築・再描画",
    "clsKnowledgeListScreen": "M-07 ナレッジ一覧画面の構築・再描画",
    "clsSearchScreen": "M-08 ナレッジ検索画面の構築・再描画",
    "clsKnowledgeViewScreen": "M-09 ナレッジ表示画面の構築・再描画",
    "clsStorageScreen": "M-10 格納先設定画面の構築・再描画",
    "clsSystemSettingsScreen": "M-11 設定画面の構築・再描画",
    "clsMigrationScreen": "M-12 フィールド反映画面の構築・再描画",
    "clsFileFormatScreen": "M-13 ファイル形式設定画面の構築・再描画",
    "clsLogScreen": "M-14 操作ログ画面の構築・再描画",
    "clsErrorBandScreen": "M-13 のエラーバンド（4 列）表示を担う連動画面クラス",
    "clsBackupMgmtScreen": "M-14 と同じ系統のバックアップ管理画面クラス",
    "clsKnowledgeManager": "ナレッジの登録・読込・採番・楽観ロックなどの業務ロジック",
    "clsFormatManager": "フォーマット定義の作成・編集・削除と、保存時の既存ナレッジ自動反映の連動",
    "clsFieldMigrator": "フォーマット変更時に既存ナレッジへフィールドを反映し、消失リスク時にバックアップを取る",
    "clsSearchEngine": "ナレッジ検索の中核。番号直接・キーワード・日付範囲フィルタとスコアリング",
    "clsLogger": "操作ログをログシートへ出力。詳細度フィルタと行数上限ローテーション",
    "clsStorageResolver": "格納先設定に基づきファイル参照リンクを解決する",
    "clsSetupOrchestrator": "ブック起動時のセットアップ一括処理（設定読込→ログ初期化→シート構築→保護→起動シート表示）",
    "clsSheetRenderer": "画面描画インターフェースのシート実装。レイアウト適用を modUILoader に委譲する",
    "clsUserFormRenderer": "画面描画インターフェースの UserForm 実装（将来用の入口）",
    "IScreenRenderer": "画面描画の抽象インターフェース（8 メソッド）",
    "modScreenRender": "各画面クラス共通のシート描画入口とログ補助",
    "modFactory": "描画クラス・画面クラスのインスタンス生成を集約するファクトリ",
    "modFormBuilder": "clsFormSpec の宣言情報から UserForm を動的に組み立てる",
    "clsButtonSpec": "ボタン 1 個分の仕様を保持する値オブジェクト",
    "clsControlSpec": "UserForm 上のコントロール 1 個の宣言情報を保持する値オブジェクト",
    "clsFieldSpec": "入力フィールド 1 件の仕様を保持する値オブジェクト",
    "clsFormSpec": "UserForm 1 つの宣言情報を保持する値オブジェクト",
    "clsScreenSpec": "1 画面分の構成情報（タイトル・セクション・ボタン・フィールド）を保持する値オブジェクト",
    "clsSectionSpec": "画面内のセクション帯 1 個の仕様を保持する値オブジェクト",
    "modConfigLoader": "起動時に <ブック名>_config.txt を読み込み modConfigHolder へ渡す（読み取り専用）",
    "modConfigHolder": "設定値をメモリに保持し、各層へ取得用のメソッドで提供する",
    "modStanzaIO": "スタンザ形式（[セクション] ＋ key=value）の汎用 read / write",
    "modUILoader": "UI 定義 .txt を読み込み、シートにレイアウトを適用する",
    "modFormatLoader": "フォーマット定義 .txt の読み書き（書き込みは管理.xlsm のみ）",
    "modKnowledgeFileIO": "ナレッジ .txt の読み書き・バックアップ（書き込みは登録修正.xlsm のみ）",
    "modFileIO": "Shift_JIS ＋ CRLF の低レベルファイル入出力",
    "modCommon": "全ブック共通の定数群",
    "modStringUtil": "文字列処理の純粋関数群",
    "modDateUtil": "日付・時刻処理の純粋関数群",
    "modImageRender": "ナレッジに添付された画像のサムネ・詳細表示",
    "modLogIds": "ログ ID 定数の定義",
    "modSheetMap": "シート名と画面 ID の対応表",
    "modUIConfig": "画面の見た目に関する既定値（フォールバック用）",
    "modFormatColumns": "フォーマット一覧の列番号定数",
    "modUILayoutExtractor": "UI レイアウトを UI 定義 .txt として書き出す開発用ツール",
    "clsLogEntry": "ログ 1 行分の値オブジェクト",
    "ClsStanzaSection": "スタンザ 1 セクションを表す値オブジェクト",
    "ClsStanzaValidationIssue": "スタンザ検証の指摘 1 件を表す値オブジェクト",
    "ClsStanzaValidationResult": "スタンザ検証結果を集約する値オブジェクト",
    "clsCellBinding": "セル 1 個の読み書きを抽象化する薄いラッパー",
    "clsCellAddrHelper": "セル番地の計算（列文字変換・オフセット等）を行うヘルパー",
    "clsCellIO": "セル値の読み書きヘルパー（ワークシート／辞書モックの両対応）",
    "clsGridIO": "グリッド（表）形式の一括読み書きヘルパー",
    "ThisWorkbook_管理": "管理.xlsm 用の ThisWorkbook。起動時の設定読込・セットアップ・終了ログ",
    "ThisWorkbook_検索": "検索.xlsm 用の ThisWorkbook。起動時の設定読込・セットアップ・終了ログ",
    "ThisWorkbook_登録修正": "登録修正.xlsm 用の ThisWorkbook。起動時の設定読込・セットアップ・終了ログ",
}

# ThisWorkbook 系はファイル名 = モジュール名が衝突するため、出力 slug を別名にする。
SLUG_OVERRIDE = {
    "ThisWorkbook_管理": "thisworkbook-admin",
    "ThisWorkbook_検索": "thisworkbook-search",
    "ThisWorkbook_登録修正": "thisworkbook-register",
}
# 画面表示用のモジュール名（ThisWorkbook 系はブック名を併記）。
DISPLAY_OVERRIDE = {
    "ThisWorkbook_管理": "ThisWorkbook（管理.xlsm）",
    "ThisWorkbook_検索": "ThisWorkbook（検索.xlsm）",
    "ThisWorkbook_登録修正": "ThisWorkbook（登録修正.xlsm）",
}


def is_test_module(name: str) -> bool:
    """テスト用モジュールかどうか。名前に Test を含むものはすべてテスト用。"""
    return "test" in name.lower()


def read_source(path: str) -> str:
    """VBA ソースを読み込む。CP932 優先、だめなら UTF-8、最後は CP932 + 置換。"""
    raw = open(path, "rb").read()
    for enc in ("cp932", "utf-8"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("cp932", errors="replace")


def module_type(name: str, ext: str) -> str:
    if name.startswith("ThisWorkbook"):
        return "ドキュメントモジュール (ThisWorkbook)"
    if ext == ".bas":
        return "標準モジュール (.bas)"
    if name == "IScreenRenderer":
        return "インターフェースクラス (.cls)"
    return "クラスモジュール (.cls)"


def is_predeclared(src: str) -> bool:
    for line in src.splitlines()[:6]:
        if "VB_PredeclaredId" in line and "True" in line:
            return True
    return False


def placement_note(name: str, ext: str, src: str) -> str:
    if name.startswith("ThisWorkbook"):
        return (
            "ブックに最初から存在する `ThisWorkbook` モジュール"
            "（VBE のプロジェクトツリーにある `ThisWorkbook`）に貼り付けます。"
            "新規モジュールとしては取り込めません。先頭の `VERSION 1.0 CLASS` から "
            "`Attribute` 行までのファイル先頭ブロックは貼り付けず、その下の本体だけを"
            "コードペインに貼り付けて既存の内容を置き換えてください。"
            "詳しい手順は[導入手順](../../setup.md)を参照してください。"
        )
    if ext == ".bas":
        return (
            f"標準モジュール（.bas）です。下記コードをコピーし、`{name}.bas` という"
            "ファイル名で保存して、VBE の「ファイル → ファイルのインポート」で"
            "取り込みます。詳しい手順は[導入手順](../../setup.md)を参照してください。"
        )
    # クラスモジュール
    note = (
        f"クラスモジュール（.cls）です。下記コードをコピーし、`{name}.cls` という"
        "ファイル名で保存して、VBE の「ファイル → ファイルのインポート」で"
        "取り込みます。先頭の `VERSION 1.0 CLASS` から始まる行はクラスモジュールの"
        "ファイル形式の一部なので、削らずにそのまま保存してください。"
    )
    if name == "IScreenRenderer":
        note += "（このクラスは画面描画クラスが `Implements` で実装するインターフェースです。）"
    note += "詳しい手順は[導入手順](../../setup.md)を参照してください。"
    return note


def render_page(name: str, ext: str, layer_slug: str, src: str) -> str:
    display = DISPLAY_OVERRIDE.get(name, name + ext)
    line_count = src.count("\n") + (0 if src.endswith("\n") else 1)
    role = ROLE_MAP.get(name, "（役割未設定）")
    book = BOOK_MAP.get(name, DEFAULT_BOOK)
    typ = module_type(name, ext)
    note = placement_note(name, ext, src)
    body = src.rstrip("\n")
    return (
        f"---\n"
        f"title: {display}\n"
        f"---\n\n"
        f"# {display}\n\n"
        f"| 項目 | 内容 |\n"
        f"|---|---|\n"
        f"| 層 | {LAYER_JP[layer_slug]} |\n"
        f"| 種別 | {typ} |\n"
        f"| 配置ブック | {book} |\n"
        f"| 役割 | {role} |\n"
        f"| 行数 | {line_count} 行 |\n\n"
        f"## 取り込み先\n\n"
        f"{note}\n\n"
        f"## ソースコード\n\n"
        f"コードブロック右上のボタンで全文をコピーできます。\n\n"
        f'```vbnet linenums="1"\n'
        f"{body}\n"
        f"```\n"
    )


INDEX_INTRO = """---
title: ソースコード
description: 現行 v2.1 の配布対象 VBA モジュールの全ソースをコピー可能な形で公開
---

# ソースコード

knowledgevba v2.1 の配布対象 VBA モジュール **{count} 本** の全ソースを、
コピー可能な形で層別に掲載しています。各ページのコードブロック右上のボタンから
コードを取得し、ご自身の `.xlsm` の VBE（Visual Basic Editor。Excel 上で
`Alt + F11` で開きます）に取り込んでご利用ください。取り込み手順の全体像は
[導入手順](../setup.md)にまとめています。

!!! note "テスト用モジュールは含みません"
    開発時に使う自動テスト用のモジュールは配布対象ではないため、本一覧には
    含めていません。掲載しているのは、ツールを動かすために実際に取り込む
    モジュールだけです。

!!! info "ファイル名のプレフィックスについて"
    VBA の慣習に従い、`mod` で始まるものは**標準モジュール**（手続き・関数の
    集まり = `.bas`）、`cls` で始まるものは**クラスモジュール**（状態を持つ
    オブジェクト = `.cls`）、`I` で始まるものは**インターフェースクラス**、
    `ThisWorkbook` は Excel がブックごとに自動で持つ**特殊なモジュール**です。

## 3 ブック構成とモジュール配置

本ツールは **登録修正.xlsm / 検索.xlsm / 管理.xlsm** の 3 ブックに分かれています。
多くのモジュールは 3 ブックに共通で取り込みますが、画面クラス・各画面の入口・
`ThisWorkbook` はブックごとに異なります。各ページの「配置ブック」欄で、その
モジュールをどのブックに取り込むかを確認できます。
"""


def main() -> int:
    if not os.path.isdir(SRC_ROOT):
        print(f"FATAL: 入力フォルダが見つかりません: {SRC_ROOT}")
        return 1

    # 出力先フォルダを用意する。
    os.makedirs(OUT_ROOT, exist_ok=True)
    for slug, _, _ in LAYERS:
        os.makedirs(os.path.join(OUT_ROOT, slug), exist_ok=True)

    # 入力フォルダの全モジュールを走査。
    present = {}
    for fn in os.listdir(SRC_ROOT):
        if not fn.endswith((".bas", ".cls")):
            continue
        name, ext = os.path.splitext(fn)
        present[name] = (fn, ext)

    written = []     # (layer_slug, name, ext, slug, line_count)
    written_paths = set()
    skipped_test = []
    unknown = []
    for name, (fn, ext) in sorted(present.items()):
        if is_test_module(name):
            skipped_test.append(name)
            continue
        layer_slug = MODULE_LAYER.get(name)
        if layer_slug is None:
            unknown.append(name)
            continue
        src = read_source(os.path.join(SRC_ROOT, fn))
        page = render_page(name, ext, layer_slug, src)
        slug = SLUG_OVERRIDE.get(name, name.lower())
        out_path = os.path.join(OUT_ROOT, layer_slug, slug + ".md")
        with open(out_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(page)
        written_paths.add(os.path.abspath(out_path))
        line_count = src.count("\n") + (0 if src.endswith("\n") else 1)
        written.append((layer_slug, name, ext, slug, line_count))

    # 定義済みだが入力に存在しないモジュールの検出。
    missing = [m for m in MODULE_LAYER if m not in present]

    # index.md を生成。
    by_layer = {}
    for layer_slug, name, ext, slug, line_count in written:
        by_layer.setdefault(layer_slug, []).append((name, ext, slug, line_count))
    index_parts = [INDEX_INTRO.format(count=len(written)), "\n## 層別一覧\n"]
    for slug, jp, desc in LAYERS:
        mods = sorted(by_layer.get(slug, []), key=lambda x: x[0].lower())
        index_parts.append(f"\n### {jp}（{len(mods)}）\n")
        index_parts.append(f"\n{desc}\n")
        index_parts.append("\n| モジュール | 配置ブック | 役割 |\n|---|---|---|\n")
        for name, ext, mslug, _ in mods:
            display = DISPLAY_OVERRIDE.get(name, name + ext)
            book = BOOK_MAP.get(name, DEFAULT_BOOK)
            role = ROLE_MAP.get(name, "")
            index_parts.append(
                f"| [{display}]({slug}/{mslug}.md) | {book} | {role} |\n")
    index_path = os.path.join(OUT_ROOT, "index.md")
    with open(index_path, "w", encoding="utf-8", newline="\n") as f:
        f.write("".join(index_parts))
    written_paths.add(os.path.abspath(index_path))

    # 旧構成のソースページ（現行に存在しないモジュール）を空にする。
    # 削除権限が無い環境では物理削除できないため、内容を空にして残骸を消す。
    emptied = []
    for root, _dirs, files in os.walk(OUT_ROOT):
        for fn in files:
            if not fn.endswith(".md"):
                continue
            p = os.path.abspath(os.path.join(root, fn))
            if p in written_paths:
                continue
            if os.path.getsize(p) == 0:
                continue
            with open(p, "w", encoding="utf-8", newline="\n") as f:
                f.write("")
            emptied.append(os.path.relpath(p, OUT_ROOT))

    # nav 断片を出力（mkdocs.yml に貼り付ける用）。
    nav_lines = ["  - ソースコード:", "      - 一覧: source/index.md"]
    for slug, jp, _ in LAYERS:
        mods = sorted(by_layer.get(slug, []), key=lambda x: x[0].lower())
        nav_lines.append(f"     