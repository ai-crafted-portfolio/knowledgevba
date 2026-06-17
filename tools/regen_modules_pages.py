#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docs/modules/ の VBA モジュールページを CP932 source から再生成する。

入力 : C:\\kvba\\publish\\dist_v2\\installer\\vba_modules\\<role>\\*.bas|*.cls
出力 : C:\\kvba\\push\\docs\\modules\\<role>\\<filename_lower>.md

仕様:
- canonical source は CP932 (Shift_JIS) + CRLF で保存されている。
  必ず encoding="cp932" で読み込み、UTF-8 で md に書き出す。
- 1 件だけ UTF-8 保存の .cls (clsSearchScreen.cls) が混ざるので、
  CP932 strict で失敗した場合のみ UTF-8 として読み込む fallback を持つ。
- md 冒頭に「メモ帳で ANSI (Shift_JIS) 保存」明示する。
- code block 内は元の VBA source を fence "```vb" で囲んで全文掲載する。
"""
from __future__ import annotations

import os
import sys

ROLES = ["admin", "register", "search", "common"]
ROLE_JP = {
    "admin": "管理.xlsm",
    "register": "登録修正.xlsm",
    "search": "検索.xlsm",
    "common": "共通モジュール",
}

SRC_ROOT_WIN = r"C:\kvba\publish\dist_v2\installer\vba_modules"
OUT_ROOT_WIN = r"C:\kvba\push\docs\modules"
import glob as _glob
_kvba = (_glob.glob("/sessions/*/mnt/kvba") or ["/sessions/none/mnt/kvba"])[0]
SRC_ROOT_LIN = _kvba + "/publish/dist_v2/installer/vba_modules"
OUT_ROOT_LIN = _kvba + "/push/docs/modules"

if os.path.exists(SRC_ROOT_WIN):
    SRC_ROOT = SRC_ROOT_WIN
    OUT_ROOT = OUT_ROOT_WIN
else:
    SRC_ROOT = SRC_ROOT_LIN
    OUT_ROOT = OUT_ROOT_LIN


def list_canonical(role):
    src_dir = os.path.join(SRC_ROOT, role)
    out = []
    for name in sorted(os.listdir(src_dir)):
        if ".bak" in name or ".broken" in name or name.startswith("."):
            continue
        if not (name.endswith(".bas") or name.endswith(".cls")):
            continue
        out.append(name)
    return out


def render_md(role, src_name, src_text, mtime_str):
    ext = ".bas" if src_name.endswith(".bas") else ".cls"
    kind_label = "標準モジュール" if ext == ".bas" else "クラスモジュール"
    if role == "common":
        save_dir = r"C:\KnowledgeMgr\installer\vba_modules\common\\"
        target_book = "共通モジュール（3 ブック共通）"
    else:
        book = ROLE_JP[role]
        save_dir = "C:\\KnowledgeMgr\\installer\\vba_modules\\" + role + "\\"
        target_book = "`" + book + "` 用の VBA モジュール"

    if src_name == "ThisWorkbook.cls":
        # ThisWorkbook はブックに最初から存在する document module。
        # ファイル Import 不可。VBE で既存 ThisWorkbook に本体をコピペする。
        # クラスモジュールヘッダ（VERSION 1.0 CLASS〜Attribute）は不要で、
        # 本体は Option Explicit から始まる。
        notice = (
            "## 取り込み方法（ThisWorkbook は特別）\n\n"
            "**【重要】`ThisWorkbook` は Excel のブック（.xlsm）に最初から存在する "
            "document module です。ファイルとして Import できません。**\n\n"
            "1. VBE（`Alt + F11`）を開き、プロジェクトツリーの `ThisWorkbook` を "
            "**ダブルクリック**してコードペインを開きます。\n"
            "2. コードペインの**既存コードをすべて削除**します。\n"
            "3. 下のコード全文をコピーして**貼り付け**ます。\n\n"
            "> VB クラスモジュールヘッダ（`VERSION 1.0 CLASS` / `BEGIN` / `END` / "
            "`Attribute` 行）は不要です。本体は `Option Explicit` から始まります。\n"
            "> 文字コード・改行はコードペインに直接貼り付けるため、ファイル保存時の "
            "ANSI/CRLF 指定は不要です。\n"
        )
    else:
        notice = (
            "## 保存方法\n\n"
            "下のコードをメモ帳に貼り付け、**[名前を付けて保存]** で次のように保存してください。\n\n"
            "- 場所: `" + save_dir + "`\n"
            "- ファイル名: `" + src_name + "`\n"
            "- ファイルの種類: **すべてのファイル**\n"
            "- 文字コード: **ANSI**（Shift-JIS）\n\n"
            "> メモ帳の文字コードを **ANSI** にしないと、VBA の日本語が文字化けして動かなくなります。\n"
            "> UTF-8 で保存すると VBA Import 時に日本語が文字化けして動かなくなります。\n"
            "> 改行コードは CRLF（Windows 標準）のままで OK です。\n"
        )
    body = (
        "---\n"
        "title: " + src_name + "\n"
        "description: " + src_name + " のソースコード（コピペ用）\n"
        "---\n\n"
        "# " + src_name + "\n\n"
        "**配置先**: " + target_book + "\n"
        "**種類**: " + kind_label + "\n"
        "**更新日**: " + mtime_str + "\n\n"
        "---\n\n"
        + notice +
        "\n---\n\n"
        "## ソースコード\n\n"
        "```vb\n"
        + src_text + "\n"
        "```\n"
    )
    return body


def gen_index(collected):
    intro = (
        "---\n"
        "title: モジュール一覧\n"
        "description: VBA モジュール（.bas / .cls）のコピペ用ソースコード一覧\n"
        "---\n\n"
        "# モジュール一覧\n\n"
        "knowledgevba を構成する VBA モジュールのソースコード一覧です。各ページに該当ファイルの全文がコピペできる形で掲載されています。\n\n"
        "[インストール手順](../install.md) の **STEP 6** で、これらを `installer\\vba_modules\\<役割>\\` 配下に保存してください。\n\n"
        "**更新日** は canonical ソースの最終更新日時です。手元のファイルより新しければ差し替えてください。\n\n"
        "!!! info \"保存時のお願い\"\n"
        "    - メモ帳で **[名前を付けて保存]** → 文字コードは **ANSI**（Shift-JIS）を選んでください\n"
        "    - UTF-8 で保存すると日本語が文字化けして動かなくなります\n"
        "    - ファイル名は **大文字小文字を区別しません**（VBE 側で正しい名前が付きます）\n\n"
        "---\n\n"
    )
    body = intro
    heads = {"admin": "## 管理.xlsm 用 (`installer\\vba_modules\\admin\\`)",
             "register": "## 登録修正.xlsm 用 (`installer\\vba_modules\\register\\`) ※v2.3 では廃止ブック（参考掲載）",
             "search": "## 検索.xlsm 用 (`installer\\vba_modules\\search\\`)",
             "common": "## 共通モジュール (`installer\\vba_modules\\common\\`)"}
    for role in ROLES:
        rows = collected.get(role, [])
        body += heads[role] + "\n\n**" + str(len(rows)) + " ファイル**\n\n"
        body += "| ファイル名 | 種類 | 更新日 |\n|---|---|---|\n"
        for fname, mt in rows:
            kind = "標準" if fname.endswith(".bas") else "クラス"
            body += "| [`" + fname + "`](" + role + "/" + fname.lower() + ".md) | " + kind + " | " + mt + " |\n"
        body += "\n"
    out_path = os.path.join(OUT_ROOT, "index.md")
    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(body)
    print("[done] index.md regenerated with dates")


def main():
    total = 0
    failures = []
    import datetime
    collected = {}
    for role in ROLES:
        src_dir = os.path.join(SRC_ROOT, role)
        out_dir = os.path.join(OUT_ROOT, role)
        os.makedirs(out_dir, exist_ok=True)
        files = list_canonical(role)
        print("[" + role + "] " + str(len(files)) + " files")
        collected[role] = []
        for fname in files:
            src_path = os.path.join(src_dir, fname)
            mtime_str = datetime.datetime.fromtimestamp(os.path.getmtime(src_path), datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M")
            collected[role].append((fname, mtime_str))
            with open(src_path, "rb") as f:
                raw = f.read()
            try:
                txt = raw.decode("cp932")
                used = "cp932"
            except UnicodeDecodeError:
                try:
                    txt = raw.decode("utf-8")
                    used = "utf-8"
                    print("  [note] " + fname + ": UTF-8 として読込（CP932 不可）")
                except UnicodeDecodeError as e:
                    failures.append((src_path, "cp932/utf8 decode fail: " + str(e)))
                    continue
            txt_lf = txt.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
            md_text = render_md(role, fname, txt_lf, mtime_str)
            out_name = fname.lower() + ".md"
            out_path = os.path.join(out_dir, out_name)
            with open(out_path, "w", encoding="utf-8", newline="\n") as f:
                f.write(md_text)
            total += 1
    gen_index(collected)
    print("[done] regenerated " + str(total) + " pages")
    if failures:
        print("[fail] " + str(len(failures)) + " files")
        for p, msg in failures:
            print("  " + p + ": " + msg)
        sys.exit(1)


if __name__ == "__main__":
    main()
