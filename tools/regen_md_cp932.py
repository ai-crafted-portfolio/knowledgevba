#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docs/modules/ の VBA モジュールページを CP932 source から CP932 md として再生成する。

入力 : C:\\kvba\\publish\\dist_v2\\installer\\vba_modules\\<role>\\*.bas|*.cls
出力 : C:\\kvba\\push\\docs\\modules\\<role>\\<filename_lower>.md  (CP932 encoded)

仕様:
- canonical source は CP932 (Shift_JIS) + CRLF で保存されている。
  必ず encoding="cp932" で読み込み、**md も CP932 で書き出す**。
- 1 件だけ UTF-8 保存の .cls (clsSearchScreen.cls) が混ざるので、
  CP932 strict で失敗した場合のみ UTF-8 として読み込む fallback を持ち、
  md は CP932 で出力する（出力時の不可文字は ? に replace）。
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
SRC_ROOT_LIN = "/sessions/confident-gracious-brown/mnt/kvba/publish/dist_v2/installer/vba_modules"
OUT_ROOT_LIN = "/sessions/confident-gracious-brown/mnt/kvba/push/docs/modules"

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


def render_md(role, src_name, src_text):
    ext = ".bas" if src_name.endswith(".bas") else ".cls"
    kind_label = "標準モジュール" if ext == ".bas" else "クラスモジュール"
    if role == "common":
        save_dir = r"C:\KnowledgeMgr\installer\vba_modules\common\\"
        target_book = "共通モジュール（3 ブック共通）"
    else:
        book = ROLE_JP[role]
        save_dir = "C:\\KnowledgeMgr\\installer\\vba_modules\\" + role + "\\"
        target_book = "`" + book + "` 用の VBA モジュール"

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
        "**配置先**: " + target_book + "  \n"
        "**種類**: " + kind_label + "\n\n"
        "---\n\n"
        + notice +
        "\n---\n\n"
        "## ソースコード\n\n"
        "```vb\n"
        + src_text + "\n"
        "```\n"
    )
    return body


def main():
    total = 0
    failures = []
    cp932_count = 0
    utf8_count = 0
    for role in ROLES:
        src_dir = os.path.join(SRC_ROOT, role)
        out_dir = os.path.join(OUT_ROOT, role)
        os.makedirs(out_dir, exist_ok=True)
        files = list_canonical(role)
        print("[" + role + "] " + str(len(files)) + " files")
        for fname in files:
            src_path = os.path.join(src_dir, fname)
            with open(src_path, "rb") as f:
                raw = f.read()
            try:
                txt = raw.decode("cp932")
                used = "cp932"
                cp932_count += 1
            except UnicodeDecodeError:
                try:
                    txt = raw.decode("utf-8")
                    used = "utf-8"
                    utf8_count += 1
                    print("  [note] " + fname + ": UTF-8 として読込（CP932 不可）")
                except UnicodeDecodeError as e:
                    failures.append((src_path, "cp932/utf8 decode fail: " + str(e)))
                    continue
            # Normalize line endings to LF inside the md code block (markdown render OK both ways)
            txt_lf = txt.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
            md_body = render_md(role, fname, txt_lf)
            out_name = fname.lower() + ".md"
            out_path = os.path.join(out_dir, out_name)
            # Write CP932 with 'replace' fallback for any chars not in CP932 (e.g. UTF-8 source file edge cases)
            # use CRLF line endings (Windows-style, since this is being pushed to a Windows repo)
            md_crlf = md_body.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
            try:
                md_bytes = md_crlf.encode("cp932")
                replaced = 0
            except UnicodeEncodeError:
                md_bytes = md_crlf.encode("cp932", errors="replace")
                replaced = sum(1 for c in md_crlf if ord(c) > 0x7F and c not in md_crlf.encode("cp932", errors="replace").decode("cp932", errors="replace"))
                print("  [warn] " + fname + ": CP932 encode replaced some chars")
            with open(out_path, "wb") as f:
                f.write(md_bytes)
            total += 1
    print("---")
    print("total files written: " + str(total))
    print("  cp932 source: " + str(cp932_count))
    print("  utf-8 source: " + str(utf8_count))
    print("  failures:     " + str(len(failures)))
    for p, msg in failures:
        print("    " + p + " : " + msg)
    return 0 if not failures else 1


if __name__ == "__main__":
    sys.exit(main())
