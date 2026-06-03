"""
ui_seed stanza 10 file -> mkdocs site md page (CP932 + CRLF) generator.
Source stanza is CP932; output md is CP932 + CRLF too (Agent S/T/U pattern).
This script itself is UTF-8.
"""
import codecs
import os

import sys
if sys.platform.startswith("linux"):
    SRC_BASE = "/sessions/confident-gracious-brown/mnt/kvba/publish/dist_v2/ui_seed"
    DST_BASE = "/sessions/confident-gracious-brown/mnt/kvba/push/docs/stanza"
else:
    SRC_BASE = r"C:\kvba\publish\dist_v2\ui_seed"
    DST_BASE = r"C:\kvba\push\docs\stanza"

# (role_ascii, filename, screen_name_jp, role_jp_dir)
TARGETS = [
    ("kanri",   "M-02.txt", "フォーマット一覧", "管理"),
    ("kanri",   "M-03.txt", "フォーマット設計", "管理"),
    ("kanri",   "M-04.txt", "プレビュー", "管理"),
    ("kanri",   "M-10.txt", "格納先設定", "管理"),
    ("kanri",   "M-11.txt", "設定", "管理"),
    ("kanri",   "M-12.txt", "フォーマット変更チェック", "管理"),
    ("kanri",   "M-14.txt", "操作ログ", "管理"),
    ("touroku", "M-05.txt", "ナレッジ登録", "登録修正"),
    ("touroku", "M-06.txt", "ナレッジ修正", "登録修正"),
    ("kensaku", "M-08.txt", "ナレッジ検索", "検索"),
]


def build_md(role_jp, fname, screen_name, stanza_body):
    mxx = fname[:-4]  # "M-02"
    md = []
    md.append("---")
    md.append("title: " + mxx + " " + screen_name)
    md.append("description: " + screen_name + "画面の設定ファイル (" + fname + ")")
    md.append("---")
    md.append("")
    md.append("# " + mxx + " " + screen_name)
    md.append("")
    md.append("`" + screen_name + "` 画面の設定ファイルです。")
    md.append("")
    md.append("## 配置先")
    md.append("")
    md.append("以下のパスに **そのままのファイル名で** 配置してください。")
    md.append("")
    md.append("```text")
    md.append("C:\\KnowledgeMgr\\ui\\" + role_jp + "\\" + fname)
    md.append("```")
    md.append("")
    md.append("## 保存手順")
    md.append("")
    md.append("1. 下の **内容** を全選択してコピー (右上のコピーボタンが便利)")
    md.append("2. **メモ帳** を開いて貼り付け")
    md.append("3. 「ファイル」 → 「名前を付けて保存」")
    md.append("4. ファイル名を `" + fname + "` にする")
    md.append("5. **文字コードを「ANSI」(Shift_JIS) に変更** ← 重要")
    md.append("6. 改行コードは「CRLF」(Windows 標準) のままにする")
    md.append("7. 上記 **配置先** のパスに保存")
    md.append("")
    md.append("> 文字コードを UTF-8 のままにすると、`管理.xlsm` 等を開いたときに")
    md.append("> 画面タイトルや項目名が **文字化け** します。必ず `ANSI` で保存してください。")
    md.append("")
    md.append("## 内容")
    md.append("")
    md.append("```ini")
    body = stanza_body.replace("\r\n", "\n").rstrip("\n")
    for ln in body.split("\n"):
        md.append(ln)
    md.append("```")
    md.append("")
    return "\r\n".join(md) + "\r\n"


def main():
    written = []
    for role, fname, screen, role_jp in TARGETS:
        src = os.path.join(SRC_BASE, role_jp, fname)
        with codecs.open(src, encoding="cp932") as f:
            body = f.read()

        md_text = build_md(role_jp, fname, screen, body)

        dst_dir = os.path.join(DST_BASE, role_jp)
        if not os.path.isdir(dst_dir):
            os.makedirs(dst_dir)
        dst = os.path.join(dst_dir, fname + ".md")
        with codecs.open(dst, "w", encoding="cp932") as f:
            f.write(md_text)

        with codecs.open(dst, encoding="cp932") as f:
            rb = f.read()
        assert rb == md_text, "round-trip mismatch: " + dst
        assert "```\r\n" in rb, "missing closing fence: " + dst

        written.append((dst, len(md_text)))
        print("OK", dst, len(md_text))

    print("=== summary: %d files ===" % len(written))


if __name__ == "__main__":
    main()
