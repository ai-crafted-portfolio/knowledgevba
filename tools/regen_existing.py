#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""drift した既存ページのみ canonical(uat/release_v2) から再生成する。
新規ページ作成・既存ページ削除は一切しない。フォーマットは regen_modules_pages.render_md を流用。"""
import os, sys, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import regen_modules_pages as gen

UAT = r"C:\kvba\uat\release_v2\installer\vba_modules"
DOCS = r"C:\kvba\push\docs\modules"
SEARCH_ORDER = {
    "admin":  ["admin", "quarantine"],
    "common": ["common", "quarantine"],
    "search": ["search", "register", "quarantine"],
}

# 再生成対象 (role, published-md-basename)
TARGETS = [
    ("admin", "modentrysettings.bas"),
    ("admin", "thisworkbook.cls"),
    ("common", "clssetuporchestrator.cls"),
    ("common", "clsuserformrenderer.cls"),
    ("common", "modbtnmessages.bas"),
    ("common", "modconfigholder.bas"),
    ("common", "modconfigloader.bas"),
    ("common", "modentryuserform.bas"),
    ("common", "modformatloader.bas"),
    ("common", "modknowledgefileio.bas"),
    ("common", "modrefresh.bas"),
    ("common", "modsetup.bas"),
    ("common", "moduserformcallback.bas"),
    ("search", "modentrysearch.bas"),
    ("search", "thisworkbook.cls"),
]

def find_canonical(role, fname):
    for sub in SEARCH_ORDER[role]:
        d = os.path.join(UAT, sub)
        if not os.path.isdir(d):
            continue
        for name in os.listdir(d):
            if name.lower() == fname.lower() and (name.endswith(".bas") or name.endswith(".cls")):
                return os.path.join(d, name), name
    return None, None

n = 0
for role, base in TARGETS:
    canon_path, canon_name = find_canonical(role, base)
    if not canon_path:
        print("[fail] no canonical for %s/%s" % (role, base)); sys.exit(1)
    with open(canon_path, "rb") as f:
        raw = f.read()
    try:
        txt = raw.decode("cp932")
    except UnicodeDecodeError:
        txt = raw.decode("utf-8")
    txt_lf = txt.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
    mtime_str = datetime.datetime.fromtimestamp(
        os.path.getmtime(canon_path),
        datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M JST")
    md_text = gen.render_md(role, canon_name, txt_lf, mtime_str)
    out_path = os.path.join(DOCS, role, base + ".md")
    with open(out_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(md_text)
    print("[done] %s/%s.md  <- %s/%s  (%s)" % (role, base, role, canon_name, mtime_str))
    n += 1
print("[ok] regenerated %d pages" % n)
