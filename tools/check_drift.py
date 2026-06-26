#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""既存の docs/modules/*/*.md の埋め込みコードを canonical(uat/release_v2) と照合し drift を報告する。"""
import os, re, sys

UAT = r"C:\kvba\uat\release_v2\installer\vba_modules"
DOCS = r"C:\kvba\push\docs\modules"
ROLES = ["admin", "common", "search"]
# canonical 探索順: 同 role -> register -> quarantine
SEARCH_ORDER = {
    "admin":  ["admin", "quarantine"],
    "common": ["common", "quarantine"],
    "search": ["search", "register", "quarantine"],
}

def read_src(path):
    with open(path, "rb") as f:
        raw = f.read()
    try:
        txt = raw.decode("cp932")
    except UnicodeDecodeError:
        txt = raw.decode("utf-8")
    return txt.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")

def find_canonical(role, fname):
    for sub in SEARCH_ORDER[role]:
        d = os.path.join(UAT, sub)
        if not os.path.isdir(d):
            continue
        for name in os.listdir(d):
            if name.lower() == fname.lower() and (name.endswith(".bas") or name.endswith(".cls")):
                return os.path.join(d, name), name
    return None, None

def extract_code(md_path):
    with open(md_path, "r", encoding="utf-8") as f:
        txt = f.read()
    m = re.search(r"```vb\n(.*?)\n```", txt, re.S)
    if not m:
        return None
    return m.group(1).replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")

drift, ok, nocanon = [], 0, []
for role in ROLES:
    d = os.path.join(DOCS, role)
    for md in sorted(os.listdir(d)):
        if not md.endswith(".md") or md in ("index.md", "test.md"):
            continue
        fname = md[:-3]  # strip .md  -> e.g. clsuserformrenderer.cls
        md_path = os.path.join(d, md)
        canon_path, canon_name = find_canonical(role, fname)
        if not canon_path:
            nocanon.append(f"{role}/{md}")
            continue
        emb = extract_code(md_path)
        src = read_src(canon_path)
        if emb is None:
            drift.append((f"{role}/{md}", "NO CODE BLOCK"))
        elif emb != src:
            drift.append((f"{role}/{md}", f"DRIFT vs {role}->{canon_name}"))
        else:
            ok += 1

print(f"OK (match): {ok}")
print(f"NO CANONICAL: {len(nocanon)}")
for x in nocanon: print("  ? " + x)
print(f"DRIFT: {len(drift)}")
for p, why in drift: print(f"  ! {p}  [{why}]")
sys.exit(1 if drift else 0)
