#!/usr/bin/env python3
# verify_vba_cp932_compat.py
#
# Static verification: every .bas / .cls / .frm / .md file under the
# given roots must round-trip through CP932 (Shift_JIS) without loss.
#
# Rule 7 (build pipeline): VBA source MUST contain only characters that
# survive a CP932 round-trip, because Excel VBE on Japanese Windows saves
# in CP932 and silently maps anything else to "?".
#
# Exit codes:
#   0  - all files pass
#   1  - one or more files fail
#   2  - bad invocation
#
# Usage:
#   python3 verify_vba_cp932_compat.py PATH [PATH ...]
#   python3 verify_vba_cp932_compat.py --quiet PATH ...   # only print fails
#   python3 verify_vba_cp932_compat.py --ext .bas,.cls PATH ...

from __future__ import annotations
import argparse
import sys
from pathlib import Path

DEFAULT_EXTS = ".bas,.cls,.frm,.md,.vb"


def collect(paths, exts, exclude_segments):
    """Collect files. exclude_segments is an iterable of path-segment names;
    any file whose path contains one of these as a directory component is
    skipped (e.g. '_archive', '__pycache__')."""
    out = []
    for p in paths:
        pp = Path(p)
        if pp.is_dir():
            for q in sorted(pp.rglob("*")):
                if not q.is_file() or q.suffix not in exts:
                    continue
                parts = set(q.parts)
                if parts & set(exclude_segments):
                    continue
                out.append(q)
        elif pp.is_file():
            out.append(pp)
        else:
            print("WARN: not found: " + str(p), file=sys.stderr)
    return out


def check(path):
    try:
        with open(path, "rb") as f:
            raw = f.read()
    except OSError as e:
        return ("READ_ERROR", [(0, str(e))])
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as e:
        return ("UTF8_ERROR", [(e.start, "invalid UTF-8: " + str(e))])
    bad = []
    for i, ch in enumerate(text):
        try:
            ch.encode("cp932")
        except UnicodeEncodeError:
            line_no = text[:i].count("\n") + 1
            bad.append((line_no, "U+{0:04X} {1!r}".format(ord(ch), ch)))
    if bad:
        return ("CP932_FAIL", bad)
    return ("PASS", [])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="+")
    ap.add_argument("--quiet", action="store_true",
                    help="Only print files that FAIL")
    ap.add_argument("--ext", default=DEFAULT_EXTS,
                    help="Comma-separated extensions (default: " + DEFAULT_EXTS + ")")
    ap.add_argument("--exclude", default="_archive,__pycache__,.git,site,node_modules",
                    help="Comma-separated directory NAMES to skip (any path "
                         "containing one of these segments is excluded). "
                         "Default: _archive,__pycache__,.git,site,node_modules")
    args = ap.parse_args()

    exts = set()
    for e in args.ext.split(","):
        e = e.strip()
        if not e.startswith("."):
            e = "." + e
        exts.add(e)

    excludes = set(s.strip() for s in args.exclude.split(",") if s.strip())

    targets = collect(args.paths, exts, excludes)
    if not targets:
        print("ERROR: no files matched", file=sys.stderr)
    