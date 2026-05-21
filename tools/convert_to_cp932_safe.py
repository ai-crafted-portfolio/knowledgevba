#!/usr/bin/env python3
# convert_to_cp932_safe.py
#
# Convert UTF-8 source files (.bas / .cls / .md / .vb) so that they
# round-trip losslessly through Windows CP932 (Shift_JIS) - the codepage
# Excel VBE saves as on Japanese Windows.
#
# Background:
#     User input on Windows-JP is CP932. When the developer writes
#     characters such as U+25B6, U+00B2, U+2264 in a .bas/.cls and the
#     file is later imported via VBE, the save round-trip silently
#     replaces them with "?", silently corrupting the source.
#
# Strategy:
#     Replace each non-CP932 character with a CP932-safe substitute that
#     preserves the meaning of comments. Replacements whose source is
#     already CP932-safe are filtered out at module-load time so that
#     valid JIS X 0208 chars (x, /, +/-, ~, etc.) are NEVER touched.
#
#     For runtime literals where a different glyph at runtime would
#     matter (button captions etc.) the developer must convert the
#     literal to ChrW(&Hxxxx) by hand - this script refuses to silently
#     change such cases (RUNTIME_LITERAL_GUARD).
#
# Usage:
#     python3 convert_to_cp932_safe.py --check FILE [FILE ...]
#     python3 convert_to_cp932_safe.py --in-place FILE_OR_DIR [...]

from __future__ import annotations
import argparse
import sys
from pathlib import Path

_RAW_REPLACEMENTS = [
    ("▶", ">"),    # BLACK RIGHT-POINTING TRIANGLE   - not CP932
    ("◀", "<"),    # BLACK LEFT-POINTING TRIANGLE    - not CP932
    ("²", "^2"),   # SUPERSCRIPT TWO                 - not CP932
    ("³", "^3"),   # SUPERSCRIPT THREE               - not CP932
    ("≤", "<="),   # LESS-THAN OR EQUAL              - not CP932
    ("≥", ">="),   # GREATER-THAN OR EQUAL           - not CP932
    ("–", "-"),    # EN DASH                         - not CP932
    ("—", "-"),    # EM DASH                         - not CP932
    # The following are CP932-safe but listed defensively. They are
    # filtered out at load time so they cannot accidentally rewrite
    # valid Japanese punctuation.
    ("≠", "!="),   # NOT EQUAL              - CP932-safe
    ("±", "+/-"),  # PLUS-MINUS             - CP932-safe
    ("×", "*"),    # MULTIPLICATION         - CP932-safe
    ("÷", "/"),    # DIVISION               - CP932-safe
    ("〜", "~"),    # WAVE DASH              - CP932-safe
    ("−", "-"),    # MINUS SIGN             - CP932-safe
]

RUNTIME_LITERAL_GUARD = {"▶", "◀"}


def _filter_to_unsafe(table):
    out = []
    for src, dst in table:
        try:
            src.encode("cp932")
        except UnicodeEncodeError:
            out.append((src, dst))
    return out


REPLACEMENTS = _filter_to_unsafe(_RAW_REPLACEMENTS)


def find_runtime_literals(text):
    hits = []
    for i, line in enumerate(text.splitlines(), start=1):
        if line.lstrip().startswith("'"):
            continue
        in_str = False
        for ch in line:
            if ch == '"':
                in_str = not in_str
            elif in_str and ch in RUNTIME_LITERAL_GUARD:
                hits.append((i, line))
                break
    return hits


def check_cp932(text):
    bad = []
    for i, ch in enumerate(text):
        try:
            ch.encode("cp932")
        except UnicodeEncodeError:
            bad.append((i, ch))
    return bad


def convert(text):
    for src, dst in REPLACEMENTS:
        text = text.replace(src, dst)
    return text


def process_file(path, write, allow_runtime_lit):
    raw = path.read_text(encoding="utf-8")
    runtime_hits = find_runtime_literals(raw)
    if runtime_hits and not allow_runtime_lit:
        return {
            "path": str(path),
            "status": "BLOCKED_RUNTIME_LITERAL",
            "runtime_hits": runtime_hits,
            "before_bad": check_cp932(raw),
        }
    converted = convert(raw)
    after_bad = check_cp932(converted)
    if after_bad:
        return {
            "path": str(path),
            "status": "STILL_BAD_AFTER_CONVERT",
            "after_bad": after_bad,
            "before_bad": check_cp932(raw),
        }
    if write and converted != raw:
        path.write_text(converted, encoding="utf-8")
    if write and converted != raw:
        status = "OK_WRITTEN"
    elif converted == raw:
        status = "OK_NOOP"
    else:
        status = "OK_DRYRUN"
    return {
        "path": str(path),
        "status": status,
        "delta_bytes": len(converted.encode("utf-8")) - len(raw.encode("utf-8")),
        "replaced_count": sum(raw.count(src) for src, _ in REPLACEMENTS),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="+")
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--in-place", action="store_true")
    ap.add_argument("--allow-runtime-literal", action="store_true")
    ap.add_argument("--ext", default=".bas,.cls,.md,.vb,.frm")
    args = ap.parse_args()

    if not (args.check or args.in_place):
        print("ERROR: pass --check or --in-place", file=sys.stderr)
        return 2

    exts = set()
    for e in args.ext.split(","):
        e = e.strip()
        if not e.startswith("."):
            e = "." + e
        exts.add(e)

    targets = []
    for p in args.paths:
        pp = Path(p)
        if pp.is_dir():
            for q in sorted(pp.rglob("*")):
                if q.suffix in exts and q.is_file():
                    targets.append(q)
        elif pp.is_file():
            targets.append(pp)
        else:
            print("WARN: not found: " + p, file=sys.stderr)

    failed = 0
    for t in targets:
        res = process_file(t, args.in_place, args.allow_runtime_literal)
        if res["status"].startswith("OK"):
            if res.get("replaced_count", 0):
                print("  {0:<12} {1}  (replaced={2}, dB={3:+d})".format(
                    res["status"], res["path"],
                    res["replaced_count"], res["delta_bytes"]))
            elif not args.check:
                print("  {0:<12} {1}".format(res["status"], res["path"]))
        else:
            failed += 1
            print("  {0:<24} {1}".format(res["status"], res["path"]))
            if "runtime_hits" in res:
                for ln, line in res["runtime_hits"][:5]:
                    print("    runtime literal L{0}: {1}".format(ln, line.strip()[:120]))
            if "after_bad" in res:
                for off, ch in res["after_bad"][:5]:
                    print("    still bad off {0}: U+{1:04X} {2!r}".format(
                        off, ord(ch), ch))
    if failed:
        print("", file=sys.stderr)
        print("{0} file(s) need manual attention.".format(failed),
              file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
