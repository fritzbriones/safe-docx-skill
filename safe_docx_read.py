#!/usr/bin/env python3
"""
safe_docx_read.py - Read-only health check for a .docx. Python 3 stdlib only.

Usage:
  python3 safe_docx_read.py input.docx              # health report + outline
  python3 safe_docx_read.py input.docx --full-text  # also every paragraph with its style
  python3 safe_docx_read.py input.docx --protect "A*" --protect "p < .05"

Reports:
  * whether you have REAL file bytes (a ZIP) or only a text extract;
  * every ZIP part with size and full SHA-256 for images;
  * styles.xml size, theme presence, paragraph styles in use (headings!);
  * table column widths from <w:tblGrid> (equal widths = likely rebuilt);
  * header/footer parts and whether they hold PAGE fields;
  * counts: protected literals (A* by default), highlights, tracked changes,
    comments, xml:space, images;
  * DAMAGE WARNINGS: signs the file was rebuilt from text.

Exit status: 0 = read OK, 2 = not a DOCX (no real bytes), 1 = other error.
Never modifies the file.
"""
import argparse, hashlib, re, sys, zipfile
from collections import Counter

_ENTITY_RE = re.compile(r"&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);")
_NAMED = {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'"}


def unescape(s):
    def _d(m):
        e = m.group(1)
        if e.startswith("#x"):
            return chr(int(e[2:], 16))
        if e.startswith("#"):
            return chr(int(e[1:]))
        return _NAMED[e]
    return _ENTITY_RE.sub(_d, s)


WT_RE = re.compile(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>", re.DOTALL)
P_RE = re.compile(r"<w:p(?:\s[^>]*)?(?<!/)>.*?</w:p>", re.DOTALL)
PSTYLE_RE = re.compile(r'<w:pStyle\s+w:val="([^"]+)"')
TBL_RE = re.compile(r"<w:tbl>.*?</w:tbl>", re.DOTALL)
GRID_RE = re.compile(r'<w:gridCol\s+w:w="(\d+)"')


def style_names(styles_xml):
    names = {}
    for m in re.finditer(r'<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>(.*?)</w:style>', styles_xml, re.DOTALL):
        n = re.search(r'<w:name\s+w:val="([^"]+)"', m.group(2))
        names[m.group(1)] = n.group(1) if n else m.group(1)
    return names


def para_list(doc_xml, names):
    out = []
    for p in P_RE.finditer(doc_xml):
        block = p.group(0)
        s = PSTYLE_RE.search(block)
        sid = s.group(1) if s else "Normal"
        text = "".join(unescape(t) for t in WT_RE.findall(block))
        out.append((names.get(sid, sid), text))
    return out


def main():
    ap = argparse.ArgumentParser(description="Read-only DOCX health check")
    ap.add_argument("docx")
    ap.add_argument("--full-text", action="store_true")
    ap.add_argument("--protect", action="append", default=[],
                    help="literal to count (A* is always counted)")
    a = ap.parse_args()
    protect = ["A*"] + [p for p in a.protect if p != "A*"]

    try:
        if not zipfile.is_zipfile(a.docx):
            print(f"NOT A DOCX/ZIP: {a.docx}")
            print("You only have text, not the real file bytes. Do NOT rebuild a DOCX from it.")
            print("Get the real file, or return a change list (Original -> Revised).")
            sys.exit(2)
        z = zipfile.ZipFile(a.docx)
    except OSError as exc:
        print(f"ERROR: {exc}"); sys.exit(1)

    names = z.namelist()
    warnings = []
    print(f"=== {a.docx} ===\nREAL BYTES: yes (valid ZIP, {len(names)} parts)\n")

    print("--- Parts ---")
    for info in z.infolist():
        line = f"  {info.filename:45s} {info.file_size:>10,d} B"
        if info.filename.startswith("word/media/"):
            line += f"  sha256={hashlib.sha256(z.read(info.filename)).hexdigest()}"
        print(line)

    if "word/document.xml" not in names:
        print("\nNo word/document.xml - this is not a Word document."); sys.exit(2)
    doc = z.read("word/document.xml").decode("utf-8", "replace")
    styles = z.read("word/styles.xml").decode("utf-8", "replace") if "word/styles.xml" in names else ""
    snames = style_names(styles)

    print("\n--- Styles & theme ---")
    print(f"  styles.xml: {len(styles.encode()):,d} B, {len(snames)} styles defined")
    has_theme = any(n.startswith("word/theme/") for n in names)
    print(f"  theme: {'present' if has_theme else 'MISSING'}")
    if len(styles.encode()) < 2000:
        warnings.append("styles.xml is under 2 KB - typical of a file rebuilt from scratch")
    if not has_theme:
        warnings.append("no word/theme/ part - Word always writes one; file may be rebuilt")

    paras = para_list(doc, snames)
    used = Counter(s for s, _ in paras)
    print(f"\n--- Paragraph styles in use ({len(paras)} paragraphs) ---")
    for s, c in used.most_common():
        print(f"  {c:5d}  {s}")
    heads = [(s, t) for s, t in paras if re.match(r"(?i)(title|heading)", s)]
    if paras and not heads:
        warnings.append("no Title/Heading styles used - headings may have been flattened to Normal")

    print("\n--- Outline ---")
    for s, t in heads[:60]:
        print(f"  [{s}] {t[:90]}")
    if len(heads) > 60:
        print(f"  ... {len(heads) - 60} more")

    tables = TBL_RE.findall(doc)
    print(f"\n--- Tables ({len(tables)}) ---")
    for i, t in enumerate(tables, 1):
        cols = [int(x) for x in GRID_RE.findall(t)]
        flag = ""
        if len(cols) >= 3 and len(set(cols)) == 1:
            # Not a warning by itself: new Word tables start with equal columns.
            flag = "  <- all equal (normal for a new table; suspicious if the original had varied widths)"
        print(f"  Table {i}: gridCol twips {cols}{flag}")

    print("\n--- Headers / footers / notes ---")
    hf = [n for n in names if re.match(r"word/(header\d*|footer\d*|footnotes|endnotes)\.xml$", n)]
    for n in hf:
        x = z.read(n).decode("utf-8", "replace")
        page = bool(re.search(r"\bPAGE\b", x))
        txt = "".join(unescape(t) for t in WT_RE.findall(x)).strip()
        print(f"  {n}: PAGE field={'yes' if page else 'no'}  text={txt[:60]!r}")
    if not hf:
        print("  none")

    print("\n--- Counts (word/document.xml) ---")
    full = "\n".join(t for _, t in paras)
    for lit in protect:
        print(f"  literal {lit!r}: {full.count(lit)}")
    for label, pat in (("highlights <w:highlight", "<w:highlight"), ("tracked insertions <w:ins ", "<w:ins "),
                       ("tracked deletions <w:del ", "<w:del "), ("comment refs", "<w:commentReference"),
                       ("xml:space", "xml:space"), ("images r:embed", "r:embed")):
        print(f"  {label}: {doc.count(pat)}")
    print(f"  words (approx): {len(full.split()):,d}")

    if a.full_text:
        print("\n--- Full text (style | text) ---")
        for s, t in paras:
            print(f"[{s}] {t}")

    print("\n--- Verdict ---")
    if warnings:
        for w in warnings:
            print(f"  WARN: {w}")
        print("  This file may already be damaged. Tell the user; do not use it as the master.")
    else:
        print("  No damage signs found.")
    sys.exit(0)


if __name__ == "__main__":
    main()
