#!/usr/bin/env python3
"""
safe_docx_read.py - Read a DOCX properly with Python stdlib ONLY.
No python-docx, no Pillow, no Office. Uses zipfile + xml.etree only.

Usage:
  python3 safe_docx_read.py input.docx
  python3 safe_docx_read.py input.docx --full-text

What it does:
  - Opens the DOCX as a ZIP (a DOCX IS a ZIP of XML + images)
  - Lists every part, with SHA256 for images (to prove they are unchanged later)
  - Lists paragraphs with their STYLE (Heading 1, Title, Normal) - not just text
  - Lists tables with their real column widths (gridCol)
  - Counts A* (critical for ABDC), highlights, tracked changes
  - NEVER rebuilds the file. Read-only.
"""
import sys, zipfile, hashlib, re
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def sha(b):
    return hashlib.sha256(b).hexdigest()[:16]

def get_text_of_p(p_elem):
    texts = []
    for t in p_elem.iter(W+"t"):
        texts.append(t.text or "")
    return "".join(texts)

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 safe_docx_read.py input.docx [--full-text]")
        sys.exit(1)
    path = sys.argv[1]
    full = "--full-text" in sys.argv
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        print(f"FILE: {path}")
        print(f"PARTS: {len(names)}")
        for n in names:
            data = z.read(n)
            extra = ""
            if n.startswith("word/media/"):
                extra = f"  bytes={len(data)} sha={sha(data)}  <-- IMAGE, never re-encode"
            print(f"  {n}  {len(data)} bytes{extra}")
        print()
        # document.xml
        xml_bytes = z.read("word/document.xml")
        xml_text = xml_bytes.decode("utf-8", errors="replace")
        print(f"CHECKS in word/document.xml:")
        print(f"  w:highlight count (should be 0 for final): {xml_text.count('w:highlight')}")
        print(f"  A* count (ABDC, must not drop *): {xml_text.count('A*')}")
        # w:ins that are real insertions vs insideH borders: look for <w:ins 
        print(f"  <w:ins tracked insertions: {xml_text.count('<w:ins ')+xml_text.count('<w:ins>')}")
        print()
        root = ET.fromstring(xml_bytes)
        # Paragraphs
        print("PARAGRAPHS (style + text):")
        for i, p in enumerate(root.iter(W+"p")):
            pPr = p.find(W+"pPr")
            style = "Normal"
            if pPr is not None:
                s = pPr.find(W+"pStyle")
                if s is not None:
                    style = s.get(W+"val", "Normal")
                    # Word stores as Heading1, Title, etc. Show as Heading 1 for humans
                    if style == "Heading1":
                        style = "Heading 1"
                    elif style == "Heading2":
                        style = "Heading 2"
                    elif style == "Heading3":
                        style = "Heading 3"
            txt = get_text_of_p(p)
            if not txt.strip():
                continue
            if full or style != "Normal" or len(txt) < 120:
                print(f"  [{i:03d}] [{style}] {txt[:180]}")
            else:
                print(f"  [{i:03d}] [{style}] {txt[:120]}...")
        print()
        # Tables
        print("TABLES (real column widths, not guessed):")
        for ti, tbl in enumerate(root.iter(W+"tbl")):
            grid = tbl.find(W+"tblGrid")
            widths = []
            if grid is not None:
                for gc in grid.findall(W+"gridCol"):
                    widths.append(gc.get(W+"w"))
            rows = list(tbl.findall(W+"tr"))
            cols = len(widths) if widths else 0
            print(f"  Table {ti}: {len(rows)} rows x {cols} cols  gridCol widths (twips) = {widths}")
            # First row text
            if rows:
                cells = rows[0].findall(W+"tc")
                hdr = []
                for c in cells:
                    t = "".join([get_text_of_p(p) for p in c.iter(W+"p")])
                    hdr.append(t[:30])
                print(f"    Header: {' | '.join(hdr)}")
        print()
        # Styles
        try:
            styles_xml = z.read("word/styles.xml").decode("utf-8", errors="replace")
            print(f"STYLES: word/styles.xml length = {len(styles_xml)} bytes")
            print(f"  Has Heading1? {'Heading1' in styles_xml or 'Heading 1' in styles_xml}")
            print(f"  Has Table Grid? {'Table Grid' in styles_xml}")
            if len(styles_xml) < 2000:
                print("  WARNING: styles.xml is tiny - file was rebuilt from scratch, formatting was lost.")
        except KeyError:
            print("  WARNING: no word/styles.xml")
        print()
        print("RULE: This script only READS. It never writes a new DOCX.")

if __name__ == "__main__":
    main()
