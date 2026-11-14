#!/usr/bin/env python3
"""
safe_pdf_read.py - Read a PDF with Python stdlib ONLY. READ-ONLY, never edit PDFs.
No PyMuPDF, no pdfminer, no Office. Uses re + zlib only.

Usage:
  python3 safe_pdf_read.py input.pdf

What it does:
  - Finds all stream objects, tries zlib decompress
  - Extracts text from (...) Tj and [...] TJ operators
  - Prints page count estimate and extracted text (first 8000 chars)

What it does NOT do:
  - It does NOT rebuild a PDF. PDFs are NOT edited by this skill.
  - If you need to change a PDF, edit the source DOCX with safe_docx_edit.py,
    then export PDF from Word on your computer. That is the only safe path.

Limitation: This is a best-effort extractor for simple PDFs (Times core fonts,
standard encoding). For complex PDFs with custom encodings, it may show
garbled text - that means do NOT trust it for content, ask the user for the DOCX source.
"""
import sys, re, zlib

def extract_streams(data):
    # Find stream ... endstream
    out = []
    for m in re.finditer(b"stream\r?\n(.*?)endstream", data, re.DOTALL):
        raw = m.group(1)
        # Try zlib decompress (FlateDecode)
        try:
            dec = zlib.decompress(raw)
            out.append(dec)
        except Exception:
            # Not compressed, or not Flate - keep raw if it looks like text
            if b"Tj" in raw or b"TJ" in raw:
                out.append(raw)
    return out

def extract_text_from_content(content):
    texts = []
    # ( ... ) Tj  - handle escaped parens
    # Simple approach: find all (...) groups
    for m in re.finditer(rb"\((?:[^()\\]|\\.)*\)", content):
        s = m.group(0)[1:-1]
        # Unescape
        s = s.replace(b"\\(", b"(").replace(b"\\)", b")").replace(b"\\\\", b"\\")
        try:
            texts.append(s.decode("latin-1"))
        except Exception:
            pass
    return " ".join(texts)

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 safe_pdf_read.py input.pdf")
        sys.exit(1)
    path = sys.argv[1]
    with open(path, "rb") as f:
        data = f.read()
    print(f"FILE: {path}  bytes={len(data)}")
    # Page count estimate: /Type /Page not /Pages
    pages = len(re.findall(rb"/Type\s*/Page[^s]", data))
    print(f"Estimated pages (by /Type /Page): {pages}")
    streams = extract_streams(data)
    print(f"Streams found: {len(streams)}  Decompressed: {len(streams)}")
    full = ""
    for s in streams:
        full += extract_text_from_content(s) + "\n"
    print(f"Extracted text chars: {len(full)}")
    print("---- FIRST 8000 CHARS ----")
    print(full[:8000])
    print("\n---- RULE ----")
    print("PDFs are READ-ONLY in this skill. Never rebuild a PDF.")
    print("To change content, edit the source DOCX with safe_docx_edit.py, then export PDF in Word.")

if __name__ == "__main__":
    main()
