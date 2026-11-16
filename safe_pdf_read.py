#!/usr/bin/env python3
"""
safe_pdf_read.py - Best-effort, READ-ONLY PDF text extraction. Python 3 stdlib only.

Usage:
  python3 safe_pdf_read.py input.pdf
  python3 safe_pdf_read.py input.pdf --max-chars 20000

How: parses PDF objects, inflates FlateDecode streams with zlib, maps
glyph codes to Unicode through each font's /ToUnicode CMap (the normal
case for Word and LibreOffice PDFs), and pulls text from Tj / TJ / ' / "
operators.

Limits: no OCR (scanned pages give no text), no compressed object streams
(/ObjStm) for fonts, no filters other than FlateDecode. When the result
looks garbled the script says so - then do NOT trust the text; ask for the
DOCX source or use a real PDF tool if your environment has one.

PDFs are never edited or rebuilt by this skill.
Exit status: 0 = text extracted, 3 = little/garbled text, 2 = not a PDF, 1 = error.
"""
import argparse, re, sys, zlib

OBJ_RE = re.compile(rb"(\d+)\s+\d+\s+obj\b(.*?)\bendobj", re.DOTALL)
STREAM_RE = re.compile(rb"stream\r?\n(.*?)\r?\n?endstream", re.DOTALL)
REF = rb"(\d+)\s+\d+\s+R"
NUM = rb"-?(?:\d+\.?\d*|\.\d+)"
TOKEN_RE = re.compile(
    rb"\((?:\\.|[^\\)])*\)"            # literal string (no nested parens)
    rb"|<[0-9A-Fa-f\s]*>"              # hex string
    rb"|/[^\s/<>\[\]()%]+"             # name, e.g. /F1
    rb"|\[|\]"
    rb"|" + NUM +
    rb"|[A-Za-z'\"*]+", re.DOTALL)
ESC = {b"n": b"\n", b"r": b"\r", b"t": b"\t", b"b": b"\b", b"f": b"\f",
       b"(": b"(", b")": b")", b"\\": b"\\"}


def literal_bytes(tok):
    s, out, i = tok[1:-1], bytearray(), 0
    while i < len(s):
        c = s[i:i + 1]
        if c == b"\\" and i + 1 < len(s):
            n = s[i + 1:i + 2]
            if n in ESC:
                out += ESC[n]; i += 2; continue
            m = re.match(rb"[0-7]{1,3}", s[i + 1:i + 4])
            if m:
                out.append(int(m.group(0), 8) & 0xFF); i += 1 + len(m.group(0)); continue
            i += 2 if n in b"\r\n" else 1
            continue
        out += c; i += 1
    return bytes(out)


def hex_bytes(tok):
    h = re.sub(rb"\s", b"", tok[1:-1])
    if len(h) % 2:
        h += b"0"
    return bytes.fromhex(h.decode())


def utf16(hexstr):
    try:
        return bytes.fromhex(hexstr.decode()).decode("utf-16-be", "replace")
    except ValueError:
        return ""


def parse_cmap(data):
    """ToUnicode CMap -> (code_width_bytes, {code_int: str})."""
    width = 2
    m = re.search(rb"begincodespacerange\s*<([0-9A-Fa-f]+)>", data)
    if m:
        width = max(1, len(m.group(1)) // 2)
    cmap = {}
    for block in re.findall(rb"beginbfchar(.*?)endbfchar", data, re.DOTALL):
        for src, dst in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>", block):
            cmap[int(src, 16)] = utf16(dst)
    for block in re.findall(rb"beginbfrange(.*?)endbfrange", data, re.DOTALL):
        for a, b, rest in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(\[[^\]]*\]|<[0-9A-Fa-f]+>)", block):
            lo, hi = int(a, 16), int(b, 16)
            if hi - lo > 65535:
                continue
            if rest.startswith(b"["):
                for k, d in enumerate(re.findall(rb"<([0-9A-Fa-f]*)>", rest)):
                    cmap[lo + k] = utf16(d)
            else:
                base = rest[1:-1]
                start = int(base, 16)
                for k in range(hi - lo + 1):
                    v = format(start + k, "0%dx" % len(base)).encode()
                    cmap[lo + k] = utf16(v)
    return width, cmap


def load_objects(raw):
    objs = {}
    for m in OBJ_RE.finditer(raw):
        body = m.group(2)
        s = STREAM_RE.search(body)
        d = body[:s.start()] if s else body
        stream = None
        if s:
            stream = s.group(1)
            if b"/FlateDecode" in d:
                try:
                    stream = zlib.decompressobj().decompress(stream)
                except zlib.error:
                    stream = None
            elif b"/Filter" in d:
                stream = None   # unsupported filter (images etc.)
        objs[int(m.group(1))] = (d, stream)
    return objs


def font_maps(objs):
    """Resource name (e.g. 'F1') -> (width, cmap). Names are merged across pages."""
    fonts = {}
    for d, _ in objs.values():
        for fm in re.finditer(rb"/Font\s*(<<.*?>>|" + REF + rb")", d, re.DOTALL):
            body = fm.group(1)
            if fm.group(2):
                body = objs.get(int(fm.group(2)), (b"", None))[0]
            for name, ref in re.findall(rb"/([^\s/<>\[\]()]+)\s+" + REF, body):
                fd = objs.get(int(ref), (b"", None))[0]
                tu = re.search(rb"/ToUnicode\s+" + REF, fd)
                if tu:
                    stream = objs.get(int(tu.group(1)), (b"", None))[1]
                    if stream:
                        fonts[name.decode("latin-1")] = parse_cmap(stream)
    return fonts


def decode_string(b, font, is_hex):
    if font:
        width, cmap = font
        out = []
        for i in range(0, len(b) - width + 1, width):
            out.append(cmap.get(int.from_bytes(b[i:i + width], "big"), ""))
        return "".join(out)
    if is_hex and len(b) >= 2 and len(b) % 2 == 0 and b[0::2].count(0) > len(b) // 4:
        return b.decode("utf-16-be", "replace")
    return b.decode("latin-1")


def text_from_content(data, fonts):
    out, arr, pending, font, last_name = [], None, [], None, None
    for tok in TOKEN_RE.findall(data):
        if tok[:1] in (b"(", b"<"):
            is_hex = tok[:1] == b"<"
            s = decode_string(hex_bytes(tok) if is_hex else literal_bytes(tok), font, is_hex)
            (arr if arr is not None else pending).append(s)
        elif tok[:1] == b"/":
            last_name = tok[1:].decode("latin-1")
        elif tok == b"[":
            arr = []
        elif tok == b"]":
            pending, arr = ["".join(arr or [])], None
        elif re.fullmatch(NUM, tok):
            if arr is not None and float(tok) < -200:   # big negative kern = word gap
                arr.append(" ")
        elif tok == b"Tf":
            font = fonts.get(last_name)
        elif tok in (b"Tj", b"TJ"):
            out.append("".join(pending)); pending = []
        elif tok in (b"'", b'"'):
            out.append("\n" + "".join(pending)); pending = []
        elif tok in (b"Td", b"TD", b"T*", b"ET"):
            out.append("\n"); pending = []
        else:
            pending = []
    text = "".join(out)
    return re.sub(r"\n{3,}", "\n\n", re.sub(r"[ \t]+\n", "\n", text))


def main():
    ap = argparse.ArgumentParser(description="Read-only PDF text extraction")
    ap.add_argument("pdf")
    ap.add_argument("--max-chars", type=int, default=0, help="truncate output (0 = all)")
    a = ap.parse_args()
    try:
        raw = open(a.pdf, "rb").read()
    except OSError as exc:
        print(f"ERROR: {exc}"); sys.exit(1)
    if not raw.lstrip().startswith(b"%PDF"):
        print("NOT A PDF (no %PDF header). You may only have a text extract, not the real file.")
        sys.exit(2)

    objs = load_objects(raw)
    fonts = font_maps(objs)
    pages = len(re.findall(rb"/Type\s*/Page(?![a-zA-Z])", raw))
    has_objstm = b"/ObjStm" in raw
    chunks = []
    for d, stream in objs.values():
        if stream and b"BT" in stream and not re.search(
                rb"/Subtype\s*/Image|/Length1|/Type\s*/(XRef|ObjStm)", d):
            chunks.append(text_from_content(stream, fonts))
    text = "\n".join(c.strip() for c in chunks if c.strip())

    printable = sum(ch.isprintable() or ch in "\n\t" for ch in text)
    ratio = printable / len(text) if text else 0
    letters = sum(ch.isalpha() for ch in text)
    print(f"=== {a.pdf} ===")
    print(f"pages (approx): {pages}   fonts with ToUnicode: {len(fonts)}   "
          f"text chars: {len(text):,d}   printable: {ratio:.0%}"
          + ("   (has compressed object streams: some fonts may be missed)" if has_objstm else ""))
    garbled = not text or ratio < 0.9 or letters < max(20, len(text) * 0.3)
    if garbled:
        print("WARNING: little or garbled text (scanned pages, custom encoding, or unsupported "
              "PDF features). Do NOT trust this output; ask for the DOCX source.")
    print("-" * 60)
    print(text[:a.max_chars] if a.max_chars else text)
    sys.exit(3 if garbled else 0)


if __name__ == "__main__":
    main()
