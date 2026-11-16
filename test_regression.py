#!/usr/bin/env python3
"""
test_regression.py - Regression tests for the Safe DOCX skill (stdlib only).

Run:  python3 test_regression.py        Exit 0 = all pass, 1 = any failure.

Editor (safe_docx_edit.py):
   1  happy path: all gates pass, header3.xml discovered and checked
   2  cross-run replacement applied, run formatting kept
   3  invalid (non-ZIP) input -> exit 1
   4-7 tamper detection: header3 changed, part removed, part added, image byte flipped
   8  &quot; / &apos; / &#8217; in an edited run are NOT double-escaped      (v3.4)
   9  every occurrence in a paragraph is replaced, NEW containing OLD is safe (v3.4)
  10  a pair that matches nothing FAILS and the output file is deleted      (v3.4)
  11  output path == input path is refused; master untouched               (v3.4)
  12  self-closing <w:p/> does not merge two paragraphs                     (v3.4)
  13  __TITLE__ alone is refused (no half-done metadata)                    (v3.4)
  14  title with backslash + & written correctly, modified date updated     (v3.2)
  15  intended A -> A* change passes Gate 6; xml:space untouched
Readers:
  16  safe_docx_read.py: exit 0 on DOCX, outline shows headings
  17  safe_docx_read.py: exit 2 on non-ZIP (text-only stub)
  18  safe_pdf_read.py: literal + ToUnicode hex text extracted
  19  safe_pdf_read.py: exit 2 on non-PDF
Editor, continued:
  20  paragraphs holding a text box are skipped with a warning           (v3.4)
"""
import os, subprocess, sys, tempfile, zipfile, importlib.util

HERE = os.path.dirname(os.path.abspath(__file__))
EDITOR = os.path.join(HERE, "safe_docx_edit.py")
READER = os.path.join(HERE, "safe_docx_read.py")
PDFREADER = os.path.join(HERE, "safe_pdf_read.py")
spec = importlib.util.spec_from_file_location("safe_docx_editor", EDITOR)
editor = importlib.util.module_from_spec(spec)
spec.loader.exec_module(editor)

PNG = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108020000"
    "907753de0000000c4944415408d763f8cfc0000003010100c9fe92ef"
    "0000000049454e44ae426082")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

DEFAULT_BODY = (
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Introduction</w:t></w:r></w:p>'
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">The governing</w:t></w:r>'
    '<w:r><w:t xml:space="preserve"> law of New York shall apply.</w:t></w:r></w:p>')


def build_docx(path, body=DEFAULT_BODY, core_inner="<dc:title>Old Title</dc:title>"
               '<dcterms:modified xsi:type="dcterms:W3CDTF">2013-01-01T00:00:00Z</dcterms:modified>'):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Default Extension="png" ContentType="image/png"/></Types>')
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           f'<w:document {W} {R}><w:body>{body}'
           '<w:sectPr><w:headerReference w:type="first" r:id="rIdHdr3"/>'
           '<w:headerReference w:type="default" r:id="rIdHdr1"/></w:sectPr></w:body></w:document>')
    styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles {W}>'
              '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
              '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>'
              '</w:styles>')
    h1 = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr {W}>'
          '<w:p><w:r><w:t>CONFIDENTIAL HEADER</w:t></w:r></w:p></w:hdr>')
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            f'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">{core_inner}</cp:coreProperties>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/theme/theme1.xml", "<a:theme/>")
        z.writestr("word/header1.xml", h1)
        z.writestr("word/header3.xml", h1.replace("CONFIDENTIAL", "FIRST PAGE"))
        z.writestr("word/media/image1.png", PNG)
        z.writestr("docProps/core.xml", core)


def run(script, *args):
    r = subprocess.run([sys.executable, "-I", script, *args], capture_output=True, text=True, timeout=60)
    return r.stdout + r.stderr, r.returncode


def edit(tmp, name, reps_text, body=DEFAULT_BODY, **kw):
    inp, outp, reps = (os.path.join(tmp, f"{name}{s}") for s in (".docx", "_out.docx", ".txt"))
    build_docx(inp, body, **kw) if kw else build_docx(inp, body)
    with open(reps, "w", encoding="utf-8") as f:
        f.write(reps_text)
    log, code = run(EDITOR, inp, outp, reps)
    return inp, outp, log, code


def part(path, name):
    with zipfile.ZipFile(path) as z:
        return z.read(name).decode("utf-8")


def tamper(src, dst, mode):
    with zipfile.ZipFile(src) as zin:
        items = {n: zin.read(n) for n in zin.namelist()}
    if mode == "tamper_header3":
        items["word/header3.xml"] = items["word/header3.xml"].replace(b"FIRST PAGE", b"TAMPERED")
    elif mode == "remove_part":
        del items["word/header3.xml"]
    elif mode == "add_part":
        items["word/evil_extra.xml"] = b"<evil/>"
    elif mode == "corrupt_image":
        items["word/media/image1.png"] = items["word/media/image1.png"][:-1] + b"\x00"
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in items.items():
            z.writestr(n, b)


def build_pdf(path):
    """Minimal PDF: one literal-string run (Helvetica) + one hex run via ToUnicode."""
    cmap = (b"/CIDInit /ProcSet findresource begin 12 dict begin begincmap\n"
            b"1 begincodespacerange <0000> <FFFF> endcodespacerange\n"
            b"2 beginbfchar <0001> <0048> <0002> <0069> endbfchar\n"
            b"1 beginbfrange <0003> <0004> <002A> endbfrange\n"
            b"endcmap CMapName currentdict /CMap defineresource pop end end")
    content = b"BT /F1 12 Tf 72 720 Td (Hello \\(PDF\\) world) Tj 0 -20 Td /F2 12 Tf <000100020003> Tj ET"
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R /F2 6 0 R >> >> /Contents 4 0 R >>",
        b"<< /Length %d >>\nstream\n" % len(content) + content + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Type /Font /Subtype /Type0 /BaseFont /X /ToUnicode 7 0 R >>",
        b"<< /Length %d >>\nstream\n" % len(cmap) + cmap + b"\nendstream",
    ]
    out = b"%PDF-1.4\n"
    for i, o in enumerate(objs, 1):
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    out += b"trailer << /Root 1 0 R >>\n%%EOF\n"
    with open(path, "wb") as f:
        f.write(out)


def main():
    tmp = tempfile.mkdtemp(prefix="sdx_test_")
    results = []

    # 1-2
    inp, outp, log, code = edit(tmp, "happy", "The governing law of New York\tThe governing law of Delaware\n")
    results.append(("1 happy path: all gates pass, header3.xml checked",
                    code == 0 and "ALL GATES PASS" in log and "Gate 3 word/header3.xml untouched" in log))
    doc = part(outp, "word/document.xml") if os.path.exists(outp) else ""
    results.append(("2 cross-run replacement applied, bold run kept",
                    "New York" not in doc and "Delaware" in doc and "<w:b/>" in doc))
    happy_in, happy_out = inp, outp

    # 3
    bad = os.path.join(tmp, "not_a_docx.docx")
    open(bad, "wb").write(b"not a zip archive")
    _, c = run(EDITOR, bad, os.path.join(tmp, "x.docx"), os.path.join(tmp, "happy.txt"))
    results.append(("3 invalid input -> exit 1", c == 1))

    # 4-7
    for mode, label in [("tamper_header3", "4 tampered header3.xml detected"),
                        ("remove_part", "5 missing part detected"),
                        ("add_part", "6 extra part detected"),
                        ("corrupt_image", "7 corrupted image byte detected")]:
        tp = os.path.join(tmp, f"tam_{mode}.docx")
        tamper(happy_out, tp, mode)
        results.append((label, len(editor.verify_package_integrity(happy_in, tp)) > 0))

    # 8 entities
    body = ('<w:p><w:r><w:t xml:space="preserve">She said &quot;hi&quot; &amp; it&apos;s &#8217;ok&#8217; to</w:t></w:r>'
            '<w:r><w:t xml:space="preserve"> the team</w:t></w:r></w:p>')
    _, outp, log, code = edit(tmp, "ent", "to the team\tto everyone\n", body)
    doc = part(outp, "word/document.xml") if os.path.exists(outp) else ""
    results.append(("8 entities not double-escaped (&quot; &apos; &#8217;)",
                    code == 0 and "&amp;quot;" not in doc and "&amp;apos;" not in doc
                    and "&amp;#" not in doc and "She said \"hi\" &amp; it's ’ok’ to everyone" in doc))

    # 9 all occurrences, NEW contains OLD
    body = '<w:p><w:r><w:t>cat and cat and cat</w:t></w:r></w:p>'
    _, outp, log, code = edit(tmp, "multi", "cat\tcat (feline)\n", body)
    doc = part(outp, "word/document.xml") if os.path.exists(outp) else ""
    results.append(("9 all occurrences replaced once each",
                    code == 0 and "cat (feline) and cat (feline) and cat (feline)<" in doc))

    # 10 unmatched pair fails + output deleted
    _, outp, log, code = edit(tmp, "miss", "Introduction\tIntro\nTypo not in doc\tX\n")
    results.append(("10 unmatched pair -> exit 1, output deleted",
                    code == 1 and "Gate 11" in log and not os.path.exists(outp)))

    # 11 same path refused
    same = os.path.join(tmp, "same.docx"); build_docx(same)
    before = open(same, "rb").read()
    reps = os.path.join(tmp, "same.txt"); open(reps, "w").write("Introduction\tIntro\n")
    log, code = run(EDITOR, same, same, reps)
    results.append(("11 output == input refused, master untouched",
                    code == 1 and open(same, "rb").read() == before))

    # 12 self-closing paragraph
    body = ('<w:p w:rsidR="00A1"/><w:p><w:r><w:t>Alpha</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Beta</w:t></w:r></w:p>')
    _, outp, log, code = edit(tmp, "selfclose", "AlphaBeta\tGAMMA\n", body)
    results.append(("12 no match across paragraphs (self-closing <w:p/>)", code == 1 and "GAMMA" not in log.split("Edit log")[-1]))

    # 20 text-box paragraphs skipped (nested <w:p>), rest still edited
    body = ('<w:p><w:r><w:t>Before box</w:t></w:r><w:r><w:drawing><w:txbxContent><w:p><w:r>'
            '<w:t>Box text</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r></w:p>'
            '<w:p><w:r><w:t>Normal para</w:t></w:r></w:p>')
    _, outp, log, code = edit(tmp, "tbox", "Normal para\tFixed para\n", body)
    ok_a = code == 0 and "text box" in log
    _, outp, log, code = edit(tmp, "tbox2", "Box text\tX\n", body)
    results.append(("20 text-box paragraphs skipped with warning; pair only in box fails",
                    ok_a and code == 1 and "Gate 11" in log))

    # 13 title alone refused
    _, outp, log, code = edit(tmp, "title1", "__TITLE__\tNew\n")
    results.append(("13 __TITLE__ without author/lastModifiedBy refused", code == 1))

    # 14 backslash + & title
    _, outp, log, code = edit(tmp, "title2", "__TITLE__\tPath\\model & results\n__AUTHOR__\tA. Author\n"
                                             "__LASTMODIFIEDBY__\tA. Author\n")
    core = part(outp, "docProps/core.xml") if os.path.exists(outp) else ""
    results.append(("14 metadata: backslash/& title, creator inserted, modified updated",
                    code == 0 and "<dc:title>Path\\model &amp; results</dc:title>" in core
                    and "<dc:creator>A. Author</dc:creator>" in core and "2013-01-01" not in core))

    # 15 intended A* change
    body = '<w:p><w:r><w:t xml:space="preserve">We used 57 A and 184 A journals.</w:t></w:r></w:p>'
    _, outp, log, code = edit(tmp, "astar", "57 A and 184 A journals\t57 A* and 184 A* journals\n", body)
    results.append(("15 intended A -> A* passes Gate 6", code == 0 and "PASS: Gate 6" in log))

    # 16-17 docx reader
    log, code = run(READER, happy_in)
    results.append(("16 docx reader: exit 0, heading in outline", code == 0 and "[heading 1] Introduction" in log))
    _, code = run(READER, bad)
    results.append(("17 docx reader: exit 2 on non-ZIP", code == 2))

    # 18-19 pdf reader
    pdf = os.path.join(tmp, "t.pdf"); build_pdf(pdf)
    log, code = run(PDFREADER, pdf)
    results.append(("18 pdf reader: literal + ToUnicode text", "Hello (PDF) world" in log and "\nHi*" in log))
    _, code = run(PDFREADER, bad)
    results.append(("19 pdf reader: exit 2 on non-PDF", code == 2))

    print("\n=== REGRESSION TESTS (v3.4) ===")
    failed = 0
    for label, ok in results:
        print(f"{'PASS' if ok else 'FAIL'}: {label}")
        failed += not ok
    print(f"\n{len(results) - failed}/{len(results)} passed.")
    if failed:
        print(f"FAIL: fix before distribution. Temp files kept in {tmp}")
        sys.exit(1)
    print("ALL REGRESSION TESTS PASS.")


if __name__ == "__main__":
    main()
