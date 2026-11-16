#!/usr/bin/env python3
"""
safe_docx_edit.py - Safe DOCX editor, v3.4. Python 3 standard library only.

Edits text inside an existing .docx WITHOUT rebuilding it:
  * copies every ZIP part byte-for-byte except word/document.xml
    (and docProps/core.xml when metadata keys are used);
  * inside word/document.xml, changes text only inside <w:t>...</w:t>
    using string surgery on the raw XML (never ElementTree re-serialization);
  * verifies the result with 12 gates and exits 1 (and deletes the output)
    if any gate fails.

Usage:
  python3 safe_docx_edit.py input.docx output.docx replacements.txt

replacements.txt - one pair per line, OLD<TAB>NEW, exact text.
  Blank lines and lines starting with # are ignored.
  Special keys (left column):
    __TITLE__            set document title    (requires __AUTHOR__ and __LASTMODIFIEDBY__ too)
    __AUTHOR__           set dc:creator
    __LASTMODIFIEDBY__   set cp:lastModifiedBy
    __PROTECT__          a literal whose count must only change as the
                         change list says (Gate 6). "A*" is always protected.

Behaviour:
  * Every occurrence of OLD in a paragraph is replaced (left to right; the
    inserted NEW text is never re-scanned, so NEW may contain OLD safely).
  * OLD may span several runs (Word splits runs). NEW goes into the first
    run, the rest of OLD is removed from the following runs; all run
    formatting, tags and attributes are kept.
  * Matching never crosses a paragraph boundary.
  * Every pair must match at least once, or the edit FAILS (catches typos).

Exit status: 0 = all gates passed; 1 = failure (output file is deleted).

Gates (all must PASS):
  1  Image bytes identical (full SHA-256)
  2  word/styles.xml byte-identical
  3  All header*/footer*/footnotes/endnotes parts byte-identical
  4  Namespace prefixes intact (r:embed, wp: counts; no ns0:/ns1:/ns2:)
  5  xml:space count unchanged
  6  Protected literals (A* + __PROTECT__) change exactly as the change list says
  7  Only document.xml (+ core.xml) changed; no parts added or removed
  8  No duplication: NEW count == before + applied
  9  No markup inside <w:t> (no phantom <w:tab/> swallowing)
  10 Shared package-integrity verifier (same code the regression tests use)
  11 Every replacement pair matched at least once
  12 Edited XML is well-formed and has no double-escaped entities

Scope of the guarantee: document-PART integrity. The ZIP container itself
is re-written (timestamps/compression), so do not claim bit-identical ZIPs.
"""
import os, re, sys, zipfile, hashlib
from datetime import datetime, timezone
from xml.dom import minidom  # parse-only check for Gate 12; never re-serialized

VERSION = "3.4"
DEFAULT_PROTECTED = ["A*"]
META_KEYS = ("__TITLE__", "__AUTHOR__", "__LASTMODIFIEDBY__")


def sha(b):
    return hashlib.sha256(b).hexdigest()


# --- XML text helpers -------------------------------------------------------
_ENTITY_RE = re.compile(r"&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);")
_NAMED = {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'"}


def xml_unescape(s):
    """Decode ALL XML entities (named + numeric). v3.4: v3.3 decoded only
    &amp; &lt; &gt;, so &quot;/&apos;/&#8217; were later re-escaped to
    &amp;quot; etc. = visible corruption that passed every gate."""
    def _d(m):
        e = m.group(1)
        if e.startswith("#x"):
            return chr(int(e[2:], 16))
        if e.startswith("#"):
            return chr(int(e[1:]))
        return _NAMED[e]
    return _ENTITY_RE.sub(_d, s)


def xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# <w:t> or <w:t attr...> - whitespace required so <w:tab/> can never match.
WT_RE = re.compile(r"<w:t(\s[^>]*)?>(.*?)</w:t>", re.DOTALL)
# A paragraph. v3.4: (?<!/) excludes self-closing <w:p .../> which in v3.3
# could swallow the following paragraph into the same block.
P_RE = re.compile(r"<w:p(?:\s[^>]*)?(?<!/)>.*?</w:p>", re.DOTALL)


def paragraph_texts(xml_str):
    """Decoded visible text of each paragraph (runs joined)."""
    return ["".join(xml_unescape(m.group(2)) for m in WT_RE.finditer(p.group(0)))
            for p in P_RE.finditer(xml_str)]


def count_text(xml_str, needle):
    if not needle:
        return 0
    return sum(t.count(needle) for t in paragraph_texts(xml_str))


# --- Replacement engine -----------------------------------------------------
def _replace_at(block, old, new, search_from, log, warnings):
    """Replace the first occurrence of OLD at/after search_from (in decoded
    paragraph text). Returns (new_block, index_after_NEW) or (block, None)."""
    matches = list(WT_RE.finditer(block))
    if not matches:
        return block, None
    inners = [xml_unescape(m.group(2)) for m in matches]
    full = "".join(inners)
    start = full.find(old, search_from)
    if start < 0:
        return block, None
    end = start + len(old)

    starts, pos = [], 0
    for t in inners:
        starts.append(pos)
        pos += len(t)
    overlapping = [i for i, t in enumerate(inners)
                   if starts[i] < end and starts[i] + len(t) > start]

    new_inners = {}
    for k, mi in enumerate(overlapping):
        inner, g0 = inners[mi], starts[mi]
        before = inner[:start - g0] if g0 < start else ""
        after = inner[end - g0:] if g0 + len(inner) > end else ""
        if k == 0:
            text = before + new + (after if len(overlapping) == 1 else "")
        elif k == len(overlapping) - 1:
            text = after
        else:
            text = ""
        attrs = matches[mi].group(1) or ""
        if text and text != text.strip() and "xml:space" not in attrs:
            warnings.append(f"run text {text[:40]!r} has leading/trailing space but no "
                            f'xml:space="preserve"; Word may drop that space - check visually')
        new_inners[mi] = xml_escape(text)

    out = block
    for mi in sorted(overlapping, reverse=True):
        m = matches[mi]
        out = out[:m.start(2)] + new_inners[mi] + out[m.end(2):]
    log.append(f"  replaced {old[:60]!r} -> {new[:60]!r} (runs {overlapping})")
    return out, start + len(new)


def replace_in_paragraph(block, old, new, log, warnings):
    """Replace ALL occurrences of OLD in one paragraph without re-scanning
    inserted text. Returns (block, count)."""
    n, pos = 0, 0
    while True:
        block, nxt = _replace_at(block, old, new, pos, log, warnings)
        if nxt is None:
            return block, n
        n += 1
        pos = nxt


def replace_in_document_xml(xml_str, reps, log, warnings, applied):
    nested_re = re.compile(r"<w:p[\s>]|<w:txbxContent")

    def repl_para(m):
        block = m.group(0)
        if nested_re.search(block, 4):
            # A text box / shape holds paragraphs inside this paragraph. The
            # regex cannot see its true end, and Word stores text boxes twice
            # (DrawingML + VML fallback). Editing here is unsafe: skip it.
            msg = ("skipped a paragraph containing a text box/shape; "
                   "edit text boxes in Word (list them in the change list)")
            if msg not in warnings:
                warnings.append(msg)
            return block
        for old, new in reps:
            block, n = replace_in_paragraph(block, old, new, log, warnings)
            if n:
                applied[old] = applied.get(old, 0) + n
        return block
    return P_RE.sub(repl_para, xml_str)


# --- Inputs -----------------------------------------------------------------
def load_replacements(path):
    reps, meta, protect, errors = [], {}, list(DEFAULT_PROTECTED), []
    with open(path, encoding="utf-8-sig") as f:   # -sig: tolerate a BOM
        for ln, line in enumerate(f, 1):
            line = line.rstrip("\r\n")
            if not line.strip() or line.startswith("#"):
                continue
            if "\t" not in line:
                errors.append(f"line {ln}: no TAB between OLD and NEW")
                continue
            old, new = line.split("\t", 1)
            if old in META_KEYS:
                meta[old] = new
            elif old == "__PROTECT__":
                if new and new not in protect:
                    protect.append(new)
            elif old == "":
                errors.append(f"line {ln}: OLD is empty")
            else:
                reps.append((old, new))
    return reps, meta, protect, errors


def apply_metadata(core_xml, meta, warnings):
    out = core_xml
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    def set_tag(tag, value):
        nonlocal out
        repl = f"<{tag}>{xml_escape(value)}</{tag}>"
        if re.search(rf"<{tag}\s*/>", out):
            out = re.sub(rf"<{tag}\s*/>", lambda m: repl, out, count=1)
        elif re.search(rf"<{tag}(\s[^>]*)?>.*?</{tag}>", out, flags=re.DOTALL):
            # lambda, never a template string: backslashes in titles are safe
            out = re.sub(rf"<{tag}(\s[^>]*)?>.*?</{tag}>", lambda m: repl, out,
                         count=1, flags=re.DOTALL)
        else:
            m = re.search(r"<cp:coreProperties[^>]*>", out)
            if not m:
                raise ValueError(f"cannot insert <{tag}>: no <cp:coreProperties> root")
            out = out[:m.end()] + repl + out[m.end():]
            print(f"  metadata: inserted missing <{tag}>")

    for key, tag in (("__TITLE__", "dc:title"), ("__AUTHOR__", "dc:creator"),
                     ("__LASTMODIFIEDBY__", "cp:lastModifiedBy")):
        if key in meta:
            set_tag(tag, meta[key])
    if re.search(r"<dcterms:modified[^>]*>.*?</dcterms:modified>", out, flags=re.DOTALL):
        out = re.sub(r"(<dcterms:modified[^>]*>).*?(</dcterms:modified>)",
                     lambda m: m.group(1) + now + m.group(2), out, count=1, flags=re.DOTALL)
    else:
        warnings.append("docProps/core.xml has no <dcterms:modified>; modified date not updated")
    return out


# --- Verification -----------------------------------------------------------
PROTECTED_PART_RE = re.compile(r"word/(header\d*|footer\d*|footnotes|endnotes)\.xml$")


def verify_package_integrity(inp, outp):
    """Package-level violations between original and edited DOCX. Shared by
    the editor (Gate 10) and the regression tests."""
    violations = []
    try:
        with zipfile.ZipFile(inp) as a, zipfile.ZipFile(outp) as b:
            na, nb = set(a.namelist()), set(b.namelist())
            added, removed = sorted(nb - na), sorted(na - nb)
            if added or removed:
                violations.append(f"ZIP parts added={added} removed={removed}")
            for name in sorted(na & nb):
                x, y = a.read(name), b.read(name)
                if x == y:
                    continue
                if name.startswith("word/media/"):
                    violations.append(f"image digest mismatch: {name}")
                elif name == "word/styles.xml":
                    violations.append("word/styles.xml changed")
                elif PROTECTED_PART_RE.match(name):
                    violations.append(f"protected part changed: {name}")
                if name not in ("word/document.xml", "docProps/core.xml"):
                    violations.append(f"unexpected changed part: {name}")
    except (OSError, zipfile.BadZipFile, KeyError) as exc:
        violations.append(f"cannot verify DOCX package: {exc}")
    return violations


DOUBLE_ESC_RE = re.compile(r"&amp;(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);")


def run_gates(inp, outp, reps, protect, applied, edited_parts):
    ok = True

    def gate(name, cond, detail=""):
        nonlocal ok
        print(f"{'PASS' if cond else 'FAIL'}: {name} {detail}".rstrip())
        ok = ok and bool(cond)

    with zipfile.ZipFile(inp) as a, zipfile.ZipFile(outp) as b:
        an, bn = a.namelist(), set(b.namelist())
        # 1
        for n in an:
            if n.startswith("word/media/"):
                same = n in bn and sha(a.read(n)) == sha(b.read(n))
                gate(f"Gate 1 image {n}", same, "" if n in bn else "MISSING IN OUTPUT")
        # 2
        if "word/styles.xml" in an:
            gate("Gate 2 styles.xml identical",
                 "word/styles.xml" in bn and a.read("word/styles.xml") == b.read("word/styles.xml"))
        # 3
        for p in (n for n in an if PROTECTED_PART_RE.match(n)):
            gate(f"Gate 3 {p} untouched", p in bn and a.read(p) == b.read(p),
                 "" if p in bn else "MISSING IN OUTPUT")
        doc_o = a.read("word/document.xml").decode("utf-8")
        doc_n = b.read("word/document.xml").decode("utf-8")
        # 4
        for pat in ("r:embed", "wp:"):
            gate(f"Gate 4 {pat} count unchanged", doc_o.count(pat) == doc_n.count(pat),
                 f"{doc_o.count(pat)}->{doc_n.count(pat)}")
        bad_ns = [p for p in ("ns0:", "ns1:", "ns2:") if p in doc_n and p not in doc_o]
        gate("Gate 4 no new ns0:/ns1:/ns2: prefixes", not bad_ns, str(bad_ns) if bad_ns else "")
        # 5
        gate("Gate 5 xml:space unchanged", doc_o.count("xml:space") == doc_n.count("xml:space"),
             f"{doc_o.count('xml:space')}->{doc_n.count('xml:space')}")
        # 6
        for lit in protect:
            expected = sum((n.count(lit) - o.count(lit)) * applied.get(o, 0) for o, n in reps)
            observed = count_text(doc_n, lit) - count_text(doc_o, lit)
            gate(f"Gate 6 protected {lit!r} delta", observed == expected,
                 f"observed {observed}, expected {expected}")
        # 7
        changed = [n for n in an if n in bn and a.read(n) != b.read(n)]
        added = sorted(bn - set(an)); removed = [n for n in an if n not in bn]
        gate("Gate 7 no parts added/removed", not added and not removed,
             f"added={added} removed={removed}" if added or removed else "")
        gate("Gate 7 only allowed parts changed", set(changed) <= set(edited_parts), str(changed))
        # 8
        for o, n in reps:
            if n == o or not n:
                continue
            before, after, k = count_text(doc_o, n), count_text(doc_n, n), applied.get(o, 0)
            gate(f"Gate 8 no duplication {n[:40]!r}", after == before + k,
                 f"before {before} + applied {k} = {before + k}, found {after}")
        # 9
        inners = [m.group(2) for m in WT_RE.finditer(doc_n)]
        gate("Gate 9 no markup inside <w:t>", not any("<" in t for t in inners),
             f"checked {len(inners)} runs")
    # 10
    v = verify_package_integrity(inp, outp)
    gate("Gate 10 package verifier", not v, "; ".join(v))
    # 11
    missed = [o for o, _ in reps if applied.get(o, 0) == 0]
    gate("Gate 11 every pair matched", not missed,
         ("NOT FOUND: " + "; ".join(repr(o[:60]) for o in missed)) if missed else "")
    # 12
    with zipfile.ZipFile(inp) as a, zipfile.ZipFile(outp) as b:
        for part in edited_parts:
            x, y = a.read(part).decode("utf-8"), b.read(part).decode("utf-8")
            try:
                minidom.parseString(y.encode("utf-8"))
                wf = True; err = ""
            except Exception as exc:  # noqa: BLE001
                wf = False; err = str(exc)[:120]
            gate(f"Gate 12 {part} well-formed", wf, err)
            d0, d1 = len(DOUBLE_ESC_RE.findall(x)), len(DOUBLE_ESC_RE.findall(y))
            gate(f"Gate 12 {part} no new double-escaped entities", d1 <= d0, f"{d0}->{d1}")
    return ok


# --- Main -------------------------------------------------------------------
def fail(msg, outp=None):
    print(f"\nFAIL: {msg}")
    if outp and os.path.exists(outp):
        os.remove(outp)
        print(f"Deleted {outp}. Do NOT send a file; return a change list instead.")
    sys.exit(1)


def main():
    print(f"safe_docx_edit v{VERSION}")
    if len(sys.argv) != 4:
        print("Usage: python3 safe_docx_edit.py input.docx output.docx replacements.txt")
        sys.exit(1)
    inp, outp, rep_path = sys.argv[1:4]

    if os.path.abspath(inp) == os.path.abspath(outp) or (
            os.path.exists(outp) and os.path.samefile(inp, outp)):
        fail("output path is the input file. Always write to a COPY; the master is never touched.")
    if not zipfile.is_zipfile(inp):
        fail(f"{inp} is not a ZIP/DOCX. You may only have a text extract, not the real file bytes.")

    reps, meta, protect, errors = load_replacements(rep_path)
    if errors:
        fail("bad replacements file:\n  " + "\n  ".join(errors))
    if not reps and not meta:
        fail("replacements file has no changes")
    if "__TITLE__" in meta and not {"__AUTHOR__", "__LASTMODIFIEDBY__"} <= meta.keys():
        fail("__TITLE__ also needs __AUTHOR__ and __LASTMODIFIEDBY__ (no half-done metadata)")

    print(f"Loaded {len(reps)} text pair(s), {len(meta)} metadata key(s); protected: {protect}")
    for o, n in reps:
        print(f"  {o[:70]!r} -> {n[:70]!r}")

    with zipfile.ZipFile(inp) as zin:
        infos = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in infos}
    if "word/document.xml" not in data:
        fail("no word/document.xml in input - not a Word document")
    if meta and "docProps/core.xml" not in data:
        fail("metadata keys given but the file has no docProps/core.xml")

    log, warnings, applied, edited = [], [], {}, []
    try:
        if reps:
            doc = data["word/document.xml"].decode("utf-8")
            data["word/document.xml"] = replace_in_document_xml(
                doc, reps, log, warnings, applied).encode("utf-8")
            edited.append("word/document.xml")
        if meta:
            core = data["docProps/core.xml"].decode("utf-8")
            data["docProps/core.xml"] = apply_metadata(core, meta, warnings).encode("utf-8")
            edited.append("docProps/core.xml")

        with zipfile.ZipFile(outp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in infos:   # same order, same names, untouched bytes
                zout.writestr(info.filename, data[info.filename],
                              compress_type=info.compress_type)
        print(f"Wrote {outp}\nApplied counts: {applied}")

        print("\n=== GATES ===")
        ok = run_gates(inp, outp, reps, protect, applied, edited)
    except Exception as exc:  # noqa: BLE001 - any crash = failed edit
        fail(f"{type(exc).__name__}: {exc}", outp)

    print("\nEdit log:")
    for line in log or ["  (nothing replaced)"]:
        print(line)
    for w in warnings:
        print(f"WARN: {w}")
    if not ok:
        fail("one or more gates failed - output discarded.", outp)
    print("\nALL GATES PASS. Open the copy in Word for a final visual check before sending.")
    sys.exit(0)


if __name__ == "__main__":
    main()
