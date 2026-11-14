#!/usr/bin/env python3
"""
safe_docx_edit_v3_2.py - v3.2 editor: PURE STRING SURGERY, no ElementTree.
Fixes 5 latent bugs (v3.1) + 3 more (v3.2): backslash crash in metadata re.sub (lambda fix), fragile 60-char log counting (direct counter), Gate 8 false-FAIL on & / < (count escaped form).
Python stdlib ONLY. No python-docx, no Office.

WHY v3: v2 used ElementTree re-serialization, which renamed namespaces
(r:embed -> ns1:, wp: -> ns2:) and rewrote header/footer even with zero edits.
v3 never parses XML into a tree. It only splices text inside <w:t>...</w:t>
via regex on raw bytes. Everything else stays byte-identical.

RULES (NEVER AGAIN):
- Only word/document.xml is edited by default (+ docProps/core.xml for metadata).
- word/header*.xml, word/footer*.xml, footnotes, endnotes are UNTOUCHABLE.
  Gate 3 enforces this: they must be byte-identical. If you need a header edit,
  return a change list, do not extend without explicit approval + same gates.
- Cross-run supported: if OLD spans multiple <w:t> in same paragraph, NEW goes
  in first overlapping <w:t>, others cleared, first run formatting kept.
  All <w:t> tags and attributes (including xml:space) are preserved verbatim.

Usage:
  python3 safe_docx_edit_v3_2.py input.docx output.docx replacements.txt

replacements.txt: OLD<TAB>NEW per line. For metadata, use special keys:
  __TITLE__<TAB>New Title
  __AUTHOR__<TAB>New Author
  __LASTMODIFIEDBY__<TAB>New Name

Gates v3.2 (9 gates, all must PASS or DISCARD):
  1. Image SHA identical
  2. styles.xml byte-identical
  3. header/footer byte-identical (untouched)
  4. Namespaces intact: r:embed unchanged, wp: unchanged, zero ns0:/ns1: prefixes
  5. xml:space count unchanged
  6. A* DELTA enforced (observed == expected from change list)
  7. Only document.xml + core.xml changed
  8. No duplication verified (NEW count == before + applied)
  9. No phantom <w:tab/> corruption (no </w:r> inside <w:t>)
"""
import sys, zipfile, hashlib, re
from datetime import datetime, timezone

def sha(b): return hashlib.sha256(b).hexdigest()[:16]

def xml_escape(s):
    return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

def xml_unescape(s):
    return s.replace("&lt;","<").replace("&gt;",">").replace("&amp;","&")

WT_RE = re.compile(r"<w:t(\s[^>]*)?>(.*?)</w:t>", re.DOTALL)  # v3.1: \s required -> <w:tab/> cannot match
P_RE = re.compile(r"<w:p(?:\s[^>]*)?>.*?</w:p>", re.DOTALL)

def replace_in_paragraph_block(block, old, new, log):
    # Find all w:t in this block
    matches = list(WT_RE.finditer(block))
    if not matches:
        return block, False
    # Build concatenated unescaped text + map each char to (match_idx, offset_in_match)
    full_parts = []
    mapping = []  # per char: (match_idx, offset)
    for mi, m in enumerate(matches):
        inner = xml_unescape(m.group(2))
        full_parts.append(inner)
        for off in range(len(inner)):
            mapping.append((mi, off))
    full = "".join(full_parts)
    if old not in full:
        return block, False
    # Only first occurrence per paragraph per replacement (call again for multiple)
    start = full.find(old)
    end = start + len(old)
    # Determine overlapping w:t indices
    overlapping = sorted(set(mapping[i][0] for i in range(start, end)))
    # Build new inner texts for each overlapping w:t
    # For each overlapping w:t, compute its slice of full that overlaps old
    new_inners = {}
    for mi in overlapping:
        m = matches[mi]
        inner = xml_unescape(m.group(2))
        # Find global range of this w:t in full
        # Compute global start of this w:t
        g_start = sum(len(xml_unescape(matches[j].group(2))) for j in range(mi))
        g_end = g_start + len(inner)
        # Portion before old, after old, inside this w:t
        keep_before = ""
        keep_after = ""
        if g_start < start:
            keep_before = inner[:start - g_start]
        if g_end > end:
            keep_after = inner[end - g_start:]
        if mi == overlapping[0]:
            # First: keep_before + NEW + (keep_after if this is also last)
            if len(overlapping) == 1:
                new_inner = keep_before + new + keep_after
            else:
                new_inner = keep_before + new
        elif mi == overlapping[-1]:
            new_inner = keep_after
        else:
            new_inner = ""
        new_inners[mi] = xml_escape(new_inner)
    # Splice block: rebuild by replacing each overlapping w:t's inner via string positions
    # Work backwards so offsets stay valid
    new_block = block
    for mi in sorted(overlapping, reverse=True):
        m = matches[mi]
        # m.start(2), m.end(2) are inner spans in original block
        new_block = new_block[:m.start(2)] + new_inners[mi] + new_block[m.end(2):]
    log.append(f"  replaced {repr(old[:60])} -> {repr(new[:60])} (overlapping w:t={overlapping})")
    return new_block, True

def replace_in_document_xml(xml_str, reps, log, applied_counts=None):
    # Process paragraph by paragraph to avoid cross-paragraph matches
    def repl_para(m):
        block = m.group(0)
        for old, new in reps:
            # Replace ONCE per paragraph per replacement. Do NOT loop:
            # NEW often contains OLD (e.g. OLD + " Table 8..."), so re-scanning
            # NEW would duplicate the insertion 10 times. One pass is correct.
            # v3.2: direct counter, NOT log-text parsing (60-char prefix collision)
            block, _changed = replace_in_paragraph_block(block, old, new, log)
            if _changed and applied_counts is not None:
                applied_counts[old] = applied_counts.get(old, 0) + 1
        return block
    # Use P_RE to find paragraphs; also handle text outside <w:p> (e.g., in tables? tables contain w:p, so covered)
    # But some w:t may be outside w:p (rare). Fall back: if no w:p, do global w:t pass.
    if "<w:p" in xml_str:
        return P_RE.sub(repl_para, xml_str)
    else:
        # Global fallback (should not happen for Word)
        block, _ = replace_in_paragraph_block(xml_str, reps[0][0], reps[0][1], log) if reps else (xml_str, False)
        return block

def load_replacements(path):
    reps = []
    meta = {}
    with open(path, encoding="utf-8") as f:
        for ln, line in enumerate(f, 1):
            line = line.rstrip("\n")
            if not line or line.startswith("#"):
                continue
            if "\t" not in line:
                print(f"WARNING line {ln} no TAB, skipped")
                continue
            old, new = line.split("\t", 1)
            if old in ("__TITLE__", "__AUTHOR__", "__LASTMODIFIEDBY__"):
                meta[old] = new
            else:
                reps.append((old, new))
    return reps, meta

def apply_metadata(core_xml, meta):
    out = core_xml
    now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    def set_or_insert(tag_open, tag_close, value, insert_after=None):
        nonlocal out
        esc = xml_escape(value)
        # Case 1: self-closing <tag/>
        if f"<{tag_open}/>" in out:
            out = out.replace(f"<{tag_open}/>", f"<{tag_open}>{esc}</{tag_close}>")
            return True
        # Case 2: existing <tag>...</tag> - v3.2: use lambda, NOT f-string as re.sub template
        # f-string template crashes on backslash (e.g. Title "Path\model" -> bad escape \m)
        if re.search(rf"<{tag_open}>.*?</{tag_close}>", out, flags=re.DOTALL):
            _repl = f"<{tag_open}>{esc}</{tag_close}>"
            out = re.sub(rf"<{tag_open}>.*?</{tag_close}>", lambda m: _repl, out, flags=re.DOTALL)
            return True
        # Case 3: v3.1 - tag entirely absent -> INSERT, do NOT silently skip (Rule 6)
        # Insert right after <cp:coreProperties ...> opening tag
        m = re.search(r"<cp:coreProperties[^>]*>", out)
        if m:
            out = out[:m.end()] + f"<{tag_open}>{esc}</{tag_close}>" + out[m.end():]
            print(f"  v3.1 metadata: inserted missing <{tag_open}>")
            return True
        print(f"  FAIL: cannot insert missing <{tag_open}> - no coreProperties root")
        return False
    if "__TITLE__" in meta:
        set_or_insert("dc:title", "dc:title", meta["__TITLE__"])
    if "__AUTHOR__" in meta:
        set_or_insert("dc:creator", "dc:creator", meta["__AUTHOR__"])
    if "__LASTMODIFIEDBY__" in meta:
        set_or_insert("cp:lastModifiedBy", "cp:lastModifiedBy", meta["__LASTMODIFIEDBY__"])
    # If any metadata changed, also update Modified date (NEVER AGAIN Rule 6)
    if meta:
        _mod_repl = f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        out = re.sub(r"<dcterms:modified[^>]*>.*?</dcterms:modified>", lambda m: _mod_repl, out, flags=re.DOTALL)
    return out

def count_pat(s, pat): return s.count(pat)

def main():
    if len(sys.argv) != 4:
        print("Usage: python3 safe_docx_edit_v3_2.py input.docx output.docx replacements.txt")
        sys.exit(1)
    inp, outp, rep_path = sys.argv[1], sys.argv[2], sys.argv[3]
    reps, meta = load_replacements(rep_path)
    print(f"Loaded {len(reps)} text replacements, {len(meta)} metadata replacements")
    for o,n in reps: print(f"  {repr(o[:70])} -> {repr(n[:70])}")
    for k,v in meta.items(): print(f"  {k} -> {repr(v[:70])}")

    with zipfile.ZipFile(inp, "r") as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}
    orig_data = dict(data)
    log = []
    # 1. document.xml via pure string surgery - v3.2: direct applied counter
    applied_counts = {}
    if reps:
        doc_str = data["word/document.xml"].decode()
        new_doc = replace_in_document_xml(doc_str, reps, log, applied_counts)
        data["word/document.xml"] = new_doc.encode()
        print(f"Applied counts (direct, not log-parsed): {applied_counts}")
        print(f"Edited word/document.xml via string surgery")
    # 2. core.xml via string replacement
    if meta:
        core_str = data["docProps/core.xml"].decode()
        new_core = apply_metadata(core_str, meta)
        data["docProps/core.xml"] = new_core.encode()
        print(f"Edited docProps/core.xml via string replacement (Title/Author/LastModifiedBy/Modified)")

    # Write output
    with zipfile.ZipFile(outp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, data[n])
    print(f"Wrote {outp}")

    # GATES
    print("\n=== GATES ===")
    ok = True
    def gate(name, cond, detail=""):
        nonlocal ok
        print(f"{'PASS' if cond else 'FAIL'}: {name} {detail}")
        if not cond: ok = False

    with zipfile.ZipFile(inp) as a, zipfile.ZipFile(outp) as b:
        # 1 Image SHA
        for n in a.namelist():
            if n.startswith("word/media/"):
                gate(f"Gate 1 Image SHA {n}", sha(a.read(n))==sha(b.read(n)), f"sha={sha(b.read(n))}")
        # 2 styles
        gate("Gate 2 styles.xml identical", a.read("word/styles.xml")==b.read("word/styles.xml"))
        # 3 header/footer untouched
        for part in ["word/header1.xml","word/footer1.xml","word/header2.xml","word/footer2.xml","word/footnotes.xml","word/endnotes.xml"]:
            if part in a.namelist():
                gate(f"Gate 3 {part} untouched", a.read(part)==b.read(part))
        # 4 namespaces
        doc_o = a.read("word/document.xml").decode(); doc_n = b.read("word/document.xml").decode()
        gate("Gate 4 r:embed unchanged", count_pat(doc_o,"r:embed")==count_pat(doc_n,"r:embed"), f"{count_pat(doc_o,'r:embed')}->{count_pat(doc_n,'r:embed')}")
        gate("Gate 4 wp: unchanged", count_pat(doc_o,"wp:")==count_pat(doc_n,"wp:"), f"{count_pat(doc_o,'wp:')}->{count_pat(doc_n,'wp:')}")
        gate("Gate 4 zero ns0:/ns1: prefixes", "ns0:" not in doc_n and "ns1:" not in doc_n and "ns2:" not in doc_n, "no ns prefixes")
        # 5 xml:space
        gate("Gate 5 xml:space unchanged", count_pat(doc_o,"xml:space")==count_pat(doc_n,"xml:space"), f"{count_pat(doc_o,'xml:space')}->{count_pat(doc_n,'xml:space')}")
        # 6 v3.2 ENFORCED via DIRECT counter (not log-text 60-char prefix match)
        expected_delta = sum((n.count("A*") - o.count("A*")) * applied_counts.get(o, 0) for o, n in reps)
        observed_delta = count_pat(doc_n,"A*") - count_pat(doc_o,"A*")
        gate("Gate 6 A* delta enforced", observed_delta == expected_delta, f"observed {observed_delta} expected {expected_delta} (applied={applied_counts})")
        # 8 v3.2 VERIFIED via ESCAPED form (doc stores &amp; &lt; etc., raw NEW would false-FAIL on & / <)
        for _o, _n in reps:
            _n_esc = xml_escape(_n)  # v3.2: count escaped, not raw
            before_n = count_pat(doc_o, _n_esc)
            after_n = count_pat(doc_n, _n_esc)
            applied = applied_counts.get(_o, 0)
            if _n == _o:
                gate(f"Gate 8 no-duplication for {repr(_n[:40])}", after_n == before_n, f"{before_n}->{after_n} applied={applied}")
            else:
                gate(f"Gate 8 no-duplication for {repr(_n[:40])}", after_n == before_n + applied, f"before {before_n} after {after_n} applied {applied} (escaped form)")
        # 7 only document+core changed (now Gate 7, plus new Gate 9 below)
        changed = [n for n in a.namelist() if a.read(n)!=b.read(n)]
        print(f"INFO Gate 7 Changed parts: {changed}")
        gate("Gate 7 only document.xml + core.xml changed", set(changed) <= {"word/document.xml","docProps/core.xml"}, f"{changed}")
        # 9 v3.1: header/footer + namespace already checked above, but also ensure no <w:tab/> corruption
        # Check that no </w:r> inside <w:t> (phantom match symptom)
        # v3.1 fix: Gate 9 must use STRICT pattern, not loose <w:t[^>]*> (loose pattern itself matches <w:tab/> and false-fails)
        _strict_inners = [m.group(2) for m in re.finditer(r"<w:t(\s[^>]*)?>(.*?)</w:t>", doc_n, flags=re.DOTALL)]
        gate("Gate 9 no phantom <w:tab/> corruption (no </w:r> inside <w:t>)", not any("</w:r>" in _inner or "<w:" in _inner for _inner in _strict_inners), f"checked {len(_strict_inners)} w:t")
    print("\nEdit log:")
    for l in log: print(l)
    if not log and reps:
        print("  (no replacements found - check OLD is exact)")
    if not ok:
        print("\nFAIL: DISCARD OUTPUT. Do NOT send it.")
    else:
        print("\nALL GATES PASS. Safe to send after Word visual check.")

if __name__ == "__main__":
    main()
