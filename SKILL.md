# Safe DOCX / PDF Skill - Read Properly, Edit Without Rebuilding

## Who this is for

**Primary audience: an AI that has ONLY**
- Python 3.12 stdlib, Node.js, curl, built-in web search/browse
- NO pip installs, NO python-docx, NO PyMuPDF, NO Pillow, NO pandas
- NO Office / LibreOffice / pandoc, NO TTF fonts
- File readers that extract TEXT + images (OCR)
- Writable: working folders + Canvas, Read/write: Knowledge Base
- Read-only: user uploads as TEXT-ONLY STUBS - never file bytes

**Full-capability AIs** (python-docx, PyMuPDF, LibreOffice available): you do not
need the stdlib-only tricks, but the GOLDEN RULES apply to you unchanged:
never rebuild a DOCX, never re-encode images, never use Markdown for DOCX text,
edit in place or return a change list, and ALWAYS verify (image SHA, styles,
A* / literal-asterisk counts) after every edit. Rebuilding "because you can"
causes the same corruption as rebuilding because you must.

This skill fixes the root cause of every rebuild mess in PAOM / JIM / BER:
text-only stubs + rebuild-from-scratch = lost headings, lost A*, equal table columns, downscaled figures.

## Golden Rules - NEVER break these

1. **NEVER use upload text stubs to edit.** Stubs have no styles, no table widths, no image bytes, no theme. They are for READING content only.
2. **NEVER rebuild a DOCX from scratch.** Never create a new DOCX by copying text into a new file. That is what destroyed PAOM Clean.
3. **NEVER re-encode images.** Copy image bytes verbatim. Never decode + re-encode JPEG/PNG.
4. **NEVER use Markdown for DOCX work.** Markdown treats `*` as italic. That is why `A*` became `A`. Do not output `*`, `**`, `#`, `---` as formatting.
5. **NEVER rebuild a PDF.** PDFs are READ-ONLY in this skill.
6. **ALWAYS edit in place.** Copy the ZIP, change only `<w:t>` text nodes, keep everything else byte-identical.
7. **ALWAYS verify after edit.** Images SHA must be identical. Styles.xml must be identical. `A*` count must be unchanged unless you intentionally changed it.

## How to get REAL file bytes (not stubs)

Uploads are stubs. Get bytes from a real source instead:

- **Google Drive:** Use the Drive connector to download the file. Do not use the upload stub.
- **GitHub:** Use `curl -L -o file.docx https://raw.githubusercontent.com/<user>/<repo>/<branch>/file.docx`
- **Canvas:** Use the Canvas connector - open the document there, it has real bytes.
- **Gmail:** Do NOT rely on Gmail for bytes - that connector cannot download attachments. Ask user to put file in Drive / GitHub / Canvas.

If you cannot get bytes, STOP. Do text-only work:
Return a change list `Original -> Revised`, NOT a new DOCX. Say: "I did not have file bytes, so I did not rebuild."

## How to READ a DOCX properly (stdlib only)

A DOCX IS a ZIP. Do not treat it as plain text.

```bash
python3 safe_docx_read.py input.docx
python3 safe_docx_read.py input.docx --full-text
```

What it checks:
- Every ZIP part + size. `word/media/image1.png` must be listed with SHA.
- `word/document.xml`: `A*` count, `w:highlight` count, tracked `<w:ins`
- Paragraphs with STYLE: `Title`, `Heading 1`, `Heading 2`, `Normal` - not just text
- Tables with REAL `gridCol` widths (twips). Example good: `1700, 2000, 3560, 2100`. Bad: `2340, 2340, 2340, 2340` = rebuilt, equal columns.
- `word/styles.xml` length. If < 2000 bytes, file was rebuilt from scratch. Good file: ~349,000 bytes for PAOM.

If you see equal columns + tiny styles.xml + no theme, the file is already damaged. Tell the user, do not use it as master.

## How to EDIT a DOCX properly (stdlib only)

Use `safe_docx_edit_v3_2.py`. It does ZIP-copy + `<w:t>` surgery only.

```bash
# 1. Make replacements.txt - OLD<TAB>NEW per line, exact text
# Example replacements.txt:
# A and A journals<TAB>A* and A journals
# 57 A and 184 A<TAB>57 A* and 184 A

# 2. Run edit
python3 safe_docx_edit_v3_2.py input.docx output.docx replacements.txt

# 3. Verify
python3 safe_docx_read.py output.docx
```

What the script guarantees:
- Copies every ZIP entry byte-for-byte, EXCEPT `word/document.xml`, `word/header*.xml`, `word/footer*.xml`, `word/footnotes.xml`, `word/endnotes.xml`
- In those XML files, replaces text ONLY inside `<w:t>` nodes. Run formatting is kept.
- If OLD spans multiple `<w:t>` runs (Word splits runs), it merges safely: NEW goes in first run, other runs cleared, first run formatting kept. No new runs, no new styles.
- After edit, it checks: Images SHA identical = PASS, styles.xml identical = PASS. If FAIL, discard output.

### What NOT to do when editing
- Do NOT use `python-docx` - you do not have it, and it rebuilds styles.
- Do NOT write a new DOCX builder from scratch.
- Do NOT set font, size, bold, spacing manually. Leave all formatting alone.
- Do NOT touch `word/media/`, `word/theme/`, `word/styles.xml`, `word/numbering.xml`.

## How to READ a PDF properly (stdlib only)

```bash
python3 safe_pdf_read.py input.pdf
```

- Uses `re` + `zlib` only to decompress streams and extract `(...) Tj` / `[...] TJ` text.
- READ-ONLY. Prints page estimate + text.
- Limitation: For complex PDFs with custom encodings, text may be garbled. If garbled, do NOT trust it. Ask user for the DOCX source.

## How to EDIT a PDF

**Do not edit PDFs.**

If user asks to change a PDF:
1. Ask for the source DOCX.
2. Edit the DOCX with `safe_docx_edit_v3_2.py`.
3. Tell user: "Export PDF from Word on your computer." You have no Office / LibreOffice / pandoc, so you cannot render a faithful PDF. Your PDF builder with Times core fonts is approximation only.

If there is no DOCX source, return a change list only, NOT a new PDF.

## Verification Checklist (LEGACY v1 - superseded by the 9 gates; kept for quick manual spot-checks)

Before you send any file, check all 5:

1. **Headings:** `Title`, `Heading 1`, `Heading 2` still present? Not all `Normal`?
2. **A*:** Search `A*`. Count must match input unless you intentionally changed it (PAOM master = 7).
3. **Tables:** `gridCol` widths still different, not all equal? Records column still centered if it was centered?
4. **Figure:** Image SHA identical to input? (Example: PAOM Figure 1 is ~1,068 KB, not ~207 KB.) If SHA changed, figure is damaged.
5. **Header:** Page number field still in `word/header1.xml`?

If any check fails, discard output, go back to input, and return a text change list instead.

## Prompt to use when user asks for a revision

Copy this to the user / to yourself at start of every task:

> I am in Safe Edit mode. I will NOT rebuild your DOCX/PDF. I will read the real file bytes via Drive/GitHub/Canvas, not upload stubs. I will return a change list Original -> Revised, and if you approve, I will apply it with safe_docx_edit_v3_2.py using in-place <w:t> surgery only. Images, styles, tables, headings will stay byte-identical. I will verify with safe_docx_read.py before sending.

## Files in this skill

- `safe_docx_read.py` - Read DOCX properly, stdlib only
- `safe_docx_edit_v3_2.py` - Edit DOCX in place, v3.2, stdlib only
- `safe_pdf_read.py` - Read PDF, stdlib only, READ-ONLY
- `SKILL.md` - This file

## Example - PAOM failure that this skill prevents

| Error in Clean | Cause | Skill fix |
|---|---|---|
| `A*` -> `A` (7 places) | Markdown `*` treated as italic | Never use Markdown, replace only in `<w:t>`, verify `A*` count |
| Headings lost, all Normal | Text stub has no styles | Read styles from `word/styles.xml`, never rebuild |
| Tables equal columns 2340 each | Builder guessed widths | Copy `tblGrid` verbatim, never set widths |
| Figure 6000px -> 1500px, 1068KB -> 207KB | Re-encoded image | Copy `word/media/*` byte-for-byte, check SHA |
| styles.xml 349KB -> 891 bytes | Custom builder from scratch | Copy `word/styles.xml` verbatim |
| Header page number lost | Rebuild dropped header | Copy `word/header1.xml` verbatim |

## If bytes are not available

Say this honestly, do not fake it:

> I only have a text stub for this file, not the real bytes, so I cannot safely rebuild it. Here is a text-only change list. Please apply it in Word in your master file. Your original file remains the source of truth for formatting.


## NEVER AGAIN - Added after PAOM Test found 2 bugs (2026-11-14)

The test file passed 5 gates but had 2 hidden bugs. Add these to every check:

**Bug 1 - xml:space stripped (370 -> 0).** The editor used ElementTree and dropped all `xml:space="preserve"` attributes. Paragraph texts still read the same, but the file is not truly byte-identical. Fix: For simple edits, use string replacement on raw XML bytes, not ET. For ET edits, verify xml:space count after. If count changes, DISCARD.

**Bug 2 - Metadata half-done.** Title + Author set, but LastModifiedBy blank, dates still 2013. Fix: If you set Title, you MUST also set Creator, LastModifiedBy, and Modified date (now, UTC, W3CDTF format).

**New gates 6 and 7 (historical: 7 gates at that point; the final v3.2 has 9 gates):**
6. xml:space count unchanged (PAOM = 370)
7. Only document.xml + core.xml changed, all other parts byte-identical

Save NEVER_AGAIN_COMMAND.md in Knowledge Base > Manuscript Preferences and paste it at start of every session.


## v3 - 2026-11-14 - ET REMOVED, STRING SURGERY ONLY

v2 bug confirmed on real PAOM: ET renamed namespaces (r:embed 1->0, wp: 6->0, ns1:/ns2:) and rewrote header1/footer1 even with zero edits requested. Gates passed but file had silent corruption.

Fix: v3 editor `safe_docx_edit_v3.py` uses pure string surgery (regex on <w:t> only). No ET. Header/footer locked by default. Added gates:
- Gate 4 Namespaces intact (r:embed, wp:, zero ns prefixes)
- Gate 3 Header/footer untouched
- Gate 8 No duplication (replace once per paragraph, NEW often contains OLD)

v1 test file (PAOM_Test_SafeEdit.docx) did NOT have namespace bug (r:embed still 1, wp: 6) - only xml:space bug (370->0). Fixed file (PAOM_Final_Manuscript.docx) uses string surgery and fixes both.


## v3.1 - 2026-11-14 - 5 latent bugs fixed by adversarial testing

Tested on real PAOM + synthetic tab file + synthetic missing-metadata file + true cross-run:
1. <w:tab/> trap: <w:t([^>]*)> matched <w:tab/> and swallowed XML. Fixed to <w:t(\s[^>]*)?> requiring whitespace or >.
2. Cross-run crash: stale group index. Fixed, verified true cross-run (NEW in first run, tail preserved).
3. Gate 6 A* was INFO-only. Now ENFORCED: observed delta == expected delta from change list.
4. Gate 8 no-duplication was claimed but never checked. Now VERIFIED: NEW count == before + applied.
5. Silent metadata no-op: if tag missing entirely, script skipped silently. Now INSERTS tag or FAILS loudly.
6. Gate 9 itself had the tab bug (loose regex in checker). Fixed to strict regex.

Validation: Real PAOM all 9 gates PASS. Synthetic tab + intentional 184 A -> 184 A* (Gate 6 delta 2 proven). Synthetic missing lastModifiedBy inserted. True cross-run verified.


## v3.2 - 2026-11-14 - 3 more bugs fixed by stress-testing

1. Backslash crash: metadata used f-string as re.sub replacement template. Title "Path\\model results" with existing <dc:title>Old</dc:title> -> re.error: bad escape \\m, run dies. PAOM masked it (tags self-closing). Fixed with lambda. Verified title with backslash writes correctly.
2. Fragile gate counting: Gates 6/8 counted applications by matching log lines against first 60 chars of OLD. Two replacements sharing 60-char prefix miscount. Fixed with direct applied counter dict passed through replace function.
3. Gate 8 false-FAIL on & / <: counted raw NEW, but doc stores escaped (&amp; &lt;). NEW with & or < would fail on good file. Fixed to count escaped form.

Validation: Real PAOM all 9 gates PASS. Synthetic backslash Case 2 (populated Title). Synthetic shared 60-char prefix (1+1 correctly, not 2+2). Synthetic & / < NEW (Gate 8 passes with escaped counting).
