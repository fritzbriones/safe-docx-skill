---
name: safe-docx
description: Edit or proofread an existing Word .docx without rebuilding it, so styles, headings, tables, images, headers and literal symbols like A* survive. Use whenever the user asks to revise, correct, update or check text in a .docx they provided (manuscripts, reports, contracts), or asks to change a PDF that has a Word source. Not for creating new documents from scratch.
---

# Safe DOCX - edit Word files without rebuilding them

A .docx is a ZIP of XML parts. Formatting lives in those parts (styles.xml,
tblGrid, media, headers), not in the text. Rebuilding a document from its
text destroys them: headings become Normal, tables get equal columns,
images are re-encoded, `A*` loses its asterisk. This skill edits the text
inside the existing file and proves nothing else changed.

All scripts are in this skill's folder and need only Python 3 (stdlib).
Use them even if python-docx or LibreOffice is available - the rules
below apply to every agent.

## Golden rules

1. **Never rebuild a DOCX** from extracted text, and never "save as new" with a library that re-writes styles.
2. **Never re-encode images.** Media bytes are copied verbatim.
3. **No Markdown in document text.** `*`, `**`, `#`, `---` are literal characters in Word. Keep `A*` as `A*`.
4. **Work on a copy.** The user's original is never overwritten.
5. **Edit text only** inside `<w:t>` runs of `word/document.xml`. Headers, footers, footnotes, styles, numbering, theme and media stay untouched.
6. **Verify every edit.** The editor's 12 gates must all pass (exit code 0). Any failure = discard the output.
7. **No real bytes, no file.** If you only have extracted text, return a change list instead.

## Step 1 - make sure you have the real file

```bash
python3 safe_docx_read.py input.docx
```

- Exit 0 = real DOCX bytes. Continue.
- Exit 2 = not a ZIP: you only have a text extract (some platforms give
  uploads as text only). Get the actual file from wherever the user keeps it
  (a connected drive, a repository, a re-upload). If you cannot, go to
  "No real bytes" below.

The report shows the styles in use, the heading outline, table column
widths, header/footer fields, image hashes and counts of protected literals
(`A*` by default; add more with `--protect "p < .05"`). If it ends with
damage warnings (tiny styles.xml, no theme, no headings), tell the user the
file may already be a damaged rebuild and should not be used as the master.

Use `--full-text` to see every paragraph with its style.

## Step 2 - propose a change list

Show the user each change as `Original -> Revised` (numbered), copying
the original text exactly from the document. Wait for approval unless the
user already told you to apply changes directly.

## Step 3 - apply to a copy

Write `replacements.txt`: one change per line, `OLD<TAB>NEW`, exact text.

```
57 A and 184 A journals	57 A* and 184 A* journals
recieved	received
__PROTECT__	p < .05
```

```bash
python3 safe_docx_edit.py input.docx input_edited.docx replacements.txt
```

How matching works:
- Exact and case-sensitive. Every occurrence in every body paragraph is replaced.
- OLD may span several runs (Word splits runs at formatting and spell-check
  boundaries). NEW takes the formatting of the first run.
- Matches never cross paragraphs. Make one pair per paragraph.
- Text in headers, footers, footnotes, endnotes, comments and text boxes nested
  inside paragraphs is out of scope; list those changes for the user to make in Word.
- A pair that matches nothing fails the run (usually a typo or a run of
  hidden characters). Copy the OLD text again from `--full-text` output.

Optional metadata keys: `__TITLE__`, `__AUTHOR__`, `__LASTMODIFIEDBY__`.
Setting a title requires all three; the modified date is updated automatically.

## Step 4 - check the result

The editor prints PASS/FAIL for each gate and exits 0 only if all pass. On
failure it deletes the output. Then:
- Exit 0: run `safe_docx_read.py` on the copy, confirm the outline and counts
  match the original, send the copy, and suggest a quick look in Word.
- Exit 1: do not send a file. Report which gate failed and give the change list.

Watch for `WARN:` lines - e.g. a run whose new text starts or ends with a
space but lacks `xml:space="preserve"` (Word may drop that space).

## PDFs

PDFs are read-only here.

```bash
python3 safe_pdf_read.py input.pdf
```

Best-effort text extraction (handles normal Word/LibreOffice PDFs; no OCR).
Exit 3 means the text is garbled or missing - do not trust it.

To change a PDF: ask for the Word source, edit that with this skill, and ask
the user to export the PDF from Word. Never regenerate a PDF yourself as a
replacement. Without a Word source, return a change list.

## No real bytes

Say so plainly and return a change list:

> I only have the text of this file, not the file itself, so I did not
> rebuild it. Here are the changes as Original -> Revised; your original
> file stays the source of truth for formatting.

## The 12 gates

1 images identical (SHA-256) · 2 styles.xml identical · 3 all headers/footers/notes identical ·
4 namespaces intact · 5 xml:space count unchanged · 6 protected literals change only as listed ·
7 only document.xml (+core.xml) changed, no parts added/removed · 8 no duplicated insertions ·
9 no markup inside text runs · 10 package verifier · 11 every pair matched ·
12 XML well-formed, no double-escaped entities

The guarantee is document-part integrity: every part except the edited ones
is byte-identical. The ZIP container itself is rewritten.

## Files

- `safe_docx_read.py` - read-only health check and outline
- `safe_docx_edit.py` - the editor with the 12 gates
- `safe_pdf_read.py` - read-only PDF text
- `test_regression.py` - 20 tests; run after changing any script
- `NEVER_AGAIN_COMMAND.md` - short rules to paste into agents without skill support
