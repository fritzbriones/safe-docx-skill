# Changelog

## v3.4 - 2026-11-16
Fixes found in review (all covered by `test_regression.py`, now 20 tests):
- **Entity corruption fixed.** v3.3 decoded only `&amp; &lt; &gt;`, so an edited run
  containing `&quot;`, `&apos;` or `&#8217;` was re-escaped to `&amp;quot;` etc. - visible
  garbage that passed every gate. All XML entities are now decoded; new Gate 12 rejects
  double-escaped entities and malformed XML.
- **All occurrences replaced.** v3.3 silently replaced only the first match per paragraph.
  Now every match is replaced, left to right, without re-scanning inserted text.
- **Unmatched pairs fail** (new Gate 11). v3.3 exited 0 if any pair matched.
- **Failed output is deleted**; any crash exits 1.
- **Output == input refused**, so the master cannot be overwritten.
- Self-closing `<w:p .../>` no longer merges with the next paragraph.
- Paragraphs holding text boxes/shapes are skipped with a warning (unsafe to regex-edit).
- Missing `core.xml` / `document.xml`, empty OLD, lines without TAB: clear error instead of crash.
- `__TITLE__` requires `__AUTHOR__` and `__LASTMODIFIEDBY__` (Rule 7 now enforced).
- Gates 6 and 8 count decoded paragraph text (catches literals split across runs).
- Gate 6 accepts extra protected literals via `__PROTECT__`.
- Warning when edited text gains leading/trailing spaces in a run without `xml:space="preserve"`.
- Cosmetic: literal `\n` in final messages fixed.
- Editor renamed `safe_docx_edit.py` (version is inside the file) so instructions don't break on upgrades.

Added the previously missing scripts:
- `safe_docx_read.py` - health check, outline, table widths, header fields, counts, damage warnings.
- `safe_pdf_read.py` - read-only text with ToUnicode CMap support (Word/LibreOffice PDFs).

Docs: SKILL.md now has `name`/`description` frontmatter, no project-specific numbers,
and does not assume uploads are text-only.

## v3.3 - 2026-11-16
- Full SHA-256 image digests (were truncated).
- Header/footer/notes parts discovered from the archive (catches header3.xml).
- Gate 7 also rejects added/removed parts.
- Exit 0 only when all gates pass; regression tests added.

## v3.2 - 2026-11-14
- Metadata `re.sub` uses a lambda (a backslash in a title crashed it).
- Gates count applications with a direct counter, not log-text prefix matching.
- Gate 8 counts the escaped form (`&amp;`, `&lt;`).

## v3.1 - 2026-11-14
- `<w:t` regex requires whitespace or `>` (the loose form matched `<w:tab/>` and swallowed XML).
- Cross-run stale-index crash fixed.
- Literal-`A*` delta enforced; no-duplication gate actually checked.
- Missing metadata tags inserted instead of silently skipped.

## v3 - 2026-11-14
- ElementTree removed. It renamed namespaces (`r:embed` -> `ns1:`) and rewrote headers
  even with zero edits. Pure string surgery on `<w:t>` only.

## v2
- Gates added for `xml:space` (ElementTree had dropped all of them) and metadata completeness.
