# Safe DOCX

**Read Word files properly. Edit them without rebuilding. Prove nothing else changed.**

Python stdlib only — no `python-docx`, no PyMuPDF, no Office, no pip installs.

---

## The problem

Ask an AI to "fix" a Word document and there is a good chance it will hand back a corrupted file:

| Damage | Why it happens |
|---|---|
| `A*` (and every `**`) became `A` | Markdown parsing treats `*` as emphasis |
| All headings flattened to plain text | The rebuild never copied paragraph styles |
| Table columns all equal width | The rebuild guessed widths instead of copying `tblGrid` |
| Figures downscaled 3× and re-encoded | The rebuild re-encoded images instead of copying bytes |
| `styles.xml` 349 KB → 891 bytes | The file was regenerated from scratch with a minimal builder |
| Page-number fields gone | Headers/footers never made it into the rebuild |

The root cause is always the same two-step:

1. The AI only **has** (or only trusts) the *text* of the document — extracted, style-free, image-free.
2. It then **rebuilds** the DOCX from that text. Formatting lives in the bytes it never had, so formatting dies.

This skill breaks that loop: treat the DOCX as what it is — a ZIP of XML parts — read the real bytes, edit only the text runs, verify everything else stayed byte-identical.

## Who this is for

**Primary audience: an AI with stdlib-only tools** — Python 3 + Node.js + curl, no pip installs, no Office/LibreOffice, uploads that arrive as text-only stubs rather than file bytes. Everything here works with `zipfile`, `re`, `hashlib`, `zlib` alone.

**If your AI has full capability** (python-docx, PyMuPDF, LibreOffice), you don't need the stdlib tricks — but the golden rules apply to you unchanged:

> **Never rebuild a DOCX. Never re-encode images. Never use Markdown for DOCX text. Edit in place or return a change list. Always verify (image SHA, styles, literal-asterisk counts) after every edit.**

Rebuilding "because you can" causes the same corruption as rebuilding because you must.

## What's in this repo

| File | What it does |
|---|---|
| `safe_docx_read.py` | Health check: parts + image SHAs, paragraph styles, real table column widths, `A*` / highlight / tracked-change counts, styles.xml size. Read-only. |
| `safe_docx_edit_v3_2.py` | The editor: ZIP copy + pure string surgery on `<w:t>` text runs. Cross-run safe, no XML tree re-serialization ever. Runs 9 verification gates on its own output. |
| `safe_pdf_read.py` | Best-effort text extraction from PDFs (re + zlib). Read-only — PDFs are never rebuilt. |
| `SKILL.md` | The full skill: rules, byte sources, workflows, legacy checks, and the version history (v3 → v3.2) with every bug that was found and fixed. |
| `NEVER_AGAIN_COMMAND.md` | The 8 rules + 9 gates, designed to be pasted at the start of every editing session. |

## Quick start

```bash
# 1. Read the file properly first
python3 safe_docx_read.py manuscript.docx

# 2. Write your changes as exact old<TAB>new pairs
#    (metadata keys: __TITLE__, __AUTHOR__, __LASTMODIFIEDBY__)
cat > replacements.txt <<'EOF'
57 A and 184 A	57 A* and 184 A*
__TITLE__	My Paper Title
EOF

# 3. Edit a copy — the master is never touched
python3 safe_docx_edit_v3_2.py manuscript.docx manuscript_edited.docx replacements.txt

# 4. The script runs all 9 gates itself and tells you PASS/FAIL per gate.
#    FAIL = discard the output. Send a change list instead.
```

If you cannot get the real file bytes (only a text stub), **stop**: return a text change list (`Original → Revised`, numbered), and say so. A change list applied by the user in Word is always safe; a rebuild never is.

## The 9 gates

Every edit is verified after writing:

1. Image SHA identical (no re-encoding)
2. `styles.xml` byte-identical
3. Headers/footers byte-identical (untouched by default)
4. Namespace prefixes intact (`r:embed`, `wp:`, no `ns0:`/`ns1:` renames)
5. `xml:space` count unchanged
6. Literal-`A*` delta enforced (observed == expected from the change list)
7. Only `document.xml` + `core.xml` changed
8. No duplication (NEW count == before + applied; escapes counted in stored form)
9. No phantom `<w:tab/>` corruption (strict regex, no `</w:r>` inside `<w:t>`)

Any FAIL → discard the output. Do not send the file.

## Why the weird engineering choices are load-bearing

- **No ElementTree.** XML tree re-serialization silently renames namespaces (`r:embed` → `ns1:`), rewrites headers, and drops `xml:space`. v2 of the editor did exactly that while *passing* its own checks. v3+ splices raw bytes inside `<w:t>...</w:t>` only.
- **Replace once per paragraph, no re-scan loop.** Insertions usually make NEW contain OLD; rescanning would duplicate the insert N times.
- **Count gates in escaped form.** The document stores `&amp;`/`&lt;`; counting raw text false-fails good files.
- **Lambda, not f-string, in `re.sub` replacements.** A title like `Path\model results` crashes a template (bad escape `\m`).
- **Strict `<w:t` regex with required whitespace.** The loose version matches `<w:tab/>` and swallows XML.

Each of these was a real bug, found by testing on a real journal manuscript and by adversarial synthetic files — the fix history is documented in `SKILL.md`.

## Verified

v3.2 was tested against a real accepted journal manuscript (byte-level) plus adversarial synthetic cases: the `<w:tab/>` trap, backslash-in-title crash, true cross-run replacement, `&`/`<` in replacement text, two replacements sharing a 60-char prefix, and literal-`A*` delta enforcement. All 9 gates and all 6 adversarial tests pass.

## Repository layout

```
safe-docx/
├── README.md
├── SKILL.md
├── NEVER_AGAIN_COMMAND.md
├── safe_docx_read.py
├── safe_docx_edit_v3_2.py
└── safe_pdf_read.py
```

## License

MIT — use it, fork it, teach your AI to stop corrupting documents.
