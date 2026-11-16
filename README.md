# Safe DOCX

**Edit Word files without rebuilding them - and prove nothing else changed.**

Python 3 standard library only. No python-docx, no Office, no pip installs.

## The problem

Ask an AI to fix a Word document and it often hands back a damaged one:

| Damage | Cause |
|---|---|
| `A*` became `A` | Markdown treated `*` as emphasis |
| Headings turned into plain text | Rebuilt from text; styles never copied |
| Table columns all equal | Widths guessed instead of copied |
| Figures downscaled, re-encoded | Images decoded and re-saved |
| `styles.xml` 349 KB -> 891 bytes | File regenerated from scratch |
| Page numbers gone | Headers not carried over |

Root cause: the AI works from the document's *text*, then *rebuilds* the file.
This skill edits the text inside the original file and verifies every other
part stayed byte-identical.

## Install

| Agent | How |
|---|---|
| Claude (claude.ai, Claude Code) | Upload the `safe-docx` folder (or this zip) as a skill. `SKILL.md` has the required frontmatter. |
| Other agents with skill/instruction folders (Codex, Cursor, etc.) | Put the folder in the project and point the agent's instruction file (e.g. `AGENTS.md`) at `SKILL.md`. |
| Chat AIs without skills (ChatGPT, Gemini, ...) | Add the scripts to the conversation/project files and paste `NEVER_AGAIN_COMMAND.md` at the start of the session. |

The agent needs to be able to run Python 3 on the real .docx file. If it only
sees extracted text, the skill tells it to return a change list instead of a file.

## Quick start

```bash
python3 safe_docx_read.py paper.docx                 # 1. check it's a real, healthy DOCX

printf '57 A and 184 A\t57 A* and 184 A*\n' > replacements.txt   # 2. OLD<TAB>NEW

python3 safe_docx_edit.py paper.docx paper_edited.docx replacements.txt   # 3. edit a copy
echo $?                                               # 0 = all 12 gates passed; 1 = discarded
```

## Files

| File | Purpose |
|---|---|
| `SKILL.md` | Instructions for the AI |
| `NEVER_AGAIN_COMMAND.md` | Short version to paste into agents without skill support |
| `safe_docx_read.py` | Read-only health check: styles, outline, table widths, header fields, image hashes, counts, damage warnings |
| `safe_docx_edit.py` | The editor: text-only surgery in `<w:t>` runs + 12 verification gates |
| `safe_pdf_read.py` | Read-only PDF text extraction (no OCR) |
| `test_regression.py` | 20 regression tests - run after any change |
| `CHANGELOG.md` | Every bug found and fixed, v2 -> v3.4 |

## The 12 gates

1. Images identical (SHA-256)
2. `styles.xml` identical
3. All headers, footers, footnotes, endnotes identical
4. Namespace prefixes intact
5. `xml:space` count unchanged
6. Protected literals (`A*`, plus any `__PROTECT__`) change only as the change list says
7. Only `document.xml` (+ `core.xml`) changed; no parts added or removed
8. No duplicated insertions
9. No markup inside text runs
10. Shared package-integrity verifier
11. Every replacement pair matched
12. Edited XML well-formed, no double-escaped entities

Exit 0 only if all pass; otherwise the output is deleted.
Guarantee scope: document-part integrity. The ZIP container is rewritten.

## Limits

- Edits body text only. Headers, footers, footnotes, comments and text boxes are
  reported for the user to change in Word.
- Matching is exact and within one paragraph.
- New text takes the formatting of the first run it replaces.
- PDF reading is best-effort; scanned PDFs need OCR.

## License

MIT - see `LICENSE`.
