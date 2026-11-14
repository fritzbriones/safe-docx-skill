# How to Use This Repo

This repo contains the complete **Safe DOCX** skill — read Word files properly, edit them without rebuilding, prove nothing else changed.

## What's inside

```
safe-docx/
├── README.md              Start here — the problem, quick start, the 9 gates
├── SKILL.md               The full skill instructions for the AI
├── NEVER_AGAIN_COMMAND.md Paste this at the start of every editing session
├── safe_docx_read.py      Health check reader (read-only)
├── safe_docx_edit_v3_2.py The editor (ZIP copy + <w:t> string surgery + 9 gates)
├── safe_pdf_read.py       PDF text extractor (read-only)
├── HOW-TO-USE.md          This file
├── LICENSE                MIT
└── .gitignore
```

## How to use the scripts

```bash
# 1. Read the file properly first
python3 safe_docx_read.py manuscript.docx

# 2. Write your changes as exact old<TAB>new pairs in replacements.txt
#    (metadata keys: __TITLE__, __AUTHOR__, __LASTMODIFIEDBY__)

# 3. Edit a copy — the master is never touched
python3 safe_docx_edit_v3_2.py manuscript.docx manuscript_edited.docx replacements.txt

# 4. The script runs all 9 gates itself: any FAIL = discard the output.
```

If you cannot get the real file bytes (only a text stub), **stop**: return a numbered text change list (`Original → Revised`) instead. A change list applied by the user in Word is always safe; a rebuild never is.

## What specs of AI this is for

**Primary target — an AI with these specs:**
- Python 3 (stdlib only) + Node.js + curl
- NO pip installs, NO python-docx, NO PyMuPDF, NO Pillow, NO pandas
- NO Office / LibreOffice / pandoc
- File uploads arrive as TEXT-ONLY STUBS (extracted text + OCR), never real file bytes
- Gets real bytes instead from: Google Drive connector, GitHub raw links, or Canvas

**If the AI has full capacity** (python-docx, PyMuPDF, LibreOffice available): the stdlib-only tricks aren't needed, but the **golden rules apply to every AI without exception**:
1. Never rebuild a DOCX from scratch
2. Never re-encode images (copy bytes verbatim)
3. Never use Markdown for DOCX text (`*` is content: `A*` must survive)
4. Edit in place (`<w:t>` surgery) or return a change list
5. Always verify after every edit: image SHA, styles.xml, literal-asterisk counts

**To install as a skill for your AI:** copy `SKILL.md` into your AI's skills folder (if it supports skills), and paste `NEVER_AGAIN_COMMAND.md` at the start of any editing session.

## Verification status

Tested against a real accepted journal manuscript (byte-level) plus 6 adversarial synthetic cases. All 9 gates and all adversarial tests pass (2026-11-14). Scripts are Python 3 stdlib only — run anywhere Python runs.
