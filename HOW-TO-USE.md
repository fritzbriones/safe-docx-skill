# How to Use This Repo

This repo contains the complete **Safe DOCX** skill.

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

## How to get the files

**Option A — GitHub web upload (easiest, no tools needed):**
1. Unzip this file to a folder.
2. Go to https://github.com/new → name the repo (e.g. `safe-docx`) → Create.
3. Click **"uploading an existing file"** on the empty repo page.
4. Drag ALL the unzipped files in (keep them at the root — no subfolder needed).
5. Commit with a message like "Safe DOCX skill v3.2 — read properly, edit without rebuilding".

**Option B — command line:**
```bash
unzip safe-docx-github.zip -d safe-docx && cd safe-docx
git init && git add -A && git commit -m "Safe DOCX skill v3.2"
git remote add origin https://github.com/<your-username>/safe-docx.git
git branch -M main && git push -u origin main
```

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

**To install as a skill for your AI:** copy `SKILL.md` into your AI's skills folder (if it supports skills, e.g. Vibe), and paste `NEVER_AGAIN_COMMAND.md` at the start of any editing session.

## Verification status

Tested against a real accepted journal manuscript (byte-level) plus 6 adversarial synthetic cases. All 9 gates and all adversarial tests pass (2026-11-14). Scripts are Python 3 stdlib only — run anywhere Python runs.
