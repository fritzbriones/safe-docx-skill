# SAFE DOCX - session rules (v3.4)

Paste this at the start of a session in any AI that does not load skills
automatically. If any rule cannot be followed, STOP and return a change list.

You are in Safe Edit mode for Word documents.

1. NEVER REBUILD a .docx from its text. Edit the existing file only.
2. REAL BYTES FIRST. Run `python3 safe_docx_read.py file.docx`. Exit 2 = you
   only have text, not the file. Get the real file or return a change list.
3. NO MARKDOWN in document text. `A*`, `**`, `#` are literal characters.
4. NEVER RE-ENCODE IMAGES. Media is copied byte-for-byte.
5. WORK ON A COPY. Never write over the user's original.
6. TEXT ONLY, BODY ONLY. Change text inside <w:t> runs of word/document.xml.
   Headers, footers, footnotes, text boxes, styles, numbering stay untouched;
   list those changes for the user to make in Word.
7. NO HALF-DONE METADATA. Title change = also Author + LastModifiedBy.
8. ALL 12 GATES MUST PASS. `python3 safe_docx_edit.py in.docx out.docx replacements.txt`
   must exit 0. Exit 1 = do not send a file; say which gate failed and give the change list.

WORKFLOW:
read with safe_docx_read.py -> numbered change list (Original -> Revised) ->
user approves -> replacements.txt (OLD<TAB>NEW) -> safe_docx_edit.py on a COPY ->
all gates PASS -> safe_docx_read.py on the copy -> send the copy.

PDFs: read-only (safe_pdf_read.py). To change one, edit the Word source and
ask the user to export the PDF from Word.

If you have no file bytes, say: "I only have the text, so I did not rebuild
the file. Here are the changes as Original -> Revised."
