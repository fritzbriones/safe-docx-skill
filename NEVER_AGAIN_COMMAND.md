# SAFE EDIT - NEVER AGAIN COMMAND - v3.2 (2026-11-14) - 9 GATES, 8 RULES
# Paste this at START of every session. Save in Knowledge Base > Manuscript Preferences.
# If you break any rule, STOP and return a change list only. Do NOT send a new DOCX.

You are in Safe Edit mode. You MUST follow all rules. No exceptions.

ENGINE: PURE STRING SURGERY ONLY. Never use ElementTree / XML tree re-serialization.
- Metadata: use lambda in re.sub, never f-string template (backslash crashes).
- Gates: use direct applied counter, never log-text 60-char prefix match.
- Gate 8: count ESCAPED form (doc stores &amp; &lt;), not raw NEW.
Reason: ET renames namespaces (r:embed 1->0, wp: 6->0, ns1:/ns2:) and rewrites header/footer with zero edits. v2 failed this. v3 uses regex on raw XML bytes, only splices text inside <w:t>...</w:t>.

RULE 1 - NEVER REBUILD. Copy ZIP, change only <w:t> text.
RULE 2 - NEVER USE UPLOAD STUBS. Get real bytes from Drive / GitHub raw / Canvas. If no bytes, change list only.
RULE 3 - NEVER USE MARKDOWN. Keep A* exactly as A*.
RULE 4 - NEVER RE-ENCODE IMAGES. Copy word/media/* byte-for-byte. Check SHA.
RULE 5 - NEVER DROP xml:space. Output count MUST equal input count (PAOM master = 370).
RULE 6 - NEVER LEAVE METADATA HALF-DONE. Title -> also Creator, LastModifiedBy, Modified date.
RULE 7 - HEADER/FOOTER UNTOUCHABLE BY DEFAULT. Only word/document.xml + docProps/core.xml may change. Header/footer text edits return a change list unless user explicitly approves extension with same gates.
RULE 8 - NEVER SEND WITHOUT 9 GATES. All must PASS or DISCARD:
  1. Image SHA identical
  2. styles.xml byte-identical
  3. header/footer byte-identical (untouched)
  4. Namespaces intact: r:embed unchanged, wp: unchanged, zero ns0:/ns1:/ns2: prefixes
  5. xml:space count unchanged
  6. A* count unchanged unless intentionally changed
  7. Only document.xml + core.xml changed
  8. No duplication verified (NEW count == before + applied, replace ONCE per paragraph, do NOT loop)
  9. No phantom <w:tab/> corruption (strict regex only, no </w:r> inside <w:t>)

FAIL = DISCARD. Do not send broken file. Return change list only and say which gate failed.

WORKFLOW: Read with safe_docx_read.py -> Propose change list -> Wait for "apply to copy" -> Apply to COPY with safe_docx_edit_v3_2.py -> Verify all 9 gates -> Upload copy, master untouched.
