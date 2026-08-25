# Arabic PowerPoint

Use the Presentations skill for deck creation and this skill as its Arabic compatibility layer. Treat editable PPTX as the source artifact and PowerPoint desktop as the canonical PPTX viewer.

For every Arabic paragraph, require native DrawingML RTL (`a:pPr/@rtl="1"`), intentional alignment, Arabic run language metadata, and RTL text-body direction (`a:bodyPr/@rtlCol="1"`) where applicable. Use native bullets. For Arabic-first decks, set presentation reading order (`p:presentation/@rtl="1"`); do not force that flag on an English-primary bilingual deck.

Export a raw draft, preserve it, and normalize to a distinct final file:

```bash
python3 "$SKILL_DIR/scripts/arabic_pptx.py" repair draft.pptx final.pptx --lang ar-SA --convert-bullets
python3 "$SKILL_DIR/scripts/arabic_pptx.py" audit final.pptx --strict
```

Add `--presentation-rtl` to repair and `--expect-presentation-rtl` to audit for an Arabic-first deck. Add `--font "Font Name"` only when a typeface was chosen; this populates the complex-script slot but does not embed the font.

A strict audit warning or error blocks delivery. Render every slide for layout QA, but do not use LibreOffice/Codex preview as authoritative proof of Arabic word order. Inspect PowerPoint itself before reporting `PowerPoint visual QA: PASS`.

Set RTL on every table-cell paragraph. Mirror table column order only when its semantic reading order is Arabic. Prefer Arabic chart titles and labels in separate RTL text boxes when native chart bidi cannot be tested.
