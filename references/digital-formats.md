# Web, UI, spreadsheets, and images

## Web and UI

Set semantic direction with `dir="rtl"` on the appropriate root or component and the correct `lang` value. Use CSS logical properties such as `margin-inline-start`, `padding-inline-end`, `inset-inline`, and `text-align: start`. Use `<bdi>` for user-generated or direction-unknown inline text and `<bdo>` only for intentional direction overrides. Do not reverse DOM order to imitate RTL.

Test keyboard navigation, icon meaning, form labels, validation, truncation, numbers, dates, currency, phone numbers, URLs, charts, and responsive layouts in a real browser. Localize formatting with locale-aware APIs; do not translate stored identifiers or numeric values.

## Spreadsheets

Set sheet or range RTL through the spreadsheet format, right-align Arabic labels intentionally, and keep formulas and values native. Mirror column order only for Arabic semantic reading order. Inspect charts, filters, frozen panes, print layout, formulas, and exported PDF.

## Images and generated visuals

Prefer adding Arabic text as editable text after image generation. Image models may produce malformed Arabic glyphs; do not accept rendered text without letter-by-letter visual checking. For final raster-only artwork, verify spelling, joins, order, punctuation, contrast, and safe margins at full resolution, and retain an editable source whenever possible.
