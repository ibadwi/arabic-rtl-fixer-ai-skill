# Arabic Word documents and PDFs

## DOCX

Use native paragraph bidi properties (`w:bidi`) and right/center/justified alignment as intended. Set section direction (`w:rtlGutter`) only when appropriate to page layout, and set table visual direction (`w:bidiVisual`) only when the table should read from right to left. Apply Arabic language metadata (`w:lang` with `w:bidi`) and a complex-script font (`w:rFonts/@w:cs`) when a typeface is specified.

Use native Word lists, headings, fields, tables, footnotes, and page numbering. Keep URLs and long English passages in LTR paragraphs or cells. Render every page and inspect headers, footers, tables, list indentation, line breaks, and pagination in addition to Arabic order.

## PDF

Create PDF from a correctly authored RTL source; do not construct visual Arabic by reversing strings. Embed or subset fonts only when licensing and the PDF engine permit it, then inspect font information. A searchable PDF must preserve sensible text extraction order; test selection/copying and extraction in addition to appearance.

If editing an existing PDF corrupts Arabic, rebuild from an editable RTL source rather than layering broken text. For scanned PDFs, distinguish OCR text from the visible page and do not claim OCR accuracy without checking representative Arabic passages.
