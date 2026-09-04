# Arabic RTL Fixer AI Skill

[![Test](https://github.com/ibadwi/arabic-rtl-fixer-ai-skill/actions/workflows/test.yml/badge.svg)](https://github.com/ibadwi/arabic-rtl-fixer-ai-skill/actions/workflows/test.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

An AI skill for creating, repairing, and validating Arabic and Arabic-English content with native right-to-left direction, clean logical-order Unicode, correct BiDi behavior, Arabic language metadata, editable lists, and Arabic-capable typography.

## Install

```bash
npx skills add ibadwi/arabic-rtl-fixer-ai-skill -s arabic-rtl-fixer-ai-skill -g -y
```

Start a new task after installation, then ask the agent to create, fix, or validate an Arabic DOCX, PPTX, PDF, web page, spreadsheet, or image with text.

Requirements: Node.js with `npx` for installation and Python 3.9+ for the bundled OOXML tools. The tools use only the Python standard library.

## Support matrix

| Format | Skill guidance | Structural audit | Automated repair | Required final QA |
|---|---:|---:|---:|---|
| PowerPoint (PPTX) | Yes | Yes | Yes | Render and inspect in PowerPoint when available |
| Word (DOCX) | Yes | Yes | Yes | Render every page and inspect in Word when available |
| PDF | Yes | Partial/manual | No | Render, inspect, and test text extraction |
| Web/UI | Yes | Manual | No | Test direction, interaction, and responsiveness in a browser |
| Spreadsheets | Yes | Manual | No | Inspect sheet direction, formulas, charts, and export |
| Images with Arabic text | Yes | Manual | No | Check every glyph, join, word, and safe margin |

Automated structural repair does not prove visual correctness or compatibility with every application.

## PowerPoint audit and repair

```bash
python3 scripts/arabic_pptx.py audit presentation.pptx --strict

python3 scripts/arabic_pptx.py repair \
  presentation.pptx \
  presentation-fixed.pptx \
  --lang ar-SA \
  --presentation-rtl \
  --convert-bullets
```

The PPTX tool can set native paragraph RTL, text-body RTL, Arabic run language, the complex-script font slot, presentation reading order, and native bullets. It removes accidental bidi controls but preserves ZWJ and ZWNJ.

## Word audit and repair

```bash
python3 scripts/arabic_docx.py audit document.docx --strict

python3 scripts/arabic_docx.py repair \
  document.docx \
  document-fixed.docx \
  --lang ar-SA \
  --font "Arial"
```

The DOCX tool audits and repairs Arabic paragraphs and runs across Word XML parts, including headers, footers, footnotes, and endnotes. It sets native `w:bidi`, Arabic language metadata, run RTL, and the complex-script font slot while preserving unrelated package parts.

Both repair commands require a distinct output path and refuse to replace an existing output unless `--force` is passed.

## Why native RTL matters

Right alignment is only visual placement; it does not establish RTL reading order. Keep source text in normal logical Unicode order:

```text
ارتفع الأداء بنسبة 25% في Q4 2026.
يعتمد النظام على GPS وTCP/IP.
البريد: support@example.com
```

Apply RTL through paragraph, container, table, or document properties. Do not reverse Arabic strings, pre-shape letters, add manual spaces, or wrap ordinary text in invisible BiDi controls.

## Quality report

The skill separates four kinds of evidence:

```text
Arabic content QA: PASS/FAIL
RTL and Unicode structural QA: PASS/FAIL
Layout QA: PASS/FAIL
Target-application QA: PASS/NOT RUN
Font portability: verified / installed-font dependent / PDF supplied
```

## Development

Run the complete standard-library test suite:

```bash
python3 -m unittest discover -s tests -v
```

The tests generate minimal DOCX and PPTX fixtures, exercise broken and mixed-language content, confirm safe output behavior, check repeatability, and ensure unrelated binary assets survive repair. See [examples](examples/README.md) and [CONTRIBUTING.md](CONTRIBUTING.md).

## Credits

Parts of the description, structure, and PPTX helper were inspired by Sultan Alsafran's MIT-licensed Arabic Presentations skill. See [THIRD_PARTY_LICENSE.md](THIRD_PARTY_LICENSE.md).

## License

MIT. See [LICENSE](LICENSE).

---

## مهارة إصلاح النصوص العربية

مهارة لإنشاء المحتوى العربي والعربي‑الإنجليزي وإصلاحه والتحقق منه باستخدام اتجاه RTL أصلي، وترتيب Unicode منطقي، ومعالجة صحيحة للنص ثنائي الاتجاه، وبيانات اللغة العربية، والخطوط الداعمة للعربية.

### التثبيت

```bash
npx skills add ibadwi/arabic-rtl-fixer-ai-skill -s arabic-rtl-fixer-ai-skill -g -y
```

تتضمن المهارة أداتين تعملان محليًا دون مكتبات Python خارجية:

- `arabic_pptx.py` لتدقيق ملفات PowerPoint وإصلاح بنيتها العربية.
- `arabic_docx.py` لتدقيق ملفات Word وإصلاح اتجاه الفقرات وبيانات اللغة والخط المركّب.

يُنشئ أمر الإصلاح ملفًا جديدًا ويحافظ على الأصل. ولا يعني نجاح الفحص البنيوي أن العرض المرئي صحيح؛ يجب عرض جميع الصفحات أو الشرائح وفحصها، ثم فتح الناتج في Word أو PowerPoint عندما يكون ذلك متاحًا.

### الاختبارات

```bash
python3 -m unittest discover -s tests -v
```

تختبر الحزمة النص العربي المختلط بالأرقام والاختصارات اللاتينية ومحارف BiDi الخفية، كما تتحقق من عدم تلف الأجزاء الأخرى داخل ملفات OOXML ومن ثبات الإصلاح عند تكراره.
