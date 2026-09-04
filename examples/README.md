# Test examples

The `broken/` directory contains small synthetic DOCX and PPTX samples with missing native RTL metadata and an accidental hidden bidi control. The corresponding files in `expected/` were produced by the bundled repair tools and pass strict structural audit.

The automated test suite generates minimal broken DOCX and PPTX packages containing representative Arabic, Latin acronyms, digits, percentages, and hidden bidi controls. It repairs them, audits the result, verifies repeatability, and confirms that unrelated binary assets remain unchanged.

Run all examples through the tests:

```bash
python3 -m unittest discover -s tests -v
```

Regenerate the downloadable examples:

```bash
python3 scripts/generate_examples.py
```

Representative mixed-direction strings:

```text
ارتفع الأداء بنسبة 25% في Q4 2026.
يعتمد النظام على GPS وTCP/IP.
البريد: support@example.com
```

Do not copy a visually reversed string into a source document. Store Arabic in logical Unicode order and apply RTL through the document format.
