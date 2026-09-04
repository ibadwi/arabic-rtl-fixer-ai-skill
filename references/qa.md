# Arabic acceptance QA

Test representative content containing Arabic-only text, Western and Arabic-Indic digits, `25%`, a leading Latin acronym, an inline acronym, parentheses, `v2.1`, dates, `TCP/IP`, a URL in an LTR container, native bullets, and multiple paragraphs.

Check separately:

1. **Content:** spelling, terminology, locale, punctuation, no accidental omissions.
2. **Unicode/structure:** logical order, no accidental bidi controls, native RTL, Arabic language metadata, native lists, appropriate font slots.
3. **Visual layout:** connected glyphs, word and number order, punctuation, wrapping, bullet side, table order, overflow, overlap, contrast, and substitution.
4. **Target application:** open the final artifact in the application named by the user. Record untested applications as `NOT RUN` rather than inferring compatibility.

For OOXML repair, also verify that the source remains unchanged, the final file opens as a valid ZIP package, unrelated binary parts remain present, and running repair again produces no additional structural changes.

Recommended delivery report:

```text
Arabic content QA: PASS/FAIL
RTL and Unicode structural QA: PASS/FAIL
Layout QA: PASS/FAIL
Target-application QA: PASS/NOT RUN
Font portability: verified / installed-font dependent / PDF supplied
```
