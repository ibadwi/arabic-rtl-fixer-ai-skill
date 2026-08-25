# Unicode and Arabic writing

Use readable, idiomatic Arabic rather than literal translation. Match terminology, numerals, punctuation, and tone to the requested locale and audience. Keep product names and established Latin abbreviations unchanged unless the user supplies an Arabic form.

## Direction and hidden characters

Direction belongs to the paragraph or container. Right alignment controls placement only; it does not establish RTL. Preserve logical source order, including mixed examples such as:

```text
ارتفع الأداء بنسبة 25% في Q4 2026.
يعتمد النظام على GPS وTCP/IP.
```

Do not wrap Arabic strings with invisible bidi controls. When inspecting supplied text, check for code points `U+202A`-`U+202E`, `U+2066`-`U+2069`, `U+200E`, `U+200F`, and `U+061C`. Report or safely remove accidental controls before delivery. Do not globally strip `U+200C` ZWNJ or `U+200D` ZWJ.

Use Unicode normalization carefully. NFC is generally suitable, but do not normalize identifiers, signatures, code, URLs, or user data when exact code points matter.

## Mixed content

Use one base direction per paragraph. Short Latin terms and numbers can remain inline in a native RTL paragraph. Split long English passages, URLs, email addresses, code, citations, and file paths into LTR containers. Avoid manual spaces and punctuation movement intended only to fix one preview.

For Arabic-first tables and processes, place the first semantic item on the right. Do not mirror logos, maps, mathematical axes, media controls, or universally directional symbols without a semantic reason.
