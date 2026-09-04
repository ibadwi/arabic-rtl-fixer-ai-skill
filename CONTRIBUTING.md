# Contributing

Contributions are welcome. Keep fixes focused on native, editable Arabic RTL behavior rather than renderer-specific visual workarounds.

Before opening a pull request:

1. Add or update a behavioral test for the change.
2. Run `python3 -m unittest discover -s tests -v`.
3. Run both command-line tools with `--help`.
4. Confirm that repair writes a new file and preserves unrelated OOXML package parts.
5. Do not commit client documents, personal data, credentials, or licensed fonts.

When reporting a document-specific bug, create a minimal synthetic reproduction whenever possible. If the original document is sensitive, do not attach it to a public issue.
