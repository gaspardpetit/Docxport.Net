# Custom kana aliases

Eight retained DOCX/Word PDF pairs cover 72 labels for Aiueo and Iroha,
in half-width and full-width forms, including zero and signed boundaries.
Exact aliases reuse the established kana list formatter.
The manifest records source/capture hashes, starts, logical Word ListString
labels and raw PDF extraction. Ignoring extraction whitespace makes every
PDF label agree. Tests preserve exact logical text in Markdown, HTML and
plain text and verify source numbering XML remains unchanged.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomKanaAliases/)

Four direct-format signed-minimum controls were captured afresh in Word on
2026-11-09. Each displays BEGINEND, confirming an empty numeric expansion.
