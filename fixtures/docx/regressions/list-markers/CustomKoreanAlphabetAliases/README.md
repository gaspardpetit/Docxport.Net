# Custom Korean alphabet aliases

Two retained DOCX/Word PDF pairs cover 34 Ganada/Chosung labels, including
cyclic output, zero, negative starts and signed boundaries. Exact aliases
reuse the established Korean alphabet formatter.
The manifest records source/capture hashes, starts, logical Word ListString
labels and raw PDF extraction. Ignoring extraction whitespace makes every
PDF label agree. Tests preserve exact logical text in Markdown, HTML and
plain text and verify source numbering XML remains unchanged.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomKoreanAlphabetAliases/)
