# Custom digit aliases

Three retained DOCX/Word PDF pairs cover 51 labels for Hindi, Thai and Taiwanese
digit aliases: zero, negative values, signed boundaries and large counters.
Exact aliases reuse the established list digit formatters, including the
Hindi/Thai rollover and the full Taiwanese counter range.
The manifest records source/capture hashes, starts, logical Word ListString
labels and raw PDF extraction. Ignoring extraction whitespace makes every
PDF label agree. Tests preserve exact logical text in Markdown, HTML and
plain text and verify source numbering XML remains unchanged.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomDigitAliases/)
