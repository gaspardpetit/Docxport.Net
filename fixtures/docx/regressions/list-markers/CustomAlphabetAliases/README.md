# Custom regional alphabet aliases

Nine retained DOCX/Word PDF pairs cover 153 labels for Arabic alphabetical and
Abjad, Hebrew numeral and alphabet, Hindi letter sets, Thai letters, and Russian
lower/upper alphabets. Exact aliases reuse the established list formatters.
The manifest records source/capture hashes, starts, logical Word ListString
labels and raw PDF extraction. Ignoring extraction whitespace and invisible
U+200C/U+200E/U+200F controls makes every PDF label agree without reordering.
Tests preserve exact logical text, including controls, in Markdown, HTML and
plain text and verify source numbering XML remains unchanged. This does not
claim general bidi layout correctness.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomAlphabetAliases/)
