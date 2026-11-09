# Custom Latin letter and Roman aliases

Four retained DOCX/Word PDF pairs cover 68 labels for custom A, a, I and i:
letter repetition, Roman rollover, zero, negative starts and signed boundaries.
The manifest records source/capture hashes, starts, logical Word ListString
labels and raw PDF extraction. Ignoring extraction whitespace makes all PDF
labels agree. Tests preserve exact logical text in Markdown, HTML and plain
text and verify the caller's numbering XML remains unchanged.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomLatinAliases/)
