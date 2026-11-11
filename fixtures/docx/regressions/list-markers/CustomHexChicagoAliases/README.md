# Custom hexadecimal and Chicago aliases

Two DOCX/Word PDF pairs cover 34 labels for the exact hexadecimal and Chicago
symbol aliases, including negative starts, zero, range limits and large starts.
The manifest records source/capture hashes, starts, logical Word labels and raw
PDF extraction. Ignoring extraction whitespace makes every PDF label agree.
Tests verify exact Markdown, HTML and plain text labels and unchanged source
numbering XML. These aliases reuse the existing list-specific formatters.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomHexChicagoAliases/)
