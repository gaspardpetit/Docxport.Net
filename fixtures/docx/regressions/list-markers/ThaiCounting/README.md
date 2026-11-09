# Thai counting list markers

Retained DOCX and Word PDF cover 271 labels: Thai words, contextual final one,
negative values, million groups, the billion boundary and the signed minimum.
The manifest records logical Word ListString labels, raw PDF extraction,
starts and artifact hashes. Ignoring PDF-inserted whitespace makes every
extracted label agree; export tests assert exact logical text.

The [Word capture](../../../../../truth/pdf-word/regressions/list-markers/ThaiCounting/)
uses the standard grouped truth tree. Tests recreate numbering in memory for
Markdown, HTML and plain text. Shared field-number formatting is unchanged.
