# Vietnamese counting list markers

Retained DOCX and Word PDF cover 271 labels: lexical forms, combining marks,
zero, empty negative expansions, hundreds and rollover after 1000.
The manifest records logical Word ListString labels, raw PDF extraction,
starts and artifact hashes. Ignoring PDF-inserted whitespace makes every
extracted label agree; export tests assert exact logical spelling and spaces.

The [Word capture](../../../../../truth/pdf-word/regressions/list-markers/VietnameseCounting/)
uses the standard grouped truth tree. Tests recreate numbering in memory for
Markdown, HTML and plain text. Shared field-number formatting is unchanged.
