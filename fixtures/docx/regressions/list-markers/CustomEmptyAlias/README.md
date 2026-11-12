# Empty custom format alias

Four DOCX/Word PDF pairs cover 68 labels captured in Word on 2026-11-11.
An exact empty custom format matches direct none: its own numeric placeholder
expands to empty, but surrounding literal text survives. Literal-only labels
also survive in both forms. Starts span negatives, zero and signed boundaries.
The manifest records source/capture hashes, starts, logical Word labels and raw
PDF extraction. Ignoring extraction whitespace makes every PDF label agree.
Tests verify exact Markdown, HTML and plain text labels and unchanged source
numbering XML. No suppression of literal-only labels is introduced.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomEmptyAlias/)
