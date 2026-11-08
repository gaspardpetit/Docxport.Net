# Arabic alphabetical and Abjad list markers

Two retained synthetic DOCX sources and visible Word PDF captures cover both
28-letter orders, repeated symbols, rollover after 392, nonpositive starts and
large counters through the signed maximum (332 labels).

`labels` records logical Word ListString values. Each alphabetical symbol has
a trailing U+200C non-joiner; each Abjad symbol has a leading non-joiner.
PDF extraction omits those controls and adds incidental spaces; removing those
differences makes every PDF label agree. `pdfExtractedLabels` retains the raw
extraction. No character reordering is performed.

`word-evidence.json` records source/capture hashes, starts and labels.
PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/ArabicAlphabets/).
Tests recreate numbering for Markdown, HTML and plain text.
Generic field formatting retains its decimal fallback.
