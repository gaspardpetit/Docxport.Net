# Thai letter list markers

The retained synthetic DOCX and visible Word PDF capture cover the complete
41-letter alphabet, repeated symbols, rollover after 1230, negative start
overrides, zero and large counters through the signed maximum.
Word expands zero to a space and negative values to empty text.
The capture contains 166 visible labels.

`word-evidence.json` records source/capture hashes, starts and visible labels.
The PDF is in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/ThaiLetters/).
Tests recreate numbering for Markdown, HTML and plain text.
Generic field formatting retains its decimal fallback.
