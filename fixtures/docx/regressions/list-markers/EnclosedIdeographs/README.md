# Enclosed ideograph list markers

The retained synthetic DOCX and visible Microsoft Word PDF capture cover
ideographEnclosedCircle across the complete symbol range, zero, negative start
overrides, large positive counters and signed integer boundaries.
Despite the format name, Word uses parenthesized ideographs U+3220 through
U+3229 for values 1 through 10. Other values use decimal text, except the
signed minimum, whose marker expansion is empty.

`word-evidence.json` records source/capture hashes, starts and visible labels.
The PDF is in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/EnclosedIdeographs/).
Tests reconstruct numbering in memory for Markdown, HTML and plain text.
Generic field formatting retains its current decimal fallback.
