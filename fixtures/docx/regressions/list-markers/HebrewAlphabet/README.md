# Hebrew alphabet list markers

The retained synthetic DOCX and visible Word PDF capture cover Hebrew2's
22-letter alphabet, prefixing tav on later cycles, rollover after 392,
nonpositive starts and large counters through the signed maximum (166 labels).

`labels` records logical Word ListString values, including its leading U+200F
right-to-left mark on positive markers. PDF extraction omits that invisible
mark and adds incidental spaces; removing those differences makes every PDF
label agree. `pdfExtractedLabels` retains raw extraction for review.
No character reordering is performed.

`word-evidence.json` records source/capture hashes, starts and labels.
The PDF is in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/HebrewAlphabet/).
Tests recreate numbering for Markdown, HTML and plain text.
Generic field formatting retains its decimal fallback.
