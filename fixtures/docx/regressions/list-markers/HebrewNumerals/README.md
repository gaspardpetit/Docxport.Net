# Hebrew numeral list markers

The retained synthetic DOCX and visible Word PDF capture cover Hebrew1 units,
tens, hundreds, the special 15/16 forms, rollover after 392, nonpositive starts
and large counters through the signed maximum. The capture contains 166 labels.

`labels` records Word ListString values in logical order. PDF text extraction
adds incidental spaces to some Hebrew runs; removing extraction whitespace
makes every PDF label equal to Word ListString. `pdfExtractedLabels` retains
that unmodified extraction for review. No character reordering is performed.

`word-evidence.json` records source/capture hashes, starts and labels.
The PDF is in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/HebrewNumerals/).
Tests recreate numbering for Markdown, HTML and plain text.
Generic field formatting retains its decimal fallback.
