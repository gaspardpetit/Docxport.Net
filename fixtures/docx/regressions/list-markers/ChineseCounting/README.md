# Chinese and Taiwanese counting list markers

Eight retained synthetic DOCX sources and visible Word PDF captures cover
chineseCounting and taiwaneseCounting. Word uses unit counting below 100,
then digit-by-digit ideographs from 100 onward. Embedded zeros use U+25CB.
Sequences cover the 99/100 transition and digit transitions at 1000 and 10000.

`word-evidence.json` records source/capture hashes, starts and visible labels.
PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/ChineseCounting/).
Tests recreate numbering for Markdown, HTML and plain text.
Generic field formatting retains its decimal fallback.
