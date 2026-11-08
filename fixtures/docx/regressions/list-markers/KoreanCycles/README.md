# Korean alphabetic list marker cycles

Retained synthetic DOCX sources and visible Microsoft Word PDF captures cover
Ganada and Chosung through repeated complete 14-symbol cycles, negative start
overrides, zero, large counters and signed integer boundaries.
Word cycles a single symbol instead of repeating or combining symbols.
Nonpositive counters use decimal text, except the signed minimum, whose
expansion is empty.

`word-evidence.json` records source/capture hashes, starts and all 264 labels.
PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/KoreanCycles/).
Tests reconstruct numbering in memory for Markdown, HTML and plain text.
Generic field formatting retains its current decimal fallback.
