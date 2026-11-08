# Japanese counting list markers

Retained synthetic DOCX sources and visible Microsoft Word PDF captures cover
japaneseCounting from zero through unit transitions at 10, 100, 1000 and 10000,
and its last supported value 999999 followed by empty expansions.
Word omits one before ten, hundred and thousand within each four-digit group,
and uses U+3007 for standalone zero. Six sequences are retained.

`word-evidence.json` records source/capture hashes, starts and visible labels.
PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/JapaneseCounting/).
Tests reconstruct numbering in memory for Markdown, HTML and plain text.
Generic field formatting retains its current decimal fallback.
