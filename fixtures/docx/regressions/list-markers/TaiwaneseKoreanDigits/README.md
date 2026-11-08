# Taiwanese and Korean digit list markers

Retained synthetic DOCX sources and visible Microsoft Word PDF captures cover
taiwaneseDigital and koreanDigital digit substitution, negative start overrides,
zero and large counters through the signed maximum.
Word leaves negative expansions empty and preserves the full nonnegative value,
including values above 32767. Taiwanese zero uses U+25CB, matching the capture.

`word-evidence.json` records source/capture hashes, starts and all 34 labels.
PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/TaiwaneseKoreanDigits/).
Tests reconstruct numbering in memory for Markdown, HTML and plain text.
Generic field formatting retains its current decimal fallback.
