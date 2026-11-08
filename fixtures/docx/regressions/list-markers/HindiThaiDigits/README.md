# Hindi and Thai digit list markers

Retained synthetic DOCX sources and visible Microsoft Word PDF captures cover
Hindi and Thai digit substitution, negative start overrides, zero, rollover
after 32767 and large counters through the signed maximum.
Word leaves negative expansions empty and wraps positive display values after
32767, independently of the underlying counter. Zero uses the respective digit.

`word-evidence.json` records source/capture hashes, starts and visible labels.
PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/HindiThaiDigits/).
Tests reconstruct numbering in memory for Markdown, HTML and plain text.
Generic field formatting retains its current decimal fallback.
