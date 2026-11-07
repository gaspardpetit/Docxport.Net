# Chicago list marker boundaries

Retained synthetic DOCX sources and visible Microsoft Word PDF captures cover
zero, the four-symbol cycle, the 30-repetition rollover, large counters,
and negative abstract starts versus standalone start overrides.

Word cycles through *, dagger, double dagger and section sign, repeating each
symbol at most 30 times before the repetition count wraps. Nonpositive counters
produce empty marker content; invalid negative abstract starts suppress numbering.
The generic field formatter retains its existing behavior.

`word-evidence.json` records source/capture hashes and every visible marker.
PDF captures are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/ChicagoBoundaries/).
Tests reconstruct these structures in memory and exercise Markdown, HTML and text.
