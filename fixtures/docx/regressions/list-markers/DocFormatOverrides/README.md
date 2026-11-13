# DOC formatting replacement activation

Twelve cases cover levels 0, 1 and 8 with no override, start only, formatting
only, and both overrides. The abstract start is 2; explicit and replacement
starts are 7. Replacement levels use upper Roman numbers, making preservation
of their formatting visible independently of their activation.

Word displays II–VII for formatting-only replacements and VII–XII for combined
overrides. The controls without replacement formatting use decimal 2–7 or
7–12. Before the fix, DOCX projection flattened replacements into abstract
definitions and incorrectly displayed VII–XII for formatting-only lists.

`source` contains the original DOCX inputs; `projected` contains API projections
of the editable binary DOCs grouped under
`fixtures/doc/regressions/list-markers/FormatOverrides`. Fresh Word captures of
all three formats are under `truth/pdf-word/regressions/list-markers/DocFormatOverrides`.
`word-evidence.json` records hashes, editable ListString values, and independently
extracted PDF labels for all 36 sequences / 216 markers. The three incorrect
pre-fix projection PDFs are retained under `before/projected` with hashes and
extracted labels in the same manifest.

Tests verify source and projected output through five text visitors, body item
preservation, and projected OOXML schema validity. The earlier `DocOverrideFlags`
tests now also check formatting-only projections. Existing missing-start controls
remain covered. Shared list ownership and other start/restart contexts remain
separate work; this change preserves each instance's existing abstract identity.

The projection now retains [OOXML level overrides](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.leveloverride?view=openxml-3.0.1)
instead of promoting them into abstract levels. Word-visible activation is the
correctness reference for these captures.
