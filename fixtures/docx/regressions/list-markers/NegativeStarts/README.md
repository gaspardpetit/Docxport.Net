# Negative-start Word evidence

These synthetic DOCX sources and captured Microsoft Word PDFs come from the
marker audit. The manifest records exact hashes and visible marker sequences.
The tests recreate the eight normalization controls across plain Markdown,
HTML and text and verify that source numbering XML remains unchanged.

Word suppresses an abstract level with a negative start and discards a
replacement level with a negative start, including its associated override.
A standalone negative startOverride remains valid. A replacement with its own
nonnegative start takes precedence over a negative startOverride; if the
replacement omits its start, a negative override invalidates that replacement.

The ninth retained control, `override-min`, demonstrates a separate formatting
defect on current main: Word shows an empty numeric expansion for Int32.MinValue,
while main prints the signed decimal value. That formatting fix is explicitly
deferred and excluded from this normalization test matrix. Its source and Word
capture are retained here for the follow-up; no passing assertion of equivalence
is made for it.

No PDF visitor or layout source changes are needed. This change builds on the
restart-import normalization merged in #199, checking invalid starts before
restart declarations in the same private numbering view.

## Evidence locations

DOCX sources, manifests and supporting data stay together in this group.
Captured Word PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/NegativeStarts/).
Each manifest records repository-relative source and capture paths.
