# Negative-start Word evidence

These synthetic DOCX sources and captured Microsoft Word PDFs come from the
marker audit. The manifest records exact hashes and visible marker sequences.
The tests recreate eight normalization controls and the signed-minimum
formatting control across plain Markdown, HTML and text and verify that source
numbering XML remains unchanged.

Word suppresses an abstract level with a negative start and discards a
replacement level with a negative start, including its associated override.
A standalone negative startOverride remains valid. A replacement with its own
nonnegative start takes precedence over a negative startOverride; if the
replacement omits its start, a negative override invalidates that replacement.

The ninth retained control, `override-min`, shows Word's empty numeric expansion
for Int32.MinValue. The decimal-boundary formatting fix now covers that case in
the negative-start regression matrix as well. The counter itself remains at the
signed minimum and advances normally on the next item.

No PDF visitor or layout source changes are needed. This change builds on the
restart-import normalization merged in #199, checking invalid starts before
restart declarations in the same private numbering view.

## Evidence locations

DOCX sources, manifests and supporting data stay together in this group.
Captured Word PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/NegativeStarts/).
Each manifest records repository-relative source and capture paths.
