# Shared DOC counter ownership

Eighteen synthetic cases cover levels 0, 1 and 8, alternating two numbering
instances in the order 1/1/2/2/1/2. Modes cover no override, explicit starts
1 and 7, formatting-only Roman replacements, combined start/formatting
replacements, and independent definitions with identical level formatting.
Combined overrides use matching starts; conflicting starts remain separate work.

DOC list instances referencing the same definition share a counter owner.
DOCX projection must retain one abstract list per DOC definition ID. A separate
abstract per instance incorrectly resets continuation and isolates explicit
starts from later items in the other instance. Independent definitions must
stay separate even when their formatting is identical. Instance-specific start
and formatting overrides remain on their individual numbering instances.

`source` contains generated DOCX inputs; `projected` contains projections of
editable DOCs under `fixtures/doc/regressions/list-markers/SharedOwner`.
Fresh Word captures of all three formats are grouped under
`truth/pdf-word/regressions/list-markers/DocSharedOwner`. The hash manifest
records 54 sequences / 324 editable list positions; independent PDF extraction
checks complete labels and Item0–Item5 body text. Three pre-fix projected DOCX
files and Word PDFs are retained under `before`, with their own hash manifest.

Baseline: fifteen shared-owner failures; three independent-definition controls
pass. Final eighteen cases pass, along with 453 prior marker/editable-field
controls. Tests check DOC definition/instance counts, projected owner identity,
replacement preservation, source numbering immutability, five text visitors,
complete labels/body and projected OOXML schema validity.
