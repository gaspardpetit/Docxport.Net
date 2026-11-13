# Sparse abstract levels through DOC export

Thirteen cases cover dense definitions, missing roots, a single non-root level,
internal gaps, missing trailing levels, and placeholders referencing missing
ancestors. Defined levels use starts of 3; omitted levels display no marker and
contribute zero when referenced by another level. Levels 1, 2 and 8 are covered.

Before the fix, nine sparse definitions were rejected. In multi-level lists,
the binary writer's generic padding also invented decimal labels starting at 1.
For `missing-tail-2`, Word displayed blank labels in the source but `1.` and `2.`
in the DOC and its projection. Private DOCX import now materializes blank,
zero-start slots before calling the binary writer, preserving that writer's
existing public-model defaults.

`source` holds DOCX inputs; `projected` holds API projections of the matching
editable files under `fixtures/doc/regressions/list-markers/SparseAbstract`.
Fresh Word PDFs of all three formats are grouped under
`truth/pdf-word/regressions/list-markers/DocSparseAbstract`. The manifest records
hashes and ListString values for 39 sequences / 234 list positions, including
empty labels; independent PDF extraction checks the complete labels and six
body items. The two pre-fix missing-tail PDFs are retained under `before`.

Tests check the imported binary slots, source numbering immutability, five text
visitors, complete body/marker output, projected schema validity, and continued
rejection of duplicate/out-of-range level indexes.

This change covers nonempty abstract definitions with an explicit non-root
level. Completely empty definitions are covered by `DocEmptyAbstract`.
Root-only definitions used at higher levels through paragraph/style/replacement
context remain separate work.
