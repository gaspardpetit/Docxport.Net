# Empty abstract list definitions

Sixteen cases cover empty abstract definitions at levels 0, 1 and 8 with no
override, a start override, a formatting replacement, and both. Four ordinary
single-root definitions control the existing supported path. Each has six
editable list paragraphs; child-level cases include parent items and resets.

Word leaves undefined levels blank. Replacement-only root markers start at
zero; explicit starts activate 7. Child replacements after parent items start
at 7 and reset to 7 after the next parent. Private DOCX import now accepts the
empty abstract and materializes the same nine blank, zero-start slots used for
sparse definitions. Existing single-root definitions retain one slot.

`source` contains DOCX inputs; `projected` contains API projections of the
editable DOCs under `fixtures/doc/regressions/list-markers/EmptyAbstract`.
Fresh Word captures of all three formats are grouped under
`truth/pdf-word/regressions/list-markers/DocEmptyAbstract`. The hash manifest
records 48 sequences / 288 editable list positions, including blank labels;
independent PDF extraction checks the complete labels and body text.

Tests verify imported slots, source numbering immutability, five text visitors,
complete labels/body, projected schema validity, and the ordinary root controls.
Root-only definitions used at higher levels through paragraph/style/replacement
context remain separate work.
