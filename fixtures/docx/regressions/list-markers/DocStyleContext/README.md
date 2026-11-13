# Paragraph/style numbering pairs

Six cases cover levels 1 and 8 with a direct level inheriting its numId,
a direct numId inheriting its level, and a fully inherited pair as controls.
Split-pair cases define only abstract level 0. Controls also define the child
level, starting at 5. All cases contain parent items and child restarts.

Word leaves undefined children blank and keeps parent markers 2 and 3.
Fully inherited controls show child markers 5 and 6, restarting at 5 after
the next parent. Previously both split-pair directions became root markers
2–7 in DOC, shifting the parent sequence. Two pre-fix DOCs and Word PDFs are
retained in `before` subgroups with their own hash manifest.

The importer now collects effective paragraph/style pairs before choosing
positional DOC slots. The DOC visitor uses the same existing style resolver
to retain both properties when direct formatting supplies exactly one of them.
Fully inherited pairs keep their existing representation. Numbering and style XML
remain unchanged in the caller's document.

`source` contains DOCX inputs; `projected` contains API projections of the
editable DOCs under `fixtures/doc/regressions/list-markers/StyleContext`.
Fresh Word captures of all three formats are grouped under
`truth/pdf-word/regressions/list-markers/DocStyleContext`. The hash manifest
records 18 sequences / 108 editable list positions, including blank labels;
independent PDF extraction checks complete labels and body text.

Tests verify imported slots, source numbering/style immutability, all five
text visitors, complete labels/body and projected schema validity. Baseline
four inheritance failures and two passing full-style controls; final focused
set including prior sparse/empty/start/override cases: 94 passed.
The expanded marker and editable-field round-trip regression set passes 441
tests, including the mixed-section styled-story control.

Explicit numId=0 cancellation is covered separately by `DocCancellation`.
