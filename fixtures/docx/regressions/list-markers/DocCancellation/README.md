# Explicit numbering cancellation

Twelve cases cover levels 1 and 8 with direct numId=0, direct numId=0 plus
ilvl, cancellation inherited from a derived paragraph style, ordinary inherited
numbering, explicit re-enabling over the cancelling style, and numbering resumed
after individual cancelled items. Parent markers remain 2 and 3; children start
at 5 and reset after a parent. Cancelled items have no list label and do not
consume a child counter.

The private DOC visitor preserves numId=0 as an explicit zero list override,
rather than treating it as absent. DOCX projection emits numId=0 for that zero
override, preserving cancellation in both direct formatting and paragraph
styles. Positive numbering IDs retain their existing path.

`source` contains DOCX inputs; `projected` contains API projections of editable
DOCs under `fixtures/doc/regressions/list-markers/Cancellation`. Fresh Word
captures of all three formats are grouped under
`truth/pdf-word/regressions/list-markers/DocCancellation`. The hash manifest
records 36 sequences / 216 editable list positions including blanks; independent
PDF extraction checks complete labels and body text. Three pre-fix DOCs and
Word PDFs are retained in `before` subgroups with a separate hash manifest:
Word incorrectly numbers cancelled children 5 and 6.

Baseline: six cancellation failures, four passing inherited/re-enabled controls.
The marker/editable-field regression set passes 451 tests, and the final twelve
group cases also pass, including two additional resume cases. Tests check source
numbering/style immutability, imported slots, all five text visitors, complete
labels/body and projected OOXML schema validity.
