# Numeric marker sizes in RTL paragraphs

These 19 synthetic documents contain only `Item0`–`Item2` and numeric list labels.
Word uses the complex-script size (`szCs`) for an all-ASCII-digit label in a
right-to-left paragraph, including a literal numeric bullet. The marker keeps
its ASCII font face. Rich HTML previously omitted that size and inherited the
body size instead (12 pt rather than Word's 15 pt in the RTL examples).

The fixtures cover direct paragraph direction, a based-on style chain, the
default paragraph style, document paragraph defaults, and explicit false values
at nearer inheritance levels. Three further controls inherit the marker size
from paragraph-mark properties, a paragraph style, or document run defaults.
The eight LTR/reset examples retain the previous HTML behavior.

`source/` contains the editable DOCX inputs. `baseline/` contains rich HTML from
main at `4d92fd0d`; `html/` contains the corrected rich HTML. The baseline producer
was the previous PR worktree, whose production sources were verified identical
to that main commit. Word PDFs are grouped under
`truth/pdf-word/regressions/list-markers/NumericRtlMarkerSize/source/`.

`word-evidence.json` records input/output hashes, Word's editable `ListString`
values, and independently extracted PDF marker fonts and point sizes. HTML
hashes use LF line endings so they also match Git's stored blobs. All 19 sequences
match Word. Eleven cases require a 15 pt HTML marker override; the
eight LTR/reset controls do not. Plain HTML and plain Markdown from the actual
CLI are byte-identical to the baseline for all 19 inputs.

The automated tests validate the source XML, check marker text and CSS, exercise
plain HTML, and verify that exporting does not mutate source XML. This change
does not implement the separate DOC contextual-formatting aliases or change PDF
layout.
