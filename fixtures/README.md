# Document fixtures

This directory contains reusable test inputs. The real-world source documents
are in `docx/samples/`; their reviewed exporter results are in `../truth/`.

DOC and DOCX inputs are grouped by source format and purpose:

- `docx/features/`: cross-cutting features, including `text-boxes/`.
- `docx/probes/`: exploratory and targeted regression probes.
- `docx/probes/shared-*`: small DOCX cases intended for reuse across PDF, HTML,
  Markdown, and text exporters. Regenerate them with
  `scripts/generate-shared-export-probes.ps1`.
- `docx/legacy/` and `doc/legacy/`: paired Word-authored compatibility cases.
- `docx/regressions/` and `doc/regressions/`: focused regression cases.
- `docx/native-oracle/` and `doc/native-oracle/`: native Word comparison cases.
- `docx/samples/`: larger real-world DOCX inputs shared by exporter tests.
- `doc/samples/`: the 14 DOC counterparts of `docx/samples/`, used to test
  native DOC import and DOC-to-DOCX projection. `TestLists.doc` also exercises
  list records, tracked revisions, and editable DOC round trips.
- `omml/`: equation inputs and paired expected-expression assets.

The `regressions/` name describes the fixture collection, not its current
pass/fail status. Exporter-specific known gaps belong in a separate reviewed
manifest. Generated exports and raster comparisons belong under `../output/`,
not here. Word's temporary `~$` lock files are not fixtures.

Sample baselines are grouped by test suite under `../truth/samples-export-html/`,
`../truth/samples-export-markdown/`, and `../truth/samples-export-text/`. Test output uses
the same suite and sample names under ignored `../output/`. After reviewing a
generated result, run `scripts/update-expected-from-tests.sh` to replace only
existing baselines with their corresponding generated files.

`docx/features/mixed-shaded-behind.docx` exercises mixed cell shading and
a page-anchored DrawingML gradient rectangle behind a table. The focused PDF
test checks that the gradient pattern and table text are both present.
DrawingML `outerShdw` on the rectangle is not yet rendered.

`docx/features/WordBorderStyles3x3.docx` places a 3×3 table on each page,
covering Word's border styles and shared cell edges. The PDF flow tests use it
to check style preservation, fixed cell widths, spacing inheritance, and double borders.

`docx/features/StyledRepeatedSpaces.docx` isolates the subtitle spacing issue
first noticed in that table fixture. Three styled lines add one space before or
after `·`; narrow table cells, split-run paragraphs, and a space group at a
run boundary check that a wrapped word is not indented by the extra blank.
All-caps styled runs check that source text transformations preserve the same spacing.
The Word reference is
`../truth/pdf-word/features/StyledRepeatedSpaces.pdf`. Regenerate the DOCX
with `tools/doc-docx/Generate-StyledRepeatedSpaceProbe.py`.
