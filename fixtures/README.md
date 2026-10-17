# Document fixtures

This directory contains reusable test inputs. The existing real-world
`samples/` tree remains in place until its source documents and expected
outputs are migrated in a separate change.

DOC and DOCX inputs are grouped by source format and purpose:

- `docx/features/`: cross-cutting features, including `text-boxes/`.
- `docx/probes/`: exploratory and targeted regression probes.
- `docx/legacy/` and `doc/legacy/`: paired Word-authored compatibility cases.
- `docx/regressions/` and `doc/regressions/`: focused regression cases.
- `docx/native-oracle/` and `doc/native-oracle/`: native Word comparison cases.
- `omml/`: equation inputs and paired expected-expression assets.

The `regressions/` name describes the fixture collection, not its current
pass/fail status. Exporter-specific known gaps belong in a separate reviewed
manifest. Generated exports and raster comparisons belong under `output/`,
not here. Word's temporary `~$` lock files are not fixtures.
