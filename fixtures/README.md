# Document fixtures

This directory contains reusable test inputs. The real-world source documents
are in `docx/samples/`; their reviewed exporter results are in `../truth/`.

DOC and DOCX inputs are grouped by source format and purpose:

- `docx/features/`: cross-cutting features, including `text-boxes/`.
- `docx/probes/`: exploratory and targeted regression probes.
- `docx/legacy/` and `doc/legacy/`: paired Word-authored compatibility cases.
- `docx/regressions/` and `doc/regressions/`: focused regression cases.
- `docx/native-oracle/` and `doc/native-oracle/`: native Word comparison cases.
- `docx/samples/`: larger real-world DOCX inputs shared by exporter tests.
- `omml/`: equation inputs and paired expected-expression assets.

The `regressions/` name describes the fixture collection, not its current
pass/fail status. Exporter-specific known gaps belong in a separate reviewed
manifest. Generated exports and raster comparisons belong under `../output/`,
not here. Word's temporary `~$` lock files are not fixtures.

Exporter baselines are grouped by test suite under `../truth/html-export/`,
`../truth/markdown-export/`, and `../truth/plain-text-export/`. Test output uses
the same suite and sample names under ignored `../output/`. After reviewing a
generated result, run `scripts/update-expected-from-tests.sh` to replace only
existing baselines with their corresponding generated files.
