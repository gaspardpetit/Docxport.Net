# Word list-marker evidence

These synthetic DOCX sources and captured Microsoft Word PDFs are selected
controls from the completed marker audit. The manifest identifies the exact
source and PDF bytes. Regression tests recreate the tested structures with
generated text and do not require Word during a normal test run.

The visible Word result takes precedence over an assumption based only on
OOXML prose. Coverage is limited to the supplied cases, rather than a claim of
universal conformance. Layout and marker-font behavior are outside this change.

## Evidence locations

DOCX sources, manifests and supporting data stay together in this group.
Captured Word PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/Resolution/).
Each manifest records repository-relative source and capture paths.
