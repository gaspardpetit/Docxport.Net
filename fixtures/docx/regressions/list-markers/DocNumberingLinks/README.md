# Numbering-style links through DOC

All fixtures are synthetic `Item0` through `Item5` paragraphs.

A level-less abstract definition can acquire its levels through `numStyleLink`, a numbering style's numId, and a matching reverse `styleLink` on the target definition. Word keeps the wrapper's counter identity separate from the target. DOC import previously copied the empty wrapper, dropping its markers.

The importer now reuses the existing numbering link resolver with DOC's existing Office 2007 compatibility-normalized view. It imports resolved levels while retaining each original abstract definition's identity. A missing reverse link still produces no wrapper markers.

Twelve cases cover depths 0, 1, and 8, valid direct links, inherited numbering-style references, a target instance start override, and missing-reverse-link controls. Source DOCX, exported DOC, and projected DOCX files have Word editable list strings and independently checked complete Word PDF text. `word-evidence.json` records relative paths, hashes, and labels. Word captures were made read-only on 2026-11-13.

Source files are in `source/`; DOC files are under `fixtures/doc/regressions/list-markers/NumberingLinks/`; projections are in `projected/`. Three pre-fix projections and Word PDFs in `before/` show the missing wrapper labels. `before-word-evidence.json` records those failures.

Tests verify numbering/style XML immutability, two independent counter owners, absence of spurious replacement levels, projected schema validity, and complete marker/body output from five text visitors. No PDF/layout implementation changes are included.
