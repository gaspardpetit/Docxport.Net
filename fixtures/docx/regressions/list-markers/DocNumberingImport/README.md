# DOC numbering import view

Eleven API-projected DOCX documents are paired with editable binary DOC exports under fixtures/doc/regressions/list-markers/NumberingImport. Original inputs reuse the NumberingProcessContent DOCX group. Both outputs were generated through DxpDocExport.Export followed by DxpDocToDocx.Project.

Fresh Word ListString and PDF captures on 2026-11-12 confirm all 33 input/DOC/projected sequences (99 labels). Nonempty ListString values confirm editable numbering. PDF extraction independently confirms visible labels and Item0-Item2 text. The manifest retains all source/output/PDF hashes and input provenance.

The fix uses a private imported numbering view in the DOC list reader, with its Office 2007 namespace compatibility profile. Newer custom-format choices use their declared fallback; text exports retain their newer compatibility profile. Other DOC format mappings, restart projection and style/story handling remain separate scopes.
