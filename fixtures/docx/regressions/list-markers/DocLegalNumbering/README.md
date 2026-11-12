# Legal numbering through binary DOC

Twelve source/DOC/projected triples cover four ancestor formats (upper Roman, decimalZero, bullet and none) with legal numbering enabled, absent and explicitly disabled. Source inputs are under `source/`; API DOCX projections are under `projected/`. Editable binary outputs are under `fixtures/doc/regressions/list-markers/LegalNumbering`.

Fresh Word ListString and PDF captures on 2026-11-12 confirm all 36 sequences (216 editable labels). PDF extraction independently confirms visible labels and Item0-Item5 body text. The manifest retains source/output/capture paths and hashes. Word PDFs are grouped by source/doc/projected under `truth/pdf-word/regressions/list-markers/DocLegalNumbering`.

The flag must survive OOXML import, binary DOC writing/reading and DOCX projection. Before this fix, the four legal cases lost the flag after a DOC hop: e.g. `BEGIN3/1END` became `BEGINIII/iEND`. Ordinary and explicit-false controls retain their original formats; decimalZero ancestors retain their leading zero under legal numbering.

The binary flag is fLegal in [MS-DOC LVLF](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/500d1b05-cd15-46af-bc8f-03bc290d1bdd). Fresh Word-visible output is the acceptance oracle. Other DOC paths, including formatting replacements and their projection, remain separate scopes. PDF/layout code is untouched.
