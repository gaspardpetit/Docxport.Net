# Bullet and no-sequence placeholders through DOC

Four original inputs reuse `../BulletPlaceholders`: bullet, decimal and Roman ancestors, plus a literal quoted-bullet control. `source/none-parent-none.docx` adds the no-sequence counterpart. Five editable binary exports are under `fixtures/doc/regressions/list-markers/BulletPlaceholders`; `projected/` contains their API DOCX projections.

Word captures on 2026-11-12 confirm all 15 source/DOC/projected sequences (90 editable ListString labels). Independent PDF extraction confirms the visible labels and all six body items in each document. The manifest retains source/output/capture paths and hashes. PDFs are grouped under `truth/pdf-word/regressions/list-markers/DocBulletPlaceholders/{source,doc,projected}`.

Before the fix, Word displayed literal `BEGIN%1/%2END` in the binary DOC where the source displayed `BEGIN3/END` or `BEGINIII/END`. Projection alone hid this defect by interpreting the literal percent references again as OOXML. Encode and decode the binary placeholder table for bullet and no-sequence labels as well as numeric labels. The literal quoted-bullet control remains unchanged. Literal markup escaping, legal-numbering flags and other DOC numbering semantics remain separate scopes.
