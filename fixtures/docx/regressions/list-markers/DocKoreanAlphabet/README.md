# Korean alphabet markers through DOC

All fixtures contain only synthetic `Item0` through `Item5` paragraphs.

Ganada and Chosung text rendering was already supported, but DOC export rejected both formats and DOCX projection rejected their binary codes. This change adds the corresponding mappings in both directions. Microsoft [MS-OSHARED 2.2.1.3 MSONFC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659) defines Ganada as 0x18 and Chosung as 0x19.

Twelve cases cover each format at depths 0 and 8, starting at 1, 13, and 19. Six paragraphs in each case cover ordinary progression, the 14-letter alphabet boundary, and continued cyclic numbering. Marker fonts explicitly use Malgun Gothic for consistent Korean glyph coverage.

`source/` contains generated DOCX inputs; `fixtures/doc/regressions/list-markers/KoreanAlphabet/` contains exported DOC files; `projected/` contains their DOCX projections. `word-evidence.json` records relative paths, SHA-256 hashes, and editable Word list strings for all 36 documents. Complete independently extracted Word PDF text is checked against every marker and paragraph. Word captures were made read-only on 2026-11-13.

`baseline-errors.json` records the 12 export failures on main. There is no before-fix DOC artifact because export threw before producing one. Tests additionally verify binary format codes, projected format declarations, OOXML schema, source numbering immutability, and five text visitors. No real document content or PDF/layout implementation changes are submitted.
