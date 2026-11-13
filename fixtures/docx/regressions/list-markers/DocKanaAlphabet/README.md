# Japanese kana markers through DOC

All fixtures contain only synthetic `Item0` through `Item5` paragraphs.

Aiueo, Iroha, AiueoFullWidth, and IrohaFullWidth already render in text exports. DOC export rejected them, and DOCX projection rejected their binary codes. This change adds all four mappings in both directions, without changing the numbering renderer.

Microsoft [MS-OSHARED 2.2.1.3 MSONFC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659) defines Aiueo as 0x0C, Iroha as 0x0D, AiueoFullWidth as 0x14, and IrohaFullWidth as 0x15.

Sixteen cases cover every format at depth 0 with starts 1, 22, and 46, and at depth 8 with start 46. Each contains six successive markers. These exercise ordinary numbering, width selection, archaic Iroha kana without half-width equivalents, and alphabet cycling. Marker fonts explicitly use Yu Gothic.

`source/` contains generated DOCX inputs; `fixtures/doc/regressions/list-markers/KanaAlphabet/` contains exported DOC files; `projected/` contains their DOCX projections. `word-evidence.json` records relative paths, SHA-256 hashes, and editable list strings for all 48 documents. Complete independently extracted Word PDF text is checked against every marker and paragraph. Word captures were made read-only on 2026-11-13.

`baseline-errors.json` records the 16 export failures on main; no DOC artifact was produced before the fix. Tests additionally verify binary codes, projected format declarations/schema, source immutability, and complete output from five text visitors. No real document content or PDF/layout implementation changes are included.
