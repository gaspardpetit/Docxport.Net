# Traditional ideographic markers through DOC

All fixtures contain only synthetic `Item0` through `Item5` paragraphs.

IdeographTraditional, IdeographZodiac, and IdeographZodiacTraditional already render in text exports. DOC export rejected them, and DOCX projection rejected their binary codes. This change adds the three mappings in both directions, leaving the numbering renderer unchanged.

Microsoft [MS-OSHARED 2.2.1.3 MSONFC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659) defines these formats as 0x1E, 0x1F, and 0x20, respectively.

Nine cases cover each format at depth 0 with starts 1 and one before its boundary, and depth 8 starting one before that boundary. Each case has six successive markers. Word uses ten stems or twelve branches followed by decimal fallback for the first two formats; the combined format cycles after sixty positions. Word uses U+620D for the eleventh branch. Marker fonts explicitly use Yu Gothic.

`source/` contains generated DOCX inputs; `fixtures/doc/regressions/list-markers/SexagenaryMarkers/` contains exported DOC files; `projected/` contains their DOCX projections. `word-evidence.json` records relative paths, SHA-256 hashes, and editable list strings for all 27 documents. Complete independently extracted Word PDF text is checked against every marker and paragraph. Word captures were made read-only on 2026-11-13.

`baseline-errors.json` records nine unsupported-format export failures on main. No baseline DOC artifact was produced. Tests additionally verify binary codes, projected format declarations/schema, source immutability, and complete output through five text visitors. No real document content or PDF/layout implementation changes are included.
