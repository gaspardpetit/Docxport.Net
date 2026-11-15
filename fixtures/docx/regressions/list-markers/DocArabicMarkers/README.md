# Arabic alphabetic markers through DOC

All fixtures contain only synthetic `Item0` through `Item5` paragraphs.

ArabicAlpha and ArabicAbjad already render in text exports. DOC export rejected them, and DOCX projection rejected their binary codes. This change adds these mappings in both directions, leaving the numbering renderer unchanged.

Microsoft [MS-OSHARED 2.2.1.3 MSONFC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659) defines these formats as 0x2E and 0x30, respectively.

Ten cases cover each format at depth 0 with starts 1, 27, 55, and 391, and depth 8 starting at 391. Each has six successive markers. These cases distinguish the two alphabet orders, repetition across alphabet boundaries, and the 392-item cycle. Expected strings are literal editable ListString captures from Word, including U+200C joining controls before or after each letter according to the format. Marker fonts explicitly use Arial.

`source/` contains generated DOCX inputs; `fixtures/doc/regressions/list-markers/ArabicMarkers/` contains exported DOC files; `projected/` contains their DOCX projections. `word-evidence.json` records relative paths, SHA-256 hashes, and editable list strings for all 30 documents. Complete independently extracted Word PDF text is checked against every visible marker and paragraph. Word PDF text omits the nonprinting U+200C characters, so `pdfLabels` exclude them; editable list strings and text visitors retain exact checks. Word captures were made read-only on 2026-11-15.

`baseline-errors.json` records ten unsupported-format export failures on main. No baseline DOC artifact was produced. Tests additionally verify binary codes, projected format declarations/schema, source immutability, and complete output through five text visitors. No real document content or PDF/layout implementation changes are included.
