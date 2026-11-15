# Hebrew markers through DOC

All fixtures contain only synthetic `Item0` through `Item5` paragraphs.

Hebrew1 and Hebrew2 already render in text exports. DOC export rejected them, and DOCX projection rejected their binary codes. This change adds these mappings in both directions, leaving the numbering renderer unchanged.

Microsoft [MS-OSHARED 2.2.1.3 MSONFC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659) defines these formats as 0x2D and 0x2F, respectively.

Ten cases cover each format at depth 0 with starts 1, 14, 21, and 391, and depth 8 starting at 391. Each has six successive markers. These cases cover the special numeric forms for 15 and 16, alphabet rollover, and the 392-item cycle. Expected strings are literal editable ListString captures from Word, including the U+200F right-to-left mark in Hebrew2. Marker fonts explicitly use Arial.

`source/` contains generated DOCX inputs; `fixtures/doc/regressions/list-markers/HebrewMarkers/` contains exported DOC files; `projected/` contains their DOCX projections. `word-evidence.json` records relative paths, SHA-256 hashes, and editable list strings for all 30 documents. Complete independently extracted Word PDF text is checked against every visible marker and paragraph. Word PDF text omits the nonprinting U+200F character, so `pdfLabels` exclude it; editable list strings and text visitors retain exact checks. Word captures were made read-only on 2026-11-14.

`baseline-errors.json` records ten unsupported-format export failures. No baseline DOC artifact was produced. Tests additionally verify binary codes, projected format declarations/schema, source immutability, and complete output through five text visitors. No real document content or PDF/layout implementation changes are included.
