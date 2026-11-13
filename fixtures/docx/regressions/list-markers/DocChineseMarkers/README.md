# Chinese counting and legal markers through DOC

All fixtures contain only synthetic `Item0` through `Item5` paragraphs.

ChineseCounting, ChineseLegalSimplified, and ChineseCountingThousand already render in text exports. DOC export rejected them, and DOCX projection rejected their binary codes. This change adds these mappings in both directions, leaving the numbering renderer unchanged.

Microsoft [MS-OSHARED 2.2.1.3 MSONFC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659) defines these formats as 0x25, 0x26, and 0x27, respectively.

Twelve cases cover each format at depth 0 with starts 1, 98, and 9998, and depth 8 starting at 9998. Each has six successive markers. These cases distinguish counting below 100 from digit substitution, simplified legal numerals, counting through ten thousand, and zero characters. Expected strings are literal editable ListString captures from Word, rather than a second implementation of the formatter. Marker fonts explicitly use Yu Gothic.

`source/` contains generated DOCX inputs; `fixtures/doc/regressions/list-markers/ChineseMarkers/` contains exported DOC files; `projected/` contains their DOCX projections. `word-evidence.json` records relative paths, SHA-256 hashes, and editable list strings for all 36 documents. Complete independently extracted Word PDF text is checked against every marker and paragraph. Word captures were made read-only on 2026-11-13.

`baseline-errors.json` records twelve unsupported-format export failures on main. No baseline DOC artifact was produced. Tests additionally verify binary codes, projected format declarations/schema, source immutability, and complete output through five text visitors. No real document content or PDF/layout implementation changes are included.
