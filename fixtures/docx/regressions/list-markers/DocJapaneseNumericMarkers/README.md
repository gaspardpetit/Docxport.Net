# Japanese numeric markers through DOC

All fixtures contain only synthetic `Item0` through `Item5` paragraphs.

JapaneseCounting, JapaneseLegal, and JapaneseDigitalTenThousand already render in text exports. DOC export rejected them, and DOCX projection rejected their binary codes. This change adds both directions of the mappings, leaving the numbering renderer unchanged.

Microsoft [MS-OSHARED 2.2.1.3 MSONFC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659) defines these formats as 0x0B, 0x10, and 0x11, respectively.

Twelve cases cover each format at depth 0 with starts 1, 98, and 9998, and depth 8 starting at 9998. Each has six successive markers. These distinguish unit counting, legal numerals, digit notation, and the transition through 10,000. Word displays an empty numeric expansion at 10,000 and above for JapaneseDigitalTenThousand, preserving the BEGIN/END literals. Expected strings are literal editable ListString captures from Word. Marker fonts explicitly use Yu Gothic.

`source/` contains generated DOCX inputs; `fixtures/doc/regressions/list-markers/JapaneseNumericMarkers/` contains exported DOC files; `projected/` contains their DOCX projections. `word-evidence.json` records relative paths, SHA-256 hashes, and exact editable list strings for all 36 documents. Complete independently extracted Word PDF text is checked against every marker and paragraph, normalizing whitespace only. Word captures were made read-only on 2026-11-15.

`baseline-errors.json` records twelve unsupported-format export failures on main. No baseline DOC artifact was produced. Tests additionally verify binary codes, projected format declarations/schema, source immutability, and complete output through five text visitors. No real document content or PDF/layout implementation changes are included.
