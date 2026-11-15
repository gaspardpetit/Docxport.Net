# Thai and Vietnamese list markers through DOC

All fixtures contain only synthetic `Item0` through `Item5` paragraphs.

ThaiLetters, ThaiNumbers, ThaiCounting, and VietnameseCounting already render in text exports. DOC export rejected them, and DOCX projection rejected their binary codes. This change adds both directions of the mappings, leaving the numbering renderer unchanged.

Microsoft [MS-OSHARED 2.2.1.3 MSONFC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659) defines these formats as 0x35 through 0x38, respectively.

Sixteen cases each contain six successive markers:

- ThaiLetters: depth 0 starts 1, 40, and 1229; depth 8 starts 1229.
- ThaiNumbers and ThaiCounting: depth 0 starts 1, 98, and 9998; depth 8 starts 9998.
- VietnameseCounting: depth 0 starts 1, 98, and 998; depth 8 starts 998.

These cases cover repetition across the 41-letter Thai alphabet and the 1230-position cycle, Thai digit and counting boundaries at 100 and 10,000, and Vietnamese counting through 1000 and its subsequent wrap at 1001. Expected strings are literal editable ListString captures from Word, preserving combining-mark sequences. Marker fonts explicitly use Leelawadee UI.

`source/` contains generated DOCX inputs; `fixtures/doc/regressions/list-markers/ThaiVietnameseMarkers/` contains exported DOC files; `projected/` contains their DOCX projections. `word-evidence.json` records relative paths, SHA-256 hashes, and exact editable list strings for all 48 documents. Complete independently extracted Word PDF text is checked against every marker and paragraph, normalizing whitespace only. Word captures were made read-only on 2026-11-15.

`baseline-errors.json` records sixteen unsupported-format export failures on main. No baseline DOC artifact was produced. Tests additionally verify binary codes, projected format declarations/schema, source immutability, and complete output through five text visitors. No real document content or PDF/layout implementation changes are included.
