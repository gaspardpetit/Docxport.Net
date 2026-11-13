# Enclosed-number markers through DOC

All fixtures contain only synthetic `Item0` through `Item5` paragraphs.

DecimalEnclosedFullstop, DecimalEnclosedParen, DecimalEnclosedCircleChinese, and IdeographEnclosedCircle already render in text exports. DOC export rejected them, and DOCX projection rejected their binary codes. This change adds all four mappings in both directions, without changing the numbering renderer. Existing DecimalEnclosedCircle is included as a control.

Microsoft [MS-OSHARED 2.2.1.3 MSONFC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659) defines these four formats as 0x1A, 0x1B, 0x1C, and 0x1D, respectively. The existing circled-decimal control uses 0x12.

Fifteen cases cover each format at depth 0 with starts 1 and one before its enclosed-glyph limit, and depth 8 starting one before that limit. Every case has six successive markers. Limits are 20 for full-stop, parenthesized, and ordinary circled decimal markers, and 10 for Chinese circled and circled ideograph markers. Boundary cases confirm Word switches to plain decimal text afterward. Marker fonts explicitly use Yu Gothic.

`source/` contains generated DOCX inputs; `fixtures/doc/regressions/list-markers/EnclosedMarkers/` contains exported DOC files; `projected/` contains their DOCX projections. `word-evidence.json` records relative paths, SHA-256 hashes, and editable list strings for all 45 documents. Complete independently extracted Word PDF text is checked against every marker and paragraph. Word captures were made read-only on 2026-11-13.

`baseline-errors.json` records 12 unsupported-format export failures and three passing controls on main. No DOC artifact was produced for unsupported formats. Tests additionally verify binary codes, projected format declarations/schema, source immutability, and complete output through five text visitors. No real document content or PDF/layout implementation changes are included.
