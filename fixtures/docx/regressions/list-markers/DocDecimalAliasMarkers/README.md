# Alternate decimal DOC numbering code

All fixtures contain only synthetic `Item0` through `Item5` paragraphs.

Microsoft [MS-OSHARED 2.2.1.3 MSONFC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659) maps both 0x00 and 0x28 to decimal. DOCX projection recognized only 0x00 and rejected 0x28. The fix adds the missing input mapping. Export continues to use canonical code 0x00.

Eight cases cover depth 0 and depth 8, with 0x28 in either the base definition or a formatting-only override. Each has six consecutive positions. Base-definition cases begin at 1 or 98, including the 99/100 boundary. Override cases declare a replacement start of 1 or 98, but Word keeps the abstract level's start of 2 because the override has no start flag. Literal expectations preserve this existing behavior.

The test source factory first exports a canonical decimal DOC, then changes only the selected LVLF.nfc byte to 0x28 using indexed structure offsets. It performs no byte-pattern searching. The DOC reader confirms that the alternate code is present in the selected definition or override. The projected DOCX must validate, use decimal declarations, and preserve the sequence through five text visitors. Re-export must use code 0x00 for every definition and override.

`source/` contains canonical decimal DOCX controls. `fixtures/doc/regressions/list-markers/DecimalAliasMarkers/` contains the synthetic DOC inputs carrying 0x28; `projected/` contains their DOCX projections. `word-evidence.json` records relative paths, SHA-256 hashes, and exact editable list strings for all 24 documents. Complete independently extracted Word PDF text is checked against every marker and paragraph, normalizing whitespace only. Word opened every input read-only for captures on 2026-11-15.

`baseline-errors.json` records eight unsupported-format projection failures on main. Word independently opened and displayed all eight baseline DOC inputs correctly. Source DOCX and baseline DOC marker sequences agree. No real document content or PDF/layout implementation changes are included.
