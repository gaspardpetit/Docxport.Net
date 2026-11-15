# Hindi list markers through DOC

All fixtures contain only synthetic `Item0` through `Item5` paragraphs.

HindiVowels, HindiConsonants, HindiNumbers, and HindiCounting already render in text exports. DOC export rejected them, and DOCX projection rejected their binary codes. This change adds both directions of the mappings, leaving the numbering renderer unchanged.

Microsoft [MS-OSHARED 2.2.1.3 MSONFC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659) defines these formats as 0x31 through 0x34, respectively.

Sixteen cases each contain six successive markers:

- HindiVowels: depth 0 starts 1, 36, and 303; depth 8 starts 303.
- HindiConsonants: depth 0 starts 1, 17, and 303; depth 8 starts 303.
- HindiNumbers and HindiCounting: depth 0 starts 1, 98, and 9998; depth 8 starts 9998.

These cases cover alphabet repetition and the 304-position cycle, decimal digit boundaries, and HindiCounting's wrap at 10,000. Word uses consonants for HindiVowels and vowels for HindiConsonants. Expected strings are literal editable ListString captures from Word. Marker fonts explicitly use Nirmala UI.

`source/` contains generated DOCX inputs; `fixtures/doc/regressions/list-markers/IndicMarkers/` contains exported DOC files; `projected/` contains their DOCX projections. `word-evidence.json` records relative paths, SHA-256 hashes, and exact editable list strings for all 48 documents. Complete independently extracted Word PDF text is checked against every marker and paragraph, normalizing whitespace only. Word captures were made read-only on 2026-11-15.

`baseline-errors.json` records sixteen unsupported-format export failures on main. No baseline DOC artifact was produced. Tests additionally verify binary codes, projected format declarations/schema, source immutability, and complete output through five text visitors. No real document content or PDF/layout implementation changes are included.
