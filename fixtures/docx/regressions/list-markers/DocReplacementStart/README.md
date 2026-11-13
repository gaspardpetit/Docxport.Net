# Replacement start precedence through DOC

Synthetic paragraphs contain only `Item0` through `Item5`. No real document content is included.

When a numbering instance has both `w:startOverride` and a replacement `w:lvl/w:start`, Word uses the explicit replacement start. DOC has only one effective start when both LFOLVL flags are set. Import now resolves that effective value before serialization, preventing the writer from replacing it with the outer override.

The 18 cases cover depths 0, 1, and 8, both conflicting directions (outer 2 / nested 7 and outer 7 / nested 2), one used instance versus two instances sharing an abstract counter, and omitted nested starts. Omitting the nested start still uses the outer override. Upper Roman replacements distinguish formatting from shared counter progress.

`source/` contains generated DOCX inputs, `fixtures/doc/regressions/list-markers/ReplacementStart/` contains exported DOC files, and `projected/` contains their DOCX projections. `word-evidence.json` records SHA-256 hashes and six editable Word list strings for all 54 documents. The corresponding Word PDF text was independently checked against all labels and paragraph text.

`before/` and `before-word-evidence.json` preserve three projections from main before the fix. Word displays the incorrect exported starts there. Word captures were made read-only on 2026-11-13. Source numbering XML remains unchanged by import. Tests additionally check binary effective starts, retained replacement formatting, shared ownership, projected OOXML validation, and five text visitors.
