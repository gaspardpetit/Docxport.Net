# Omitted non-root replacement starts through DOC

Every fixture contains only synthetic `Item0` through `Item5` paragraphs.

Word uses zero for a non-root replacement level whose `w:start` is omitted, once a parent establishes the level's start context. DOC import previously supplied one. This changed the sequence `3, 0, 1, 4, 0, 1` to `3, 1, 2, 4, 1, 2` after export.

The fix changes only the omitted non-root replacement fallback. Root activation retains its existing fallback; explicit starts and startOverride precedence retain their paths. A child list beginning before its parent still starts from the abstract definition, as Word displays.

Fifteen cases cover depths 0, 1, and 8; omitted, explicit zero, and explicit seven replacement starts; child-first lists and parent/child/restart sequences. Tests assert binary starts, source immutability, projected schema, and complete marker/body output through five text visitors. All 96 retained audit start/restart cases also pass with this change; only the focused generic fixtures are submitted.

`source/` contains the generated DOCX inputs, `fixtures/doc/regressions/list-markers/ReplacementDefault/` contains exported DOC files, and `projected/` contains their DOCX projections. `word-evidence.json` records relative paths, SHA-256 hashes, and six editable Word list strings per document. Independently extracted complete Word PDF text is checked against all labels and paragraph text. Word captures were made read-only on 2026-11-13.

Two pre-fix projections and Word PDFs are preserved in `before/`, with their incorrect labels recorded in `before-word-evidence.json`. No real document content or PDF/layout implementation changes are included.
