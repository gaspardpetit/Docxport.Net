# Direct paragraph numbering compatibility

Fourteen documents cover direct numbering ID, level, paragraph-style and paragraph-property wrappers/choices, fallback selection, inherited list levels, ignored IDs, namespace bindings declared on ancestors, and rebound prefixes. The disable-wrapper control must suppress all three markers while preserving the paragraph text.

Fresh Microsoft Word ListString/PDF captures on 2026-11-12 confirm 39 visible labels and three suppressed markers. See `word-evidence.json` for exact values and paired source/PDF hashes. Tests also compare the same inputs with Open XML SDK compatibility preprocessing and verify unchanged source document XML across plain/rich text visitors.

Compatibility is normalized only for marker membership resolution. General paragraph layout/style resolution, style-definition compatibility and DOC round trips remain separate scopes.
