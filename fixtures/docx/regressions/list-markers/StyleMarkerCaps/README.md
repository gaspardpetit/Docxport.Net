# Style and default capitalization of list labels

These 32 synthetic documents contain three `Item0`–`Item2` paragraphs and either a Latin `ab` or Greek `αβ` label. Each script covers 16 settings: document defaults on/off, default paragraph style on/off, named style on/off, base style on with child absent/off/on, defaults/style precedence in both directions, style/paragraph-mark precedence in both directions, style/numbering-level precedence in both directions, and an unused style with caps enabled.

Word uses the nearest style run-properties declaration for marker capitalization; in particular, enabling caps in both a base and child style leaves the label uppercase. Document run defaults are the fallback. Explicit paragraph-mark and numbering-level declarations retain precedence, including off. Each body run explicitly disables caps so the marker rule can be checked independently of body text.

All `source/` inputs validate against the Office 2013 schema. Word PDF captures are under `truth/pdf-word/regressions/list-markers/StyleMarkerCaps/source/`. `word-evidence.json` records source/PDF hashes, visible expected labels, Word ListString values, PDF text, and CLI comparison hashes. ListString exposes unformatted labels; expectations follow the visible PDF text, whose complete character multiset was independently verified against labels and body text.

Main at `df9213c4` fails 12 cases and passes 20 controls. `baseline/` and `html/` retain the rich HTML before/after pairs for the 12 corrections. Across actual CLI rich HTML, plain HTML, and plain Markdown exports, 36 outputs change only label capitalization and all 60 controls are byte-identical. Output hashes normalize CRLF to LF for portable Git blob verification.

`MarkerStyleCapsTests` checks four visitors (rich HTML, plain HTML, plain Markdown and plain text), schema validity, expected labels and unchanged document/styles/numbering XML. Word-compatible paragraph-mark properties nested inside style paragraph properties are outside this increment.
