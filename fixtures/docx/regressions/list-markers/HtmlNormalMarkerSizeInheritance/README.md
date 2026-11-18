# Normal HTML marker size inheritance

These 45 synthetic documents capture Word's visible list-marker sizes. Sources
are grouped in `source/`; matching Word-exported PDFs are under
`truth/pdf-word/regressions/list-markers/HtmlNormalMarkerSizeInheritance/source/`.
`word-evidence.json` binds each source and PDF by SHA-256 and records Word list
strings, extracted PDF text, and marker font sizes. Small PDF rounding differences
(for example, 27.96 instead of 28 points) are expected.

The 37 target/control cases cover direct paragraph-mark sizes, paragraph styles,
default paragraph styles, based-on styles, document defaults, absent/empty/font-only
level properties, explicit level sizes, table paragraphs, and markup-compatibility
Choice/Fallback selection. Body-run sizes deliberately differ from marker sizes.
Five cases cover Word-compatible style paragraph-mark properties: within one style,
its run size wins over its paragraph-mark size; the nearest style wins over its base.
The existing numeric RTL inheritance path has a control case.

Eight guard cases retain existing behavior for inherited complex-script/RTL flags,
font hints, superscript, and RTL paragraphs. This change does not resolve those
contexts or non-ASCII marker sizing. It only adds normal LTR ASCII inheritance
where script selection and glyph scaling are unambiguous.

The tests validate packages after compatibility preprocessing and check rich/plain
HTML labels, marker CSS, and source XML immutability. Schema exceptions are limited
to the exact style paragraph-mark and complex-script elements that Word accepts.
Other fixtures must validate without errors.

CLI comparisons cover rich/plain HTML, rich/plain Markdown, and accepted/rejected
plain text. Only rich HTML marker font sizes change; the other exports and guarded
cases remain byte-identical. Representative baseline/current rich HTML captures
are retained in `baseline/` and `html/`, with normalized-LF hashes in the manifest.
The adjacent `HtmlMarkerLevelSize` missing-normal/missing-both cases now inherit
the 12-point document default; their comparison records are included here.
