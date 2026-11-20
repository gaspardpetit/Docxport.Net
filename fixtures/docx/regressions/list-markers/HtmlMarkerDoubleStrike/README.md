# Explicit double-strike HTML list markers

The 69 synthetic DOCX sources in `source/` have matching Word-exported PDFs in
`truth/pdf-word/regressions/list-markers/HtmlMarkerDoubleStrike/source/`.
`word-evidence.json` binds each source/PDF pair by SHA-256 and records Word list
strings, PDF text, and separate strike/underline rectangles. Double strike has
two strokes per marker, single strike has one, and disabled cases have none.
The body text is undecorated throughout.

The matrix covers decimal, bullet, and literal (`none`) labels; absent, single,
and explicitly disabled underline; disabled, single, double, combined single/double,
and explicitly disabled double strike. Additional cases cover implicit enabled
double strike, RTL paragraphs, tables, and compatibility Choice/Fallback selection.
Word displays two strike lines when single and double strike are both enabled.

HTML represents double strike on a dedicated inner span. Underline stays on the
outer marker so that a single underline remains single. Single-strike and disabled
cases keep their existing markup. This applies to rich and plain HTML, whose
existing marker decoration support is retained. Inherited double strike, underline
styles/colors, body-run decorations, and PDF layout are separate work.

Tests validate all packages after compatibility preprocessing, check labels and
independent decoration spans in rich/plain HTML, and ensure source XML is unchanged.
CLI comparisons cover six export modes, permitting only the double-strike marker
decoration change in HTML; other exports and controls must remain byte-identical.
Four representative rich HTML baseline/current pairs are retained under `baseline/`
and `html/`, with normalized-LF hashes in the manifest.

An additional browser check renders the exported marker components in isolation
with Arial at 12 points. The manifest records the Chromium/Edge version, PDF hashes,
and strike/underline rectangles. All 69 current renders match Word's stroke counts;
33 baseline renders have only one strike instead of two. These component renders
verify decoration behavior, not whole-document layout or exact stroke geometry.
