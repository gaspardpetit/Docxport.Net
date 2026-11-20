# Double-underlined HTML list markers

The 78 synthetic DOCX sources in `source/` have matching Word-exported PDFs under
`truth/pdf-word/regressions/list-markers/HtmlMarkerDoubleUnderline/source/`.
The manifest binds source/PDF pairs by SHA-256 and records Word list strings, PDF
text, and separate strike/underline rectangles. Body text stays undecorated.

The matrix covers decimal, bullet, and literal (`none`) labels; absent, disabled,
single, and double underline; disabled, single, double, and combined strike flags;
RTL paragraphs; tables; compatibility Choice/Fallback; and paragraph-mark, style,
and document-default inherited single/double strike with explicit double underline.

Word displays two underline strokes independently of its one/two strike strokes.
HTML therefore gives double underline its own outer decoration style. When single
strike is also enabled, a dedicated inner span keeps that strike solid and single.
The existing double-strike span already provides the same separation. Single
underline and disabled controls keep their existing markup.

C# tests check rich/plain HTML labels and independent decorations, source XML
immutability, and schema validation after compatibility preprocessing. All sources
must validate without errors. CLI comparisons permit only double-underline marker
changes in HTML; controls and other exporters remain byte-identical. Four rich
HTML baseline/current pairs are retained under `baseline/` and `html/`, with LF
hashes in the manifest.

Browser component renders in Arial 12pt compare both strike and underline counts
with Word; version, PDF hashes, and rectangles are recorded in the manifest. These
checks cover marker decoration, not exact geometry or whole-page layout. Inherited
underline, other underline styles/colors, body-run formatting, and PDF layout are
separate work.
