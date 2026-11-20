# Inherited double-strike HTML list markers

The 117 synthetic sources in `source/` have matching Word-exported PDFs under
`truth/pdf-word/regressions/list-markers/HtmlMarkerDoubleStrikeInheritance/source/`.
The manifest binds each pair by SHA-256 and records Word list strings, PDF text,
and marker strike rectangles. Body runs explicitly disable the relevant strike
flags so the captures distinguish marker inheritance from body formatting.

Coverage includes decimal, bullet, and literal labels; paragraph marks, styles,
based-on/default paragraph styles, and document defaults; absent/empty/font-only
level properties; explicit-off and enabled overrides; Word-compatible style
paragraph marks; compatibility Choice/Fallback; RTL paragraphs; and tables.
The mixed-flag cases check that single strike and double strike inherit independently:
a nearer single-strike declaration does not clear inherited double strike, and
explicitly disabling double strike does not clear inherited single strike.

Unlike single strike, Word does not apply a separate style paragraph-mark toggle
for double strike. Within each style its run declaration takes precedence over
its paragraph mark; the nearest style takes precedence over its base. Direct
level and paragraph-mark declarations override the inherited result, including false.

Tests check rich/plain HTML decorations, labels, source XML immutability, and schema
validation after compatibility preprocessing. The only accepted schema errors are
the exact style `pPr/rPr` locations that Word supports. The exporter uses the same
resolved double-strike value for the inner decoration span and outer marker CSS.

CLI comparisons permit only double-strike marker decoration changes in HTML;
controls and other export formats must stay byte-identical. Representative rich
HTML baseline/current pairs are retained under `baseline/` and `html/`, with LF
hashes in the manifest. Browser component renders in Arial 12pt additionally compare
strike counts with Word; their version, PDF hashes, and rectangles are recorded.
These checks concern marker decoration, not exact geometry or whole-page layout.
Underline styles/colors, body-run formatting, and PDF layout are separate work.
