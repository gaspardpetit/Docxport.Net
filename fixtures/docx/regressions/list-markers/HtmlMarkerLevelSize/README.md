# Explicit HTML list-marker sizes

These 69 generic three-item documents cover rich HTML sizing of ASCII list markers from the numbering level. Labels, sizes and slot selection were checked against fresh Word PDF exports. All documents validate as OOXML after standard MC preprocessing.

The matrix covers level `sz`/`szCs`, `cs` and `rtl` on/off/absent, paragraph bidi on/off, numeric and Latin mixed labels, common numeric delimiters, level MC wrappers, paragraph/style bidi choices/fallbacks, missing sizes, and body/paragraph-mark cs controls. Word uses the explicit normal size unless level cs is on and rtl is off. In RTL paragraphs numeric labels with common delimiters `()[] .-` use the complex size. Inherited cs does not replace the normal slot when an explicit level size resets it.

The helper is intentionally bounded to printable ASCII and the verified numeric delimiters. Broader inherited sizes and other scripts remain separate. Five cases with omitted selected sizes or non-ASCII markers retain their previous output. Existing pure numeric RTL size inheritance remains supported, including a missing-level-size control. Plain output and Markdown receive no new size styling.

`word-evidence.json` records labels, expected CSS half-points (null means preserve omitted marker-size CSS), source/PDF hashes, Word text and marker trace sizes, and actual CLI before/after hashes. Word PDF rounding can differ by 0.04pt. PDFs are grouped under `truth/pdf-word/regressions/list-markers/HtmlMarkerLevelSize/source/`. Four representative rich HTML changes are retained in `baseline/` and `html/`; hashes cover every CLI pair.

Tests exercise rich/plain HTML, exact marker labels, size-slot selection, XML validation and source document/styles/numbering immutability. Existing inherited numeric RTL tests remain covered. Production changes affect only marker size resolution and rich HTML CSS; they do not change body formatting or PDF layout.
