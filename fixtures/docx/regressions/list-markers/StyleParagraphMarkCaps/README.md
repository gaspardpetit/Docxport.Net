# Style paragraph-mark capitalization

These 52 synthetic documents cover Word's handling of `w:style/w:pPr/w:rPr/w:caps` for list markers. Each has three literal Latin or Greek labels and generic body text with direct caps disabled.

Word accepts this location although it is outside the OOXML schema. Tests assert the specific styles-part validation errors instead of treating these documents as conformant OOXML. Visible output from fresh Word PDF exports is the oracle; `ListString` alone does not report capitalization formatting.

The matrix covers default and named paragraph styles, document defaults, based-on chains, mark-only declarations, coexisting style-run and style-mark declarations, and direct paragraph-mark and numbering-level overrides. A mark-only declaration supplies the nearest style value. When both locations coexist, Word inherits the mark declaration independently as a toggle. This rule is confined to marker capitalization.

`word-evidence.json` records expected visible labels, source/PDF hashes, Word list strings and PDF text, and actual CLI before/after hashes for rich HTML, plain HTML and plain Markdown. Representative changed rich HTML pairs are retained in `baseline/` and `html/`. Matching Word PDFs are in `truth/pdf-word/regressions/list-markers/StyleParagraphMarkCaps/source/`.

The tests exercise all four text visitors and assert that export preserves the input document, styles and numbering XML. Existing schema-valid capitalization fixtures remain in `StyleMarkerCaps` and `ParagraphMarkerCaps`.
