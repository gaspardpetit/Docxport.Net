# Style paragraph-mark capitalization with MC branches

These 48 generic documents extend `StyleParagraphMarkCaps` with markup-compatibility choices and fallbacks. Word accepts `style/pPr/rPr` despite its being outside the OOXML schema; the SDK represents this run-properties node as an unknown element. The exporter's typed normalization previously skipped alternate content inside it.

The matrix includes Latin and Greek literal markers, style-run caps on/off, selected style-mark caps on/off, supported choices, unsupported choices with fallbacks, wrappers around either caps or the whole rPr, nested alternate content, and direct paragraph-mark/numbering-level overrides. Both uppercase and lowercase results are checked. Whole-rPr wrappers and precedence cases provide controls for already-correct behavior. Body runs have caps explicitly disabled.

All visible expectations were confirmed by fresh Word PDF exports. `ListString` is recorded separately because it omits capitalization formatting. `word-evidence.json` records source/PDF hashes and actual before/after CLI hashes for rich HTML, plain HTML and plain Markdown. Four representative changed rich HTML pairs are retained in `baseline/` and `html/`. Word PDFs are grouped under `truth/pdf-word/regressions/list-markers/StyleParagraphMarkCapsMc/source/`.

Tests exercise four text visitors, preserve the original document/styles/numbering XML, and assert only the known styles-part schema errors after standard MC preprocessing. These are Word compatibility cases, not examples of schema-conformant OOXML.
