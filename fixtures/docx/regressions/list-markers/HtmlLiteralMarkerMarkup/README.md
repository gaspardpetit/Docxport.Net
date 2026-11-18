# Literal marker text and HTML font wrappers

These 29 generic documents cover rich/plain HTML labels containing angle brackets, tag-like text, entity-like text, ampersands and Unicode, plus numeric substitutions and caps. Each pattern is exercised as a numbered label, a bullet without a font wrapper, and an Arial bullet with the formatter's generated wrapper. Ordinary labels and genuine generated wrappers provide controls.

Word displays the label characters literally. Previously, HTML guessed that any angle-bracket pair was markup, while span-like content could also be stripped during normalization. The fix recognizes a generated bullet-font wrapper only when it matches the resolved numbering format and font metadata. Its payload remains literal text and is escaped before rendering; plain output extracts that payload without stripping tag-like label characters. Marker size classification also reads the literal payload.

All visible expectations were confirmed using fresh Word PDFs, grouped under `truth/pdf-word/regressions/list-markers/HtmlLiteralMarkerMarkup/source/`. Source/PDF hashes, Word strings/text, exact expected labels and actual CLI before/after hashes are in `word-evidence.json`. Four representative rich HTML pairs are retained in `baseline/` and `html/`.

Tests assert exact rich/plain HTML marker text, preservation of the intended generated font span, absence of unintended label elements, OOXML validation and source document/styles/numbering immutability. CLI checks compare changes inside marker payloads, preserve marker CSS and the surrounding document, and check existing generic sample exports. This change is scoped to HTML rendering; formatter output and other visitors retain their existing contracts.
