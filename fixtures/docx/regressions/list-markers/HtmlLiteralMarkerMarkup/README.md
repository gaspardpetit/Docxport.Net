# Literal marker text and generated font wrappers

These 29 generic documents cover rich/plain HTML labels containing angle brackets, tag-like text, entity-like text, ampersands and Unicode, plus numeric substitutions and caps. Each pattern is exercised as a numbered label, a bullet without a font wrapper, and an Arial bullet with the formatter's generated wrapper. Ordinary labels and genuine generated wrappers provide controls.

Word displays the label characters literally. Previously, HTML guessed that any angle-bracket pair was markup, while span-like content could also be stripped during normalization. The fix recognizes a generated bullet-font wrapper only when it matches the resolved numbering format and font metadata. Its payload remains literal text and is escaped before rendering; plain output extracts that payload without stripping tag-like label characters. Marker size classification also reads the literal payload.

All visible expectations were confirmed using fresh Word PDFs, grouped under `truth/pdf-word/regressions/list-markers/HtmlLiteralMarkerMarkup/source/`. Source/PDF hashes, Word strings/text, exact expected labels and actual CLI before/after hashes are in `word-evidence.json`. Four representative rich HTML pairs are retained in `baseline/` and `html/`.

HTML tests assert exact rich/plain HTML marker text, preservation of the intended generated font span, absence of unintended label elements, OOXML validation and source document/styles/numbering immutability. CLI checks compare changes inside marker payloads, preserve marker CSS and the surrounding document, and check existing generic sample exports. The HTML fix preserves the formatter output contract.


Plain-text coverage reuses these same 29 source documents and matching Word PDFs. `PlainTextLiteralMarkerTests` checks accept/reject exports, exact labels and source XML preservation, plus established Symbol and Wingdings translation controls. The plain-text visitor removes only the metadata-confirmed generated wrapper; it preserves tag-like characters in its literal payload.

`plain-text-evidence.json` binds its CLI results to the unchanged Word evidence manifest and records before/after hashes for accept/reject text, rich/plain HTML and rich/plain Markdown. Four representative text pairs are retained under `plain-text/baseline/` and `plain-text/after/`. Source documents and PDFs are shared rather than duplicated for this follow-up.


Markdown coverage reuses the 29 original sources and adds eight context variants under `markdown/source/`: four label patterns each as marker-only paragraphs and as paragraphs with spacing that selects the rich HTML paragraph wrapper. Their fresh Word PDFs are under the corresponding `truth/pdf-word/.../HtmlLiteralMarkerMarkup/markdown/source/` path; `markdown-word-evidence.json` records source/PDF hashes and Word-visible text.

`MarkdownLiteralMarkerTests` checks exact labels, source XML preservation, ordinary/standalone rendering, ordered-list delimiter escaping and Symbol/Wingdings controls. The marker normalization recognizes only metadata-confirmed font wrappers and encodes literal HTML characters. General Markdown punctuation handling remains separate.

`markdown-evidence.json` records actual before/after CLI hashes and markdown-it-py CommonMark rendering checks with HTML enabled. Plain Markdown is also checked with HTML disabled. All 74 rich/plain renders match the 37 Word-confirmed inputs. Four representative plain Markdown pairs are under `markdown/baseline/` and `markdown/after/`. Their paragraph-ending blank lines are preserved as exact CLI evidence.
