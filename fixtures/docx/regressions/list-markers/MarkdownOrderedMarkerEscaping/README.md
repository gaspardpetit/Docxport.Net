# Ordered markers across Markdown and HTML contexts

Word displays `1.` and `1)` without a backslash. Markdown export previously used backslash escapes to avoid starting an ordered list. Those escapes can become visible when the marker falls inside a raw HTML block, including the first paragraph after a section/body, header/footer, table or footnote opening.

The exporter now uses numeric character references for the generated ordered delimiter: `1&#46;` and `1&#41;`. They preserve literal punctuation in both Markdown text and raw HTML, and do not start an unintended Markdown list. This changes the raw representation in plain Markdown too; already-correct rendered labels remain the same. Literal entity text is escaped separately and is not decoded into a generated delimiter.

The 17 generic inputs cover periods/closing parentheses in body, standalone, rich paragraph-wrapper, table, header, footer and footnote contexts, plus literal entity, font-wrapped bullet and digit-only controls. All inputs validate as OOXML and have fresh Word PDF captures under `truth/pdf-word/regressions/list-markers/MarkdownOrderedMarkerEscaping/source/`. The Word manifest records source/PDF hashes, list strings and independently checked PDF text.

C# tests check four configurations (rich, plain, rich without document colors, rich without paragraph-layout wrappers), exact labels and source-part XML preservation. Existing ordered-marker, empty/restarted marker, inline drawing and revision-boundary tests retain their behavior with the new encoding.

`cli-evidence.json` records actual before/after CLI hashes and markdown-it-py CommonMark rendering with HTML enabled. All 34 rich/plain renders match Word; 20 rendered baseline mismatches are fixed. Plain rendering is also checked with HTML disabled for body/standalone/wrapped inputs. Other exporters are byte-identical. Four representative rich Markdown pairs and their rendered HTML are retained in `baseline/` and `after/`.

Production changes are confined to the two delimiter-encoding returns and their comment. No paragraph boundaries, body formatting, container markup, CSS or PDF/layout code changes are needed.
