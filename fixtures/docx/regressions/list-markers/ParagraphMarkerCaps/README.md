# Paragraph-mark capitalization of list labels

These 18 synthetic documents contain only three `Item0`–`Item2` paragraphs and a Latin `ab` or Greek `αβ` list label. Each script covers absent, enabled, and disabled `w:caps` on both paragraph-mark run properties and numbering-level run properties. Styles and document defaults do not declare capitalization.

Word renders paragraph-mark capitalization when the level has no capitalization declaration. An explicit level declaration wins, including `w:caps w:val="0"`. Body text remains unchanged. This includes levels with no run properties at all.

`source/` holds the schema-valid DOCX inputs. Corresponding Word PDF captures are under `truth/pdf-word/regressions/list-markers/ParagraphMarkerCaps/source/`. `word-evidence.json` records source/PDF hashes, expected visible labels, Word ListString results, PDF text, and CLI output hashes. Word's ListString API reports lowercase unformatted labels even when the PDF renders uppercase; visible expectations follow the PDF.

The two cases `latin-mark-on-level-absent` and `greek-mark-on-level-absent` fail on main at `76b9f830`; the other 16 cases are controls. `baseline/` and `html/` retain rich HTML before and after those two corrections. Their only output change is label capitalization. CLI plain HTML and Markdown also receive the corrected labels. All 48 CLI control outputs are byte-identical to main. HTML hashes normalize CRLF to LF for portable Git blob verification.

`MarkerParagraphCapsTests` validates each source with the Office 2013 schema, checks rich HTML, plain HTML, plain Markdown and plain text, and verifies export leaves document/styles/numbering XML unchanged. Style/default capitalization inheritance is outside this increment.
