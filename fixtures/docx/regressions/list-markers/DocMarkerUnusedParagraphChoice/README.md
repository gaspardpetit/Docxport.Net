# Ignore rejected paragraph list references during DOC export

Fifty-four synthetic documents contain three Item0-Item2 paragraphs with [7], [8], [9] markers in the body, primary header or primary footer. Cases cover levels zero/eight and compatibility choices around paragraph properties, numbering properties or numId. Each scope has a supported Choice, an unsupported Choice selecting Fallback, and a direct-property control. A rejected branch references a separate hexadecimal list, intentionally unsupported by the binary DOC exporter. No ordinary style references that rejected list.

On main b4a2206a, twenty-four cases throw NotSupportedException while importing the rejected branch's list; thirty controls export successfully. DOC list discovery now reads the selected paragraph numbering properties using the existing normalization helper, alongside the effective style/direct numbering pair. It continues collecting normalized style definitions. Three additional guards verify that actively selected unsupported lists still fail in all three stories.

Tests check immutable source document/styles/numbering/header/footer XML, projected DOCX schema validity, and exact markers through plain/rich Markdown and HTML. PlainText verifies body markers and its intentional suppression of header/footer content.

Microsoft Word opened each source, exported DOC and projected DOCX read-only. word-evidence.json records 162 exact ListString sequences and input/PDF hashes. Complete PDF text and embedded marker glyph outlines/sizes match across each triple. word-shape-evidence.json records relative glyph-origin comparisons within 0.02 pt; these checks do not assert whole-page layout equivalence.

word-baseline.json records twenty-four pre-fix export exceptions against the same source inputs; those cases produced no baseline DOC. Binary fixtures are grouped under fixtures/doc/regressions/list-markers/MarkerUnusedParagraphChoice, Word PDFs under truth/pdf-word/regressions/list-markers/DocMarkerUnusedParagraphChoice, and source/projected DOCX here. All fixtures are synthetic. No PDF visitor or layout implementation changes are included.
