# List paragraphs in story compatibility fallbacks

Eighteen generic documents contain Item0-Item2 paragraphs with [7], [8], [9] markers at level zero/eight in the body, primary header or primary footer. Each combination has direct content, a compatibility fallback, and nested fallbacks. Unknown Choice namespaces select Fallback in Word and in the current DOC writer. Rejected branches contain paragraphs referencing a hexadecimal list intentionally unsupported by DOC export. No ordinary style references that rejected list.

On main 89af5680, twelve cases fail DOC export while discovering lists in rejected paragraphs; six direct controls pass. Text export also omits the selected paragraphs because the block dispatcher lacks an AlternateContent case. source-text-baseline.json records twelve empty CLI text results and six correct controls using the previously built cd63017e CLI; its library/CLI source matches the main base. word-baseline.json records the twelve pre-fix DOC export exceptions against input hashes.

The block dispatcher now uses the existing alternate-content dispatcher. DOC list discovery checks ancestor branch selection before resolving a paragraph, using the DOC visitor's choice policy and caching selected branches. This covers nested choices without importing rejected list references. The fixtures exercise unsupported Choice plus Fallback, not new namespace support. Separate reader guards cover accepting/rejecting a Choice and skipping the opposite branch; three export guards preserve failures for selected unsupported lists.

Tests check immutable source document/styles/numbering/header/footer XML, projected schema validity, and markers through plain/rich Markdown and HTML. PlainText checks body markers and intentional header/footer suppression.

Microsoft Word opened every source, exported DOC and projected DOCX read-only. word-evidence.json records 54 exact ListString sequences and input/PDF hashes. Complete PDF text and embedded marker glyph outlines/sizes match across each triple. word-shape-evidence.json records relative glyph-origin comparisons within 0.02 pt; these checks do not assert whole-page layout equivalence.

DOC fixtures are grouped under fixtures/doc/regressions/list-markers/MarkerStoryChoiceDiscovery; Word PDFs under truth/pdf-word/regressions/list-markers/DocMarkerStoryChoiceDiscovery. All fixtures are synthetic. No PDF visitor or layout implementation changes are included.
