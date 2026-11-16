# Compatibility-selected style numbering through DOC

Thirty-six synthetic documents contain three Item0-Item2 paragraphs with [7], [8], [9] markers inherited from a paragraph style. They cover levels zero/eight and compatibility choices wrapping a complete style, its paragraph properties, numbering properties, numId or ilvl. A separate scope selects a custom default paragraph style for paragraphs without pStyle. Each scope has a supported Choice, an unsupported Choice with Fallback, and a direct-property control. Rejected branches supply a different list or level with visibly different markers.

On main 9ad25b1e, twenty-two cases lose their numbering or use the wrong level during DOC export and projection; fourteen controls pass. DOC export now enumerates the existing compatibility-normalized style view instead of raw styles. This preserves selected style properties, including default style selection. No PDF visitor or layout implementation changes are included.

Microsoft Word opened each source, exported DOC and projected DOCX read-only. word-evidence.json records 108 exact ListString sequences and hashes of each input and Word PDF. Independent checks compare complete PDF text, embedded marker glyph outlines/sizes, and relative glyph origins within 0.02 pt; word-shape-evidence.json records the comparisons. These checks do not compare whole-page layout.

Generic pre-fix DOC/PDF pairs retain the failures and are indexed by word-baseline.json. DOCs are grouped under fixtures/doc/regressions/list-markers/MarkerStyleChoice and PDFs under truth/pdf-word/regressions/list-markers/DocMarkerStyleChoice. Source and projected DOCX remain grouped here.
