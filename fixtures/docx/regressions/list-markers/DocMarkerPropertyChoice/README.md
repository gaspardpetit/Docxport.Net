# Compatibility-selected paragraph list properties through DOC

Thirty synthetic documents contain three Item0-Item2 paragraphs with [7], [8], [9] markers. They cover levels zero/eight and compatibility choices wrapping paragraph properties, numbering properties, numId, ilvl or pStyle. Each scope has a supported Choice, an unsupported Choice with Fallback, and a direct-property control. Rejected branches point to a different style, list or level with visibly different markers.

On main 3443bab3, twelve cases wrapping pPr, numPr or pStyle lose their markers on DOC export and projection. The eighteen direct/numId/ilvl controls pass. DOC export now reuses the existing paragraph numbering compatibility selector for style and list membership. Other paragraph formatting continues through its existing path; no PDF visitor or layout implementation changes are included.

Microsoft Word opened each source, exported DOC and projected DOCX read-only. word-evidence.json records 90 exact ListString sequences and hashes of each input and Word PDF. Independent checks compare complete PDF text, embedded marker glyph outlines/sizes, and relative glyph origins within 0.02 pt; word-shape-evidence.json records the comparisons. These checks do not compare whole-page layout.

Twelve generic pre-fix DOC/PDF pairs retain the missing markers and are indexed by word-baseline.json. DOCs are grouped under fixtures/doc/regressions/list-markers/MarkerPropertyChoice and PDFs under truth/pdf-word/regressions/list-markers/DocMarkerPropertyChoice. Source and projected DOCX remain grouped here.
