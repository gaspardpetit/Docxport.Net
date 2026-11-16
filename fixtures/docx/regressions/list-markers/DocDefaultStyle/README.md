# Numbering inherited from the default paragraph style

Fourteen synthetic documents contain three Item0-Item2 paragraphs. Their default paragraph style supplies decimal numbering starting at seven at level zero or eight. Cases cover a default style named Normal, a custom DefaultList ID, a custom default alongside a separate non-default Normal style, explicit paragraph style references, and explicit numId=0 cancellation.

On main c62a4315, paragraphs without pStyle always fall back to Normal or built-in style zero during DOC export. Four custom-default cases therefore lose their [7], [8], [9] markers in both editable DOC and projected DOCX. The ten controls pass. Select the declared default paragraph style when pStyle is omitted; retain the existing fallback when no default is declared. No PDF visitor or layout implementation changes are included.

Microsoft Word opened each source, exported DOC and projected DOCX read-only. word-evidence.json records 42 exact ListString sequences and hashes for each input and Word PDF. Independent checks compare complete PDF text, embedded marker glyph outlines and sizes, and relative glyph origins within 0.02 pt. word-shape-evidence.json records the comparisons; this is not a whole-page layout comparison.

Four generic pre-fix DOC/PDF pairs retain the missing markers and are indexed by word-baseline.json. DOCs are grouped under fixtures/doc/regressions/list-markers/DefaultStyle; Word PDFs under truth/pdf-word/regressions/list-markers/DocDefaultStyle; source and projected DOCX remain grouped here.
