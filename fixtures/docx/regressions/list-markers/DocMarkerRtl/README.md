# RTL list marker formatting through DOC

Forty synthetic three-paragraph documents cover Latin, Greek, Hebrew, Arabic, mixed scripts, digits and punctuation, with absent, disabled and enabled RTL/complex-script flags. Marker font slots are Arial, Times New Roman, MS Mincho and Courier New, with ordinary size 12 pt and complex-script size 15 pt. Body text Item0-Item2 deliberately retains explicit RTL and complex-script flags.

On main 830de1f9, Word's DOC marker path applies RTL's forced complex-script formatting to the whole marker, differing from the DOCX path. Ten baseline exports have different marker glyphs or sizes; fourteen automated cases fail slot/flag assertions. The fix is confined to DOC list markers that have an explicit ASCII face and no recognized legacy Symbol face in any font slot: clear marker RTL and any explicit complex-script forcing flag, keep script-specific sizes and faces, and use the ASCII face for the high-ANSI slot. Unicode supplies natural direction. The preceding forced-complex-script export behavior still applies when RTL is disabled. Ordinary body runs and PDF/layout code are unchanged.

Eight additional automated controls cover Symbol/Wingdings in each font slot and retain their original RTL encoding path. Source numbering stays unchanged, body flags survive, and projected documents pass schema validation.

Microsoft Word opened all 120 source/DOC/projected documents read-only on 2026-11-15 (local date). word-evidence.json records exact editable ListString values, paths and SHA-256 hashes. Independent PDF verification checks complete text, embedded marker glyph outlines and sizes, and relative glyph origins within each marker (0.02 pt tolerance). word-shape-evidence.json records these comparisons. This is marker verification, not a claim of matching whole-page layout.

Generic pre-fix DOC/PDF pairs are retained under baseline and indexed by word-baseline.json. DOC files are grouped under fixtures/doc/regressions/list-markers/MarkerRtl; Word PDFs under truth/pdf-word/regressions/list-markers/DocMarkerRtl. Source/projected DOCX remain grouped here.
