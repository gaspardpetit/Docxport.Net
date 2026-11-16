# Direct list marker font hints through DOC

Thirty-two synthetic documents use three Item0-Item2 paragraphs with Greek alpha/beta bullets. They cover levels 0 and 8, abstract definitions and formatting-only overrides, default/eastAsia/cs hints and absent-hint controls, with either explicit marker font names or a hint alone. Fonts are Arial, Times New Roman, MS Mincho and Courier New; document defaults provide the same four faces for hint-only cases.

On main 870f0642, explicit rFonts hints disappear on DOC export and projection. Eight eastAsia cases visibly render in Times New Roman rather than source Word's MS Mincho. Twenty-four automated hint-preservation cases fail; eight absent-hint controls pass.

Preserve the direct hint in the shared DOC character-property codec and project it back to OOXML, including hint-only RunFonts. This requires an additive nullable FontHint property, native reader/writer support, and OOXML mapping. No font-slot selector, PDF visitor or layout code changes. Style/default hint inheritance is a separate follow-up.

The native mapping follows [MS-DOC character properties, sprmCIdctHint](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/7022285b-9621-42e9-ad4d-4e02c115ef18): 0 default, 1 eastAsia, 2 cs, FF no guidance. FF has no ST_Hint equivalent and is omitted from OOXML. Codec controls cover all four operands, invalid values, and clearing direct hints with a plain reset.

Microsoft Word opened the source, exported DOC and projected DOCX documents read-only. word-evidence.json records exact ListString values, relative paths and SHA-256 hashes. Independent PDF verification compares complete text, embedded marker glyph outlines/sizes, and relative glyph origins within 0.02 pt. word-shape-evidence.json records those comparisons. Whole-page layout is outside this check.

Generic pre-fix DOC/PDF pairs are retained under baseline and indexed by word-baseline.json. DOCs are grouped under fixtures/doc/regressions/list-markers/MarkerFontHint and Word PDFs under truth/pdf-word/regressions/list-markers/DocMarkerFontHint. Source/projected DOCX remain grouped here.
