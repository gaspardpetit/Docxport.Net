# Complex-script-only list marker fonts through DOC

Eight generic documents use six Item0–Item5 paragraphs and a single Hebrew alef (U+05D0) marker. They exercise levels 0 and 8, abstract definitions and formatting-only overrides, and either an explicit complex-script font alone or all four font slots. The marker complex-script font is Courier New; mixed-slot controls use Arial/Times New Roman/MS Mincho for the other three slots. Body defaults are Arial. Marker run properties explicitly enable complex-script formatting.

On main 093bbea2, four complex-script-only cases lose their resolved font name during native list reading. The binary font index is present, but list reading calls ResolveFonts only when an ASCII, high-ANSI or East Asian font index exists. Projected DOCX therefore drops the complex-script font and Word renders these markers in Arial. All four mixed-slot controls already preserve Courier New.

The fix adds the complex-script font index to that existing resolution guard. No DOC writer, marker resolver, shaping or layout code changes. A separate Latin forced-complex-script DOC export mismatch found during probing remains outside this input-font fix.

Microsoft Word was opened read-only on all 24 source/DOC/projected documents on 2026-11-15 (local date). word-evidence.json records exact editable ListString values, relative paths and source/PDF SHA-256 hashes. Complete independently extracted PDF text matches all markers and Item0–Item5, with whitespace normalization only. All six marker spans in every PDF use CourierNewPSMT: 144 verified positions. Body spans are excluded from marker-font assertions.

word-baseline.json retains four pre-fix projected DOCX/PDF pairs whose marker spans use ArialMT. These generic files are under the baseline subfolders. DOC fixtures are under fixtures/doc/regressions/list-markers/MarkerComplexFontSlot; Word PDFs are under truth/pdf-word/regressions/list-markers/DocMarkerComplexFontSlot.

DocMarkerComplexFontSlotTests checks the native font index and resolved name, unchanged ASCII slot behavior, projected complex-script font and flag, source numbering immutability, projected schema validity and complete marker/body text through five text visitors.
