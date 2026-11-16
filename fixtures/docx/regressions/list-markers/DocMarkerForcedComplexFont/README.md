# Forced complex-script marker faces on DOC export

Twelve synthetic documents use six Item0-Item5 paragraphs and a Latin BEGINabEND bullet. They cover levels 0 and 8, abstract definitions and formatting-only overrides, with CS-only fonts, mixed font slots, or explicitly disabled forcing. The complex-script face is Courier New. Mixed slots use Arial, Times New Roman and MS Mincho; body defaults use Arial.

On main e5146718, Word renders the eight forced source DOCX markers in Courier New but the exported DOC markers in Arial. Projected DOCX already renders Courier New. Four disabled-forcing controls use Arial throughout. The fix materializes an explicitly forced non-symbol complex-script face into the ordinary DOC marker slots. It is limited to list export; body run formatting and PDF/layout code are unchanged. Symbol faces retain their existing encoding slots and are covered by two additional automated controls.

Microsoft Word opened all 36 source/DOC/projected documents read-only on 2026-11-15 (local date). word-evidence.json records editable ListString values, relative paths, SHA-256 hashes and independently extracted PDF marker fonts. Complete PDF text agrees with all six markers and Item0-Item5 using whitespace normalization only. All 216 marker positions use CourierNewPSMT when forced, ArialMT when disabled.

word-baseline.json retains eight generic pre-fix DOC/PDF pairs demonstrating ArialMT before the fix. DOC files are under fixtures/doc/regressions/list-markers/MarkerForcedComplexFont; Word PDFs are under truth/pdf-word/regressions/list-markers/DocMarkerForcedComplexFont. Source and projected DOCX stay grouped here.

DocMarkerForcedComplexFontTests checks native and projected font selection, forcing flags, source numbering immutability, projected schema validity and complete marker/body text through five visitors. The earlier DocMarkerComplexFontSlotTests reads its original checked-in DOC fixtures for native CS-only slot assertions, preserving the reader regression independently of the export normalization added here.
