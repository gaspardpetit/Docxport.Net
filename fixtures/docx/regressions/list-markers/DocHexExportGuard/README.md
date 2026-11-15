# Hexadecimal DOC export guard

Eight generic six-paragraph documents use only Item0 through Item5, with BEGIN/END marker literals. Each format is exercised at level 0 and level 8, as an abstract definition or a formatting-only instance override, starting at 14. Hexadecimal crosses E, F, 10; decimalHalfWidth provides the accepted 0x0F control.

The DOC and projected DOCX files were generated from main fd82a038 before the guard. Installed Microsoft Word was opened read-only on all 24 files on 2026-11-15. Word refused all four hexadecimal DOC files with the error retained in word-rejections.json. Those binaries are retained deliberately to verify historical input projection; they are not valid Word-openable output examples. The source and projected hexadecimal DOCX files open and display the same labels. All twelve half-width decimal source/DOC/projected sequences open and agree.

word-evidence.json records the other 20 editable ListString sequences, source and PDF SHA-256 hashes, and repository-relative paths. Complete PDF text was independently extracted and compared with all labels and Item0–Item5, normalizing whitespace only. PDFs are under truth/pdf-word/regressions/list-markers/DocHexExportGuard; DOC files are under fixtures/doc/regressions/list-markers/HexExportGuard.

DocHexExportGuardTests rejects new hexadecimal DOC exports while retaining five text-visitor renderings for DOCX and historical DOC projection, and writable half-width decimal controls. It verifies native codes and schema-valid projected DOCX.

[MS-DOC LVLF](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/500d1b05-cd15-46af-bc8f-03bc290d1bdd) forbids nfc 0x08, 0x09, 0x0F and 0x13. Word accepts 0x0F here, so that existing behavior is deliberately preserved. Chicago and decimalFullWidth2 output are already unsupported; their historical input mappings remain unchanged.
