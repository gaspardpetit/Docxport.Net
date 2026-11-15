# Explicit list-level paragraph-style links

Eight generic six-paragraph documents contain only Item0 through Item5 and BEGIN/END marker literals. They exercise a nested paragraph style at levels 1 and 8, with or without an explicit list-level w:pStyle association, each with ordinary or legal numbering. The Nested style retains its own numbering properties in all controls.

On main fd98c60e, DOC export writes the no-style sentinel (0x0FFF) for every LSTF.rgistdPara slot. Word opens the files and displays the correct existing numbers, but ListTemplate.ListLevels.Item(level + 1).LinkedStyle is empty even for the explicitly linked input. This is loss of editable list metadata, rather than an existing-paragraph sequence correction. The fix preserves the explicit link through DOC output, native input reading and DOCX projection.

Microsoft Word was opened read-only on all 24 source/DOC/projected documents on 2026-11-15. In the four explicitly linked cases the exported DOC now reports LinkedStyle = Nested, matching source and projection. Unlinked DOC controls retain an empty linked style. Word infers Nested in source/projected DOCX controls from their paragraph-style numbering; this PR preserves explicit level metadata only and does not add inference for unlinked DOC output.

word-evidence.json records all 24 editable ListString sequences and LinkedStyle observations, repository-relative paths and source/PDF SHA-256 hashes. Complete independently extracted PDF text matches every marker and Item0â€“Item5, normalizing whitespace only. PDFs are under truth/pdf-word/regressions/list-markers/DocListStyleLinks; DOC files are under fixtures/doc/regressions/list-markers/ListStyleLinks.

word-baseline.json and the two generic baseline DOC files retain the previous empty LinkedStyle observations. DocListStyleLinkTests checks the binary style index, projected w:pStyle resolution to the Nested paragraph style, source numbering/style immutability, projected schema validity and five text visitors. Unlinked controls ensure no association is introduced without an explicit source link.

[MS-DOC LSTF](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/6682e74f-365f-45e6-8b06-0e898bbf69a1) defines rgistdPara as the linked paragraph-style indexes, with 0x0FFF indicating no link. [STSH](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/c8ee0f39-02c3-4caa-b27a-6a97600130fe) defines the valid style-index range.

Codex review identified a schema-order defect in the initial implementation when a level combined w:pStyle and w:isLgl. Both linked legal-numbering cases reproduced the schema failure. Schema-aware child insertion fixes their order; all eight projected documents pass OpenXmlValidator, including unlinked legal-numbering controls.
