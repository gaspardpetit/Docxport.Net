# Independent DOC start and formatting override flags

Twelve source documents cover levels 0, 1 and 8 with no override, start only,
formatting only, and both overrides. The abstract start is 2; explicit starts
and replacement-level starts are 7. Each document has six editable list items.

Word displays 7–12 for start-only and combined overrides, and 2–7 for the
ordinary and formatting-only controls. Before this fix, the combined cases
exported to DOC displayed 2–7: the writer lost the independent start flag.
The reader also discarded the start flag when formatting was present.

The matching binary files are grouped under
`fixtures/doc/regressions/list-markers/OverrideFlags`. Fresh Word captures of
both formats are under `truth/pdf-word/regressions/list-markers/DocOverrideFlags`.
`word-evidence.json` records file hashes, editable ListString labels and
independently extracted PDF labels (24 sequences, 144 labels).
The three pre-fix DOC Word PDFs are retained in the `before` subgroup with
their hashes and extracted 2–7 sequences in the same manifest.

The focused tests inspect the binary flags and decoded start values for every
case, and plain Markdown projection for all twelve cases. The formatting-only
projection issue originally deferred from this group is now covered by
`DocFormatOverrides`, including distinct Roman replacement formats and fresh
Word source/DOC/projected captures.

The production fix preserves independent LFOLVL fStartAt/fFormatting bits and
reads the combined start from the nested LVL, leaving the ignored outer start
field unused. See [MS-DOC LFOLVL](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/2555c6ec-604b-48fd-8484-a2fdbb95b42c).
