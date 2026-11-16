# Inherited marker size without level run properties

Word displays an all-digit marker at its inherited complex-script size in an RTL
paragraph even when the numbering level has no `w:rPr`. Rich HTML previously
returned before resolving that inherited size. Treating absent level run
properties as an empty, detached formatting view preserves inheritance without
changing the source XML.

The 36 synthetic cases combine:

- Size inherited from paragraph-mark properties, the default paragraph style,
  or document run defaults.
- Absent, empty, or font-only level run properties.
- Generated decimal labels (`1`, `2`, `3`) or literal numeric bullets (`12`).
- Explicit RTL and LTR paragraph direction.

Word's editable `ListString` values, complete PDF text, and independently
extracted PDF marker sizes were checked for every source. RTL markers are 15 pt;
LTR controls are 12 pt. Main at `d62d4286` fails the six absent-properties RTL
cases; the other 30 rich HTML outputs remain byte-identical after the fix.
The actual CLI's plain HTML and Markdown outputs are byte-identical for all 36
inputs as well.

`source/` contains the generic DOCX inputs. The corresponding Word PDFs are under
`truth/pdf-word/regressions/list-markers/HtmlMarkerInheritedSize/source/`.
`baseline/` and `html/` retain before/after rich HTML for the six changing cases;
each difference is solely the marker's `font-size:15pt` attribute. The baseline
producer's production sources were verified identical to the main base.

`word-evidence.json` records source/PDF hashes, Word sequences and marker sizes,
rich output hashes for all 36 cases, paths for the retained HTML, and hashes of
the byte-identical plain output pairs. HTML hashes explicitly use LF line endings
to match Git's stored blobs. The automated tests validate source XML, check rich
and plain marker text/CSS, and assert source XML immutability.
