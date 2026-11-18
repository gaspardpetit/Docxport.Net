# Single-strike inheritance for HTML list markers

These 90 synthetic DOCX files cover decimal, bullet, and literal (`none`) labels.
Sources are under `source/`, and matching Word-exported PDFs are under
`truth/pdf-word/regressions/list-markers/HtmlMarkerStrikeInheritance/source/`.
The manifest binds sources and PDFs by SHA-256 and records Word list strings,
PDF text, and the thin filled rectangles Word uses to strike the markers.
Each enabled case has one rectangle per marker; those rectangles cross marker
text and never cross the explicitly unstruck body text.

Coverage includes paragraph-mark, paragraph-style, based-on style, default
paragraph-style, and document-default declarations; absent, empty, and font-only
level properties; explicit level/paragraph overrides; nearest style values;
Word-compatible style paragraph-mark properties; compatibility Choice/Fallback;
RTL paragraphs; and table paragraphs.

Word treats style paragraph-mark strike as a separately inherited toggle when
the same style also declares run strike. A mark-only declaration supplies the
style value. The style-chain cases preserve both effects, including a child with
only a disabled mark and an inherited enabled mark toggle. Direct paragraph-mark
and level declarations override the inherited result, including explicit false.

C# tests check rich and plain HTML, source XML immutability, and package validity
after compatibility preprocessing. The only permitted validation errors are the
exact style `pPr/rPr` locations accepted by Word. All other sources must validate.
HTML's existing plain configuration retains marker decorations, so both HTML
configurations receive the resolved strike value.

CLI comparisons check that HTML changes are confined to marker strike CSS and
that Markdown and text exports remain byte-identical. Four representative rich
HTML baseline/current pairs are retained under `baseline/` and `html/`, with
normalized-LF hashes in the manifest. Double-strike and underline rendering are
separate work; no body-run or PDF layout code changes are included.
