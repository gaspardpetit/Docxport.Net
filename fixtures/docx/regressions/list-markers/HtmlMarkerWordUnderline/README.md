# Word-only HTML marker underline

These 96 synthetic DOCX sources have matching Word-exported PDFs under
`truth/pdf-word/regressions/list-markers/HtmlMarkerWordUnderline/source/`.
The manifest records source/PDF hashes, Word list strings, per-marker underline
runs, and independent strike rectangles. All new packages validate without errors.

Word excludes ordinary spaces, tabs, nonbreaking spaces, and ideographic spaces
from word-only underline. Captured em spaces and narrow nonbreaking spaces remain
underlined within a word. The fixtures cover these distinctions, repeated spaces,
decimal/bullet/literal labels, font/no-font paths, single/double strike, RTL, tables,
compatibility branches, and literal markup, quotes, and entity-looking text.

The renderer decorates escaped text segments and preserves their original encoding.
Entity decoding is used only to classify a token as a gap; decoded text is never
inserted as markup. The generated-font HTML flag allows the existing presentation
span's text nodes to be decorated without treating literal label text as HTML.

C# tests verify exact visible label text, expected underlined words, independent
strike, rich/plain HTML, source XML immutability, and schema validity. CLI checks
reverse only the new decoration markup and require the original bytes otherwise,
including whitespace and entity representation. Existing controls are reused from
`HtmlMarkerDoubleUnderline` and `HtmlMarkerPatternedUnderline`; no control sources
or PDFs are duplicated. Representative rich HTML baseline/current pairs are kept
under `baseline/` and `html/`, with normalized-LF hashes in the manifest.

Browser captures compare underline coverage with the expected words and require
excluded gaps to stay unpainted. Native browser ink skipping inside a word is
retained. Strike counts are checked separately. These checks concern word coverage,
not exact stroke geometry, whitespace layout, line breaking, or whole-page layout.
Other underline variants, inheritance/color, body-run formatting, and PDF layout
remain separate work.
