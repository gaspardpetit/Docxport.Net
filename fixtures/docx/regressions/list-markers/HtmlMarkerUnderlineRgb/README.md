# Explicit RGB marker underline color

These 90 synthetic DOCX sources have matching Word-exported PDFs under
`truth/pdf-word/regressions/list-markers/HtmlMarkerUnderlineRgb/source/`.
The manifest records source/PDF hashes, labels, underline colors, black glyphs,
and independent black strike strokes.

There are 75 direct RGB targets, 12 auto/disabled controls, and three theme-cache
guards. Coverage includes decimal/bullet/literal labels, single/double/dotted/
dashed/wavy/word-only underlines, single/double strike, RTL, tables, and compatibility
Choice/Fallback. All packages must validate, and source XML must remain unchanged.

Underline color uses `text-decoration-color`, never glyph `color`. Single strike
gets a separate solid span when needed to retain its own color; double strike
already has one. Word-only underline applies color to the word spans rather than
the outer marker. The existing `EmitRunColor` option controls this presentation
independently of font emission. Rich/plain presets and both option overrides are
tested.

Theme references take precedence over cached RGB. The three guards have a green
theme and a conflicting red cache: Word paints green. The historical CLI/browser captures in this manifest are from PR #319 and retain
its black guard output. The current assertions now require Word's green theme color;
fresh theme-resolution captures are in the adjacent HtmlMarkerUnderlineThemeBase group.

The historical PR #319 CLI checks required only rich HTML underline-color changes for direct RGB targets.
Plain HTML, disabled/auto/theme guards, and other exporters must stay byte-identical.
Controls reuse the prior double-underline, patterned-underline, and word-underline
groups, without duplicating their sources/PDFs. Four representative rich HTML
baseline/current pairs are retained under `baseline/` and `html/`, with LF hashes.

Browser component captures check visible underline hue, black glyphs, and continuous
black strike strokes, including raster-painted wavy lines. The historical 87 supported target/
control cases matched Word; the three theme guards preserved their baseline in #319.
Exact geometry and whole-page layout are not claimed. Underline inheritance, theme
resolution, glyph color, body formatting, and PDF layout remain separate work.
