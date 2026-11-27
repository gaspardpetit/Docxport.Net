# Explicit RGB glyph colors on HTML list markers

These 86 generic DOCX sources have matching Word-exported PDFs under
`truth/pdf-word/regressions/list-markers/HtmlMarkerGlyphRgb/source/`.
The manifest records source/PDF hashes, Word labels and colors, CLI comparisons,
and browser component captures.

There are 77 direct RGB targets, six auto/absent controls, and three theme-color
guards. Coverage includes decimal/bullet/literal labels, single/double/dotted/
dashed/wavy/word-only or disabled underline, single/double/no strike, blue/red/
green glyphs, RTL, tables, and compatibility Choice/Fallback. Word paints direct
RGB glyphs and strike strokes in the requested color while retaining an explicit
red underline. Body text stays black. All packages validate and source XML stays
unchanged during export.

EmitRunColor controls glyph color independently of font emission. Rich/plain
presets and both option overrides are tested. Direct colors must be six hex
characters. A theme reference overrides its RGB cache: the three theme guards
have a green theme and red cache; Word paints green. This manifest's CLI/browser
captures are historical PR #322 evidence and retain its black guard output.
Current assertions now require green; fresh base-theme captures and CLI comparisons
are in the adjacent HtmlMarkerGlyphThemeBase group. Color inheritance remains
separate work.

In the historical PR #322 comparison, only the outer marker's rich HTML color changes. The browser uses currentColor
for independent strike strokes; the existing underline color remains separate.
Plain HTML, Markdown and text stay byte-identical to main. CLI controls reuse 461
prior RGB/theme/endpoint/double/patterned/word-underline inputs without duplicating
sources/PDFs. Four representative rich HTML before/after pairs are retained with
LF-normalized hashes.

Browser components use Arial 12pt. Captures compare glyph and strike RGB with
Word, preserve independent underline hue (including raster-painted wavy lines),
and retained theme glyph guards at #322. Exact geometry and whole-page layout are outside
scope. The separate PDF/layout work is untouched.
