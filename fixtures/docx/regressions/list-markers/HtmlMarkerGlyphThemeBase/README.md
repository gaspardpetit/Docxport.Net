# Base theme glyph colors on HTML list markers

These 96 synthetic DOCX sources have matching Word-exported PDFs under
`truth/pdf-word/regressions/list-markers/HtmlMarkerGlyphThemeBase/source/`.
The manifest records source/PDF hashes, Word labels and colors, CLI comparisons,
and browser component captures.

There are 93 base-theme targets, a themeColor=none control, and two tint/shade
guards. Cases cover decimal/bullet/literal labels, six underline styles, independent
single/double/no strike, auto/conflicting/near-matching RGB caches, direct accent
colors, document mappings including chains/cycles/self, RTL, tables, and
compatibility Choice/Fallback. Glyph colors require w:color/@val, so absent-cache
cases that would violate that schema are deliberately omitted.

Word ignores glyph RGB caches even when they differ by only one channel, and
resolves document color mappings once. The marker-only resolver uses those rules;
the body-run resolver retains its default traversal and cache policy. Explicit
red underlines remain independent of the theme-colored glyphs and strike strokes.
ThemeColor=none paints black despite its red cache and retains existing HTML.

The two intermediate tint/shade guards keep black HTML output. Word paints them
light/dark green; they explicitly do not match and are excluded from supported
color claims. All glyph modifier calculations, including endpoints, remain a
separate path. This change does not alter the shared body/PDF color helper.

EmitRunColor controls color independently of font emission. Rich/plain presets
and both option overrides are tested, along with Office 2013 package validation
and source XML immutability. Only rich HTML marker color changes. Plain HTML,
Markdown and text stay byte-identical to main. The manifest reuses 547 prior
RGB/theme/endpoint/double/patterned/word-underline inputs: three former glyph RGB
theme guards now become green; all other reused HTML stays identical. Four
representative rich HTML before/after pairs are retained with LF hashes.

Browser components use Arial 12pt. They compare glyph/strike RGB with Word,
retain independent underline hue including raster-painted wavy lines, and
preserve known modifier guards. Exact geometry and whole-page layout are outside
scope. The separate local PDF/layout work is untouched.
