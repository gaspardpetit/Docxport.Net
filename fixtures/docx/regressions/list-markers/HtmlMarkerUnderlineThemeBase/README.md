# Base theme marker underline colors

These 104 generic DOCX sources have matching Word-exported PDFs under
`truth/pdf-word/regressions/list-markers/HtmlMarkerUnderlineThemeBase/source/`.
The manifest records source/PDF hashes, Word labels and decoration colors, CLI
comparisons, and browser component captures. Source packages validate against
Office 2013 after compatibility processing and remain unchanged during export.

The 102 base-theme cases cover decimal, bullet and literal labels; single,
double, dotted, dashed, wavy and word-only underlines; independent single/double
strike; conflicting, absent, auto and near-matching RGB caches; accent colors;
document color mappings; RTL, tables, and compatibility Choice/Fallback.
Word ignores the cached RGB even when it differs by only one channel. Its
marker mapping is a single alias lookup, including chained and cyclic mappings.
The body-run resolver retains its existing default traversal and cache policy.

Two tint/shade guards preserve existing HTML behavior. Word paints them light
and dark green; their HTML remains black, and they are explicitly excluded from
claims of matching Word. Modifier calculations and underline inheritance are
separate work.

Only rich HTML underline presentation changes. Glyphs and independent strike
strokes stay black. Word-only underlines color their child word spans. The
existing EmitRunColor setting controls color independently of font emission.
Plain HTML, Markdown and text stay byte-identical to main. The manifest also
reuses 333 prior RGB/double/patterned/word-underline inputs: three former RGB
theme guards now become green; all other reused HTML output stays identical.
Four representative rich HTML before/after pairs are retained under baseline/
and html/ with LF-normalized hashes.

Browser checks compare underline hue with Word, verify black glyphs and continuous
black strike strokes, and include raster-painted wavy lines. Components use
Arial 12pt; exact geometry and whole-page layout are not claimed. PDF layout
and glyph color are outside this change.
