# Theme modifier endpoints on marker underlines

These 24 synthetic sources have matching Word-exported PDFs under
`truth/pdf-word/regressions/list-markers/HtmlMarkerUnderlineThemeEndpoints/source/`.
The manifest records source/PDF hashes, Word labels and decoration colors, CLI
comparisons, and browser component captures.

Eighteen endpoint cases cover themeTint/themeShade 00 and FF, primary, mixed,
gray and near-gray theme colors, conflicting RGB caches, and both attributes.
Word paints 00 tint white and 00 shade black. FF preserves the base theme color.
When FF tint and 00 shade coexist, Word treats the tint as inactive and paints
black; 00 tint with FF shade paints white. Source themeTint/themeShade values
are recorded independently in the manifest.

Six intermediate-modifier guards retain their existing black HTML output.
They explicitly do not match Word and are excluded from the supported-color
claims. The broader investigation found one-channel color differences in the
intermediate HLS calculation; it remains separate work. This change performs
no HLS calculation and does not alter the shared body/PDF color helper.

Only rich HTML underline presentation changes. Glyphs and independent strikes
stay black. The EmitRunColor option controls this independently of font emission.
Rich/plain presets and both option overrides are tested, as are Office 2013
package validation and source XML immutability. Plain HTML, Markdown and text
must stay byte-identical to main. CLI controls reuse 437 earlier RGB/base-theme/
double/patterned/word-underline fixtures without duplicating their sources/PDFs.
Four representative rich HTML before/after pairs are retained with LF hashes.

Browser components use Arial 12pt and a gray background to make white underline
paint measurable. Checks compare actual RGB with Word, verify black glyphs and
continuous black strikes, and preserve six known intermediate guard outputs.
Exact geometry and whole-page layout are not claimed. PDF/layout work is outside
this change.
