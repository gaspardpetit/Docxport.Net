# Patterned HTML marker underlines

These 69 new synthetic targets cover explicit dotted (`dotted`), dashed (`dash`),
and wavy (`wave`) OOXML underlines. Sources are under `source/`; matching Word PDFs
are under `truth/pdf-word/regressions/list-markers/HtmlMarkerPatternedUnderline/source/`.
The manifest binds each source/PDF pair by SHA-256 and records Word list strings,
PDF text, underline pattern measurements, and independent strike rectangles.

Coverage includes decimal, bullet, and literal labels; no/single/double strike;
RTL paragraphs; tables; compatibility Choice/Fallback; and inherited single/double
strike from paragraph styles. Existing Word-backed controls are reused from
`HtmlMarkerDoubleUnderline`, including absent/disabled/single/double underline and
strike combinations, rather than duplicating those sources or PDFs.

HTML maps these styles to CSS `dotted`, `dashed`, and `wavy`. Single strike stays
solid on its existing separate span; double strike retains its own double style.
C# tests check rich/plain HTML, source XML immutability, and schema validation.
All new packages must validate without exceptions.

CLI comparisons require only marker style changes for the new HTML targets,
byte-identical HTML for the reused controls, and unchanged Markdown/text outputs.
Representative rich HTML baseline/current pairs are retained under `baseline/`
and `html/`, with normalized-LF hashes in the manifest. Two existing TestLists
markers were independently confirmed dashed in a fresh Word export; their four
HTML snapshots change only by adding the marker underline style.

Browser component renders in Arial 12pt compare visible shape families with Word.
The raster probe distinguishes repeated short dots, longer dash runs, and waves
whose vertical position varies across the label. This also covers browser PDF
painting that is not exposed as ordinary vector paths. Strike strokes are checked
separately and must remain continuous rectangles. Versions, PDF/crop hashes, and
measurements are recorded in the manifest. These checks concern shape family and
independent strike count, not exact spacing, thickness, or whole-page layout.

Other underline variants, underline inheritance/color, body-run formatting, and
PDF layout remain separate work.
