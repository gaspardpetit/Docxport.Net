# PDF paragraph continuations

An OOXML paragraph can produce several MigraDoc paragraphs when it contains a manual page break. The flow visitor now records each rendered fragment's source paragraph, first rendered-text offset, and fragment index in `DxpPdfParagraphFragments`. STYLEREF, exact-line baseline classifications, and fit-text metrics use that source relationship instead of copied entries.

## State ownership

- **Source-wide:** paragraph style, alignment, spacing policy, source borders, and classifications calculated from the whole OOXML paragraph belong to the source. Rendering a fragment should resolve these through its source relationship rather than copy every classification to a new dictionary key.
- **Fragment-local:** text ranges, strike/underline/effect positions, highlights, inline images, bookmarks, and render coordinates belong to the fragment that contains them. Copying a range list or overlay to another fragment would duplicate or misplace content.
- **Break transition:** a new fragment inherits paragraph formatting but not first-line indent or before-spacing. The preceding fragment loses after-spacing. Active run and hyperlink state continue into the new fragment; text-offset counters restart at zero while the registry records the cumulative rendered-text offset.

Inherited value lookup checks a fragment-specific value first, then its root. Inherited classification flags are read from the root. This allows an explicit fragment override without duplicating the source-wide default.

## Migration boundary

The registry is a first step, not a blanket replacement for the existing metadata transfers in `VisitBreak`. Each paragraph-keyed collection must be classified before migration. In particular, moving all metadata by reflection or copying every dictionary entry would be incorrect for offset-bearing ranges and overlays. The post-layout footnote splitter is a separate continuation path and should join this model only after its offset and layout semantics are tested.

Next tests should combine two or more manual breaks with source-wide formatting and fragment-local effects, then cover in-paragraph column breaks using the same transition policy. The current column flow still advances at the next source paragraph, not immediately within one.
