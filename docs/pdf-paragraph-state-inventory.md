# PDF paragraph state inventory

`DxpFlowPdfVisitor` declares 94 MigraDoc-paragraph-keyed collections, including the separate footnote-split map. This inventory covers those collections, not paragraph-valued reverse maps or lists of render records. A manual page break creates a new MigraDoc paragraph for the same OOXML paragraph. A dictionary lookup against that new paragraph is not automatically equivalent to a lookup against the source fragment.

The registry's flag lookup accepts membership on either the fragment itself or its root. A local-membership regression failed before this change. A full-corpus rescore found unchanged pagination and one 0.000003 smoothed-RMSE improvement (`ruby-raise-probe__ruby-exact-raise-0.docx`); milestone failure counts were unchanged.

## Runtime capability gate

During PR #144, the worktree pinned `ext/PDFsharp` at `95e11d7`; that renderer exposed none of the visitor's reflection-probed callbacks. This is a historical inventory from the Pdfxport 7.0.1 integration. The tracked submodule has since been removed; a separate source checkout can still be selected with `PdfSharpSourceRoot`. Some callbacks, including `TextWidthAdjustment`, were active at that point, while others, including `TabLeaderLayout` and `LineEndOverhang`, were absent. A patch to a callback still needs a test proving that the selected renderer executes it. The callback-backed collections below are classified by intended ownership, **not** automatically marked as validated rendering behavior.

## Ownership rules

- **Source-wide:** computed once from OOXML paragraph/style/context, independent of rendered-text offset. Read through `DxpPdfParagraphFragments`' root relationship. A fragment-specific dictionary value takes precedence.
- **Materialized on transition:** the break changes or clones the actual MigraDoc paragraph state, or a consumer requires a physical per-fragment entry. Keep the explicit transfer until that consumer is migrated.
- **Fragment-local:** ranges, word positions, inline objects, or post-layout measurements. Never inherit these solely by paragraph identity; split or recompute them using `RenderedTextStart` and the new fragment's coordinates.
- **Unresolved:** ownership or consumer semantics need a targeted fixture before migration. A name containing “paragraph” is not evidence that inheritance is safe.

## Classified collections

**Source-wide, resolved through the fragment root in the relevant consumer:**

`_exactCjkLineParagraphs`, `_exactStylelessPlainParagraphs`, `_exactStylelessImplicitAptosParagraphs`, `_topAlignedFontSizes`, `_topAlignedHeaderParagraphs`, `_verticalAlignedFonts`, `_fitTextWidths`, `_fitTextCharacterCounts`, `_mixedCjkFitBodyParagraphs`, `_latinTimesFitParagraphs`, `_mixedCalibriYuMinchoFitParagraphs`, `_latinYuMinchoFitParagraphs`, `_fittedBodyAfterSecondCjkHeader`, `_latinCalibriFitParagraphs`, `_fitTextTracking`, `_hangulBaselineShiftParagraphs`, `_mixedHighAnsiFittedHeaders`, `_compactArialBodyParagraphs`, `_directBorderParagraphs`, `_footerParagraphShadeClearances`, `_footerLeaderParagraphs`, `_narrowJustifiedBodyParagraphs`, `_borderedTabInsets`, `_unshiftedGridGlyphParagraphs`.

The Hangul baseline read is active post-layout state. A two-fragment PDF regression showed a 0.28-point baseline mismatch before root lookup and no mismatch afterward. On the 1,156-document corpus, the change caused no pagination or smoothed-RMSE difference among the 1,155 sources with Word references; one source remains unscored because its Word reference is missing.

**Source-wide, migrated but callback-only in this build:**

`_frameWidthParagraphs`, `_implicitTableParagraphs`, `_tocLeaderParagraphs`, `_positionalLeaderParagraphs`, `_clearedTabLeaderParagraphs`, `_trackedLayoutParagraphs`, `_mixedFullWidthPunctuationParagraphs`, `_justifiedShrinkFactors`, `_twoWordOverhangParagraphs`.

These values are computed from the OOXML paragraph or its surrounding style/context. All paragraph-keyed reads now resolve through the fragment registry, but the pinned PDFsharp build does not expose their `TabLeaderLayout`, `TextWidthAdjustment`, line-overhang, justified-shrink, or mixed-font-metrics hooks. A two-page tab-leader probe rendered identically with and without inherited lookup. The migration preserves the ownership model, but it is **not** evidence of rendering fidelity until those hooks become active and receive observable continuation probes. Hook integration remains a separate PDFxport task.

The callback-read migration passed all 60 focused flow tests. A 1,156-document export rerun retained all page counts. Among 1,155 nonempty PDFs, 1,148 were pixel-identical at the scorer's 400×520 resolution; the other seven retained identical Word-reference smoothed RMSE to six decimals on focused rescore. The callback capability gate remains closed.

`_narrowJustifiedBodyParagraphs` differs: its read in `JustifyContinuedPageLines` is active. A tracked-revision paragraph split by a manual page break had 2.52-point final-line gaps before root lookup and 3.07-point gaps afterward; the regression asserts the latter. Its physical transfer for post-layout footnote splits remains in place.

The subsequent 1,156-document rerun again had no pagination changes. At the visual scorer's 400×520 resolution, 1,150 exported PDFs were pixel-identical to the prior run. Five differed in pixels but had the same Word-reference smoothed RMSE to six decimals after focused rescoring. The remaining export is the pre-existing empty `WordPageBorderSectionSetup` result. Thus the existing 1,155-document Word scorecard has no measured score change from this migration.

`_borderedTabInsets` is computed from source-wide run borders and tab stops. A bordered, tabbed run split by a manual page break had a 14.5-point border-to-last-glyph clearance on its first fragment but only 9.5 points on its continuation. Root lookup restores 14.5 points on both fragments; the PDF geometry regression checks that clearance.

The full 1,156-document rerun after this migration had no pagination changes. Of 1,155 nonempty PDFs, 1,146 were pixel-identical to the preceding run at the scorer's 400×520 resolution. Nine changed in pixels, but focused Word-reference rescoring found unchanged smoothed RMSE to six decimals for all nine. The existing empty export remains unchanged.

**Materialized on a physical transition or by `AddParagraph`:**

`_sourceBodyParagraphs`, `_sourceCellParagraphs`, `_suppressedLineNumberParagraphs`, `_paragraphShades`, `_manualParagraphShades`, `_characterScales`, `_characterSpacings`, `_tabbedRunShadings`, `_coloredUnderlineParagraphs`, `_bodyParagraphs`, `_mirroredParagraphs`, `_layoutWidthParagraphs`, `_compressedBodyParagraphs`.

**Fragment-local; do not inherit wholesale:**

`_emphasisRanges`, `_strikeRanges`, `_revisionUnderlineRanges`, `_revisionParagraphs`, `_formattingBalloons`, `_strikeTabBoundaries`, `_textEffectRanges`, `_partialFitGroups`, `_rubyGuides`, `_rubyNaturalLineHeights`, `_rubyOriginalLineSpacingRules`, `_rubyOverflowBaselineOffsets`, `_textBorderRanges`, `_paragraphImages`, `_mathFractions`, `_inlineMathFractions`, `_nestedInlineMathFractions`, `_inlineMathMatrices`, `_inlineMathCombinedScripts`, `_highlightParagraphTexts`, `_crossRunLeadingWords`, `_footnoteSplitContinuations`, `_standaloneInlineMathRadicals`, `_mixedInlineMathRadicals`, `_headerRunBorderParagraphs`, `_footerRunBorderParagraphs`, `_borderResetParagraphs`, `_emergencyBreakParagraphs`, `_oversizedGridLatinParagraphs`, `_oversizedStoryLatinParagraphs`, `_centeredSpacingOffsets`, `_listMarkerUnderlines`, `_listMarkerStrikes`, `_bulletBaselineParagraphs`, `_gridGlyphParagraphs`, `_gridBoundaryParagraphs`, `_packedGridParagraphs`, `_packedGridOffsets`, `_mathRadicals`, `_inlineMathRadicals`, `_tallInlineMathRadicals`, `_noBreakParagraphs`, `_superscriptHeaderParagraphs`, `_cjkNumberedBodyParagraphs`, `_firstBodyDisplayFractions`, `_mixedDisplayMathParagraphs`, `_mixedDisplayRunFractionParagraphs`, `_firstBodyDoubleBorderInset`.

`_compressedBodyParagraphs` is source-wide for manual-break fragments, but a post-layout footnote split creates a new physical paragraph outside the manual-break registry. A shortened-page variant of `Right to Repair.docx` produces repeated synthetic splits: without physical transfer, the second continuation loses compressed spacing and renders its spaces at 8.5 rather than 9.5 points. Transfer this flag to each synthetic continuation. The same regression checks `_layoutWidthParagraphs` transfer for another compressed paragraph in that sample.

After this transfer, all 64 focused flow tests pass. The 1,156-document rerun retains every page count. Among 1,155 nonempty PDFs, 32 differ in 400×520 raster pixels from the preceding run; 1,042 documents have Word references in the current scorer directories. Compared with the earlier Word scorecard, only 11 smoothed-RMSE values change, all improving by 0.000001. The other changed PDFs are not covered by that scorecard and need Word references before claiming an absolute fidelity result.

`_layoutWidthParagraphs` temporarily gives each actual MigraDoc paragraph a one-point wider layout right indent and restores it after page preparation. A dense tracked paragraph with a manual page break left its continuation at −1 point while the root was restored to zero. Transferring membership to the continuation restores both physical paragraph formats; a regression inspects the rendered MigraDoc state and PDF page count. The separate natural-page and footnote-split matching in `RestoreLayoutWidthContinuationPaint` remains unchanged.

The following 1,156-document rerun had no pagination changes. Among 1,155 nonempty PDFs, 1,151 were pixel-identical to the previous run at the scorer's 400×520 resolution; the other four retained identical Word-reference smoothed RMSE to six decimals on focused rescore. The pre-existing empty export was unchanged.

`_compactArialBodyParagraphs` is assigned from paragraph formatting and read by underline/border overlays. A two-page underline probe showed different rectangle heights (1 versus 0.25 points) before root lookup; the same geometry follows both fragments afterward. A second full-corpus pass found no changed pagination or smoothed RMSE among 1,155 scored documents. List-marker and math collections cannot be inferred from their names alone because the relevant glyphs or objects may occur on only one fragment.

`_directBorderParagraphs` maps only paragraphs with direct OOXML borders to their source. The footnote post-layout split keeps its physical transfer; manual page-break fragments resolve the mapping through the root. A two-page, two-paragraph red-side-border probe had a 6-point gap on both sides before root lookup; the continuation now supplies both bridge strokes. The 1,155 scored corpus documents had unchanged pagination and smoothed RMSE after this change.

The optional `TopBorderDistanceAtAreaStart` diagnostic callback also resolves this map through the root. That callback is absent from the pinned renderer, so its migration is ownership-consistent but has no current rendering evidence.

`_footerParagraphShadeClearances` is computed from the source paragraph's run borders. A footer with a shaded paragraph split by a manual page break painted only one wide shading rectangle before root lookup; both fragments now receive their clearance-based shading rectangle. The existing shading-color and physical-transfer paths remain unchanged. The full-corpus rerun kept the same pagination and milestone-failure counts; one low-error ruby probe varied by 0.000003 in smoothed RMSE between repeated runs.

`_footerLeaderParagraphs` also affects an active double-run-border overlay, independently of its inert `TabLeaderLayout` callback. A footer split by a manual page break painted the continuation's outer border two points lower before root lookup. The continuation regression now checks equal outer-border positions.

The full 1,156-document export rerun after this change had no pagination differences from the previous run. Of the 1,155 documents in the prior Word scorecard, 1,042 were rescored against the same Word reference PDFs with no smoothed-RMSE change. The other 113 had identical 400×520 raster pixels on every page between the prior and current flow PDFs, so their visual scores against the unchanged Word references cannot change either. One document still lacks a Word reference.

`_headerRunBorderParagraphs`, `_footerRunBorderParagraphs`, and `_borderResetParagraphs` are set while adding run text, rather than when entering the source paragraph. Their meaning is tied to the fragment that contains the bordered or unbordered run; inheriting them would misclassify a continuation with different runs.

`_emergencyBreakParagraphs`, `_oversizedGridLatinParagraphs`, and `_oversizedStoryLatinParagraphs` are also set while emitting the particular run that triggers them. `_centeredSpacingOffsets` is calculated from the current fragment's rendered-text length. List-marker underline, strike, and bullet-baseline overlays must attach to the fragment that actually contains the marker; inheriting them would let later fragments select unrelated words or glyphs.

Grid-glyph and packed-grid flags and offsets are set when the triggering CJK text is emitted. Math-radical values and flags are set while visiting the specific math object. No-break and superscript-header flags are likewise attached to the fragment containing the triggering content. These must not be inherited into a continuation that has no corresponding content. `_unshiftedGridGlyphParagraphs`, by contrast, depends on the source paragraph's adjusted-indent property and document grid. A manual break inside one run left the continuation's grid-glyph baseline 0.5 points higher before root lookup; both baselines match after lookup. The regression exercises the active post-layout metric path.

The grid-state migration passed all 62 focused flow tests. A full 1,156-document rerun retained all page counts. Of 1,155 nonempty PDFs, 1,152 were pixel-identical to the preceding run at the scorer's 400×520 resolution; the other three retained identical Word-reference smoothed RMSE to six decimals on focused rescore. The pre-existing empty export was unchanged.

The CJK-numbered-body correction is attached to the fragment where the list marker is emitted. Display-fraction and mixed-display-math flags are attached while visiting the math object and accompany fragment-local fraction records; inheriting those flags without the corresponding math content would corrupt unrelated continuation text.

`_firstBodyDoubleBorderInset` compensates only the first fragment's extra before-spacing. A first-body double-run-border probe produced identical border rectangles on both sides of a manual page break with exact-key lookup. Temporarily inheriting the value moved the continuation border 4.7 points and failed the regression, so this value must remain fragment-local despite being computed from the source paragraph.

## Separate continuation paths

`_footnoteSplitContinuations` records post-layout split paragraphs; it must not be treated as a manual page-break fragment. Column breaks currently advance on the next source paragraph and need their own transition policy. Section breaks change section geometry, so neither paragraph-root inheritance nor a copied format is sufficient by itself.

## Next evidence gates

1. When the PDFxport renderer exposes the callback hooks, add observable continuation fixtures for tab leaders, overhang/justification, vertical alignment, and tracked layout; compare page-level geometry to sources without breaks.
2. Audit `GetPageStoryRenderInfos`, `RenderEmphasisMarks`, and footnote/column paths independently before extending the manual-page-break registry to other continuation kinds.
3. Re-score changed fixtures against Word and compare the full corpus to the parent commit before requesting another automated review. The current 1,156-document comparison establishes this migration's incremental effect, not complete Word-reference coverage.
