# Pdfxport 7.0.1 renderer integration

The PDF exporter uses the pinned Pdfxport 7.0.1 packages by default. Renderer properties provided by that version should be assigned directly, so a renamed or missing API fails the build rather than silently disabling a layout feature.

This slice binds the available contour, line-fit, measured-width, justified-space, and table-fragment reservation callbacks. In particular, the old reflection lookup for `TableFragmentTopBorderWidth` never matched the 7.0.1 property `TableFragmentTopBorderReservation`.

Other optional reflection hooks in `DxpFlowPdfVisitor` are **not** made available by this change. Some are diagnostic experiments; others need separate Pdfxport API work and focused rendering evidence before they can be required. A passing build or unit suite does not imply those hooks ran. In particular, activating the table-fragment reservation does not by itself resolve the observed page-count gap for `OverpageManyRows.docx` on the default 7.0.1 package.

The source fallback remains available through `UsePdfSharpSource=true` or `PdfSharpSourceRoot`; neither is required for this integration.
