# Custom full-width decimal alias

Three DOCX/Word PDF pairs cover 51 labels for the exact full-width decimal
alias and both direct full-width decimal formats. Direct controls captured in
Word on 2026-11-11 confirm empty signed-minimum expansions for both formats.
The alias agrees with decimalFullWidth, including ordinary negative digits;
decimalFullWidth2 suppresses all negative expansions instead.
The manifest records source/capture hashes, starts, logical Word labels and raw
PDF extraction. Ignoring extraction whitespace makes every PDF label agree.
Tests verify exact Markdown, HTML and plain text labels and unchanged source
numbering XML.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomFullWidthAlias/)
