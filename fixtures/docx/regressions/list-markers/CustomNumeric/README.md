# Direct custom numeric markers

Ten retained DOCX/Word PDF pairs cover 170 labels: canonical two- through
five-digit padding, unsupported six-digit and single-token numeric patterns,
unknown text, signed boundaries and legal-numbering controls.
Word recognizes fixed format strings rather than arbitrary numeric masks.
Manifests preserve source/capture hashes, starts, logical Word labels and raw
PDF extraction. Ignoring extraction whitespace makes every PDF label agree.
Tests recreate the definitions for Markdown, HTML and plain text and preserve
source numbering XML. Compatibility-branch selection is a separate scope.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomNumeric/)
