# Custom ordinal and dash aliases

Four DOCX/Word PDF pairs cover 68 labels for the exact ordinal and dashed-number
aliases and corresponding direct formats.
Direct controls captured in Word on 2026-11-11 agree with all 34 retained alias
labels, including negative starts, zero, range transitions and signed boundaries.
The manifest records source/capture hashes, starts, logical Word labels and raw
PDF extraction. Ignoring extraction whitespace makes every PDF label agree.
Tests verify exact Markdown, HTML and plain text labels and unchanged source
numbering XML. Word suppresses the signed minimum while retaining ordinary
negative ordinal and dashed-number labels. Shared field formatting is unchanged.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomOrdinalDashAliases/)
