# Custom Greek numerals

Eight DOCX/Word PDF pairs cover 356 labels for upper/lower Greek numeric
extensions, abstract zero and legal-numbering controls. Word combines additive
symbols, prefixes thousands with a comma and cycles positive values through
9999. Negative expansions are empty and zero expands to a space.
Four legal-numbering controls captured afresh in Word on 2026-11-11 confirm
decimal output when isLgl is enabled and Greek output when explicitly disabled.
The manifest records source/capture hashes, starts, logical Word labels and raw
PDF extraction. Ignoring extraction whitespace makes every PDF label agree.
Tests verify exact Markdown, HTML and plain text labels and unchanged source
numbering XML. Shared field formatting remains unchanged.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomGreekNumerals/)
