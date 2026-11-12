# Custom English word aliases

Six DOCX/Word PDF pairs cover 102 labels for the exact cardinal and ordinal word
aliases and corresponding direct formats.
Direct controls captured in Word on 2026-11-11 agree with all 34 retained alias
labels, including negative starts, zero, range transitions and signed boundaries.
The manifest records source/capture hashes, starts, logical Word labels and raw
PDF extraction. Ignoring extraction whitespace makes every PDF label agree.
Tests verify exact Markdown, HTML and plain text labels and unchanged source
numbering XML. Two additional direct controls capture exact-thousand forms,
32767/32768 and 999999/1000000 boundaries in Word on 2026-11-11.
Word supports English list words from zero through 999999 and suppresses
negative/out-of-range expansions. Shared field formatting retains its prior
range; the list extension applies only to the built-in English provider.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomEnglishWordAliases/)
