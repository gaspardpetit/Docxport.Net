# Custom Turkish alphabets

Ten DOCX/Word PDF pairs cover 642 labels for upper/lower Turkish alphabet
extensions, exact-identifier fallback controls, and abstract-zero controls.
Word recognizes the historical circumflex identifiers, uses Turkish letters
with breve in the alphabet, repeats each letter and wraps after 30 repetitions.
The similar identifiers containing breve fall back to decimal. Negative
expansions are empty and zero expands to a space for recognized identifiers.
The manifest records source/capture hashes, starts, logical Word labels and raw
PDF extraction. Ignoring extraction whitespace makes every PDF label agree.
Tests verify exact Markdown, HTML and plain text labels and unchanged source
numbering XML. Four legal-numbering controls captured afresh in Word on
2026-11-11 confirm decimal output when isLgl is enabled and Turkish output
when explicitly disabled, across another 68 labels.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomTurkishAlphabets/)
