# Hindi letter list markers

Two retained synthetic DOCX sources and visible Word PDF captures cover the
complete Hindi letter sets, repetitions, rollover after 304, negative start
overrides, zero and large counters through the signed maximum.
Word uses consonant symbols for hindiVowels and vowel symbols for
hindiConsonants, despite the enum names. The vowel set includes two multi-codepoint
symbols. Zero expands to a space and negatives to empty text.
The captures contain 332 labels.

`word-evidence.json` records source/capture hashes, starts and visible labels.
PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/HindiLetters/).
Tests recreate numbering for Markdown, HTML and plain text.
Generic field formatting retains its decimal fallback.
