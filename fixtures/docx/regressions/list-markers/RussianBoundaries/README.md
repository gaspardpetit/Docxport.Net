# Russian alphabetic list marker boundaries

Retained synthetic sources and visible Microsoft Word PDF captures cover both
cases of the Russian alphabet, first and repeated letters, the 30-repetition
rollover, large counters, zero sequences and negative starts versus overrides.

Word repeats one letter at a time and wraps the repetition count after 30.
Zero expands to a space and negative overrides to empty text. Invalid negative
abstract starts suppress numbering. Shared field formatting stays unchanged.

`word-evidence.json` records source/capture hashes and every visible marker.
PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/RussianBoundaries/).
Tests exercise Markdown, HTML and plain text without requiring Word.
