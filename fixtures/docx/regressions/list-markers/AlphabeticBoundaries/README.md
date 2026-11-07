# Alphabetic marker boundary evidence

Seven synthetic DOCX sources and retained Microsoft Word PDF captures cover
zero, negative start overrides and signed overflow for Latin alphabetic list
markers. The manifest records exact hashes, paths and numeric expansions.
Captures are in the matching [truth group](../../../../../truth/pdf-word/regressions/list-markers/AlphabeticBoundaries/).

Word expands zero to a space and negative values to empty text. Positive
letters retain their established repetition/rollover behavior. Tests check
counter progress, literal label text, paragraph content and source immutability
in plain Markdown, HTML and text. A lowercase zero control covers a–e as well.

Shared alphabetic formatting used by fields retains signed/zero values and
positional letters such as AB. Other alphabets and PDF/layout algorithms are
outside this change.
