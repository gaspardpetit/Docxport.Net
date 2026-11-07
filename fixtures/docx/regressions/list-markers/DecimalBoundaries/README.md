# Decimal marker boundary evidence

Twelve synthetic DOCX inputs and retained Microsoft Word PDF captures cover
decimal and decimalZero counters near signed overflow, at zero, and from
negative start overrides. The manifest records their paths, exact hashes and
six-item Word sequences. Captures are in the matching
[`truth/pdf-word` group](../../../../../truth/pdf-word/regressions/list-markers/DecimalBoundaries/).

Word's decimal marker at Int32.MinValue has an empty numeric expansion. Other
negative decimal values remain visible. Zero-padded decimal markers suppress
all negative expansions. Literal text around the placeholder and paragraph
content remain present, and the signed counter continues advancing, including
through overflow from Int32.MaxValue. Zero remains `0` or `00`, respectively.

The in-memory tests reproduce these sequences in plain Markdown, HTML and
text and verify source immutability. The shared number formatter used by fields
retains its previous behavior. Other numbering formats and marker layout are
outside this PR's scope.
