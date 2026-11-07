# Hexadecimal marker boundary evidence

Seven synthetic DOCX sources and retained Microsoft Word PDF captures cover
zero, negative start overrides, signed overflow, and the 65535 boundary for
hexadecimal list markers. The manifest records source/capture paths and hashes,
visible numeric expansions and each fixture's label delimiters. Captures are
in the matching [truth group](../../../../../truth/pdf-word/regressions/list-markers/HexBoundaries/).

Word emits hexadecimal list expansions only for counter values 0–65535.
The boundary sequence is `FFFE`, `FFFF`, empty, empty. Negative values and
large positive values have empty numeric expansions, preserving surrounding
literal text, paragraph content and counter progress. These are observations
from the supplied Word controls, rather than a universal conformance claim.

Tests recreate the structures in memory and check plain Markdown, HTML and
text with source immutability. Shared hexadecimal number formatting used by
fields retains its existing larger-value and signed behavior. Other formats
and PDF/layout algorithms are outside this change.
