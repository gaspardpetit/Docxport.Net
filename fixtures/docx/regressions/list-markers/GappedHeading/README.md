# Gapped heading counter evidence

The synthetic DOCX has a level-3 heading followed by two level-0 body
paragraphs, with levels 1 and 2 omitted. Microsoft Word displays `(2)`,
`[2]`, `[3]`; see the captured PDF and machine-readable evidence.
Applying the heading override consumes its unused ancestor override even
across the missing levels. The former `[88]`, `[89]` test expectation
contradicted Word. `PlainMarkdownContinuesParagraphNumberAfterHeadingUsesSameInstance`
recreates this case without a Word dependency during tests.

The other counter tests embed synthetic source XML and expected sequences
from the audit's installed-Word captures. They check source immutability,
OOXML validation and plain Markdown, text and HTML output. The McGregor
regression recreates only numbering transitions with generated paragraph text;
the original judgment is not included.

## Evidence locations

DOCX sources, manifests and supporting data stay together in this group.
Captured Word PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/GappedHeading/).
Each manifest records repository-relative source and capture paths.
