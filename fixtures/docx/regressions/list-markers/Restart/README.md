# Word restart-import evidence

The sources and Microsoft Word PDF captures are synthetic controls retained
from the marker audit. `word-evidence.json` records exact source/PDF hashes,
visible marker sequences and, for the boundary controls, Word's saved numbering
levels. Tests recreate the list structures and need no installed Word.

The boundary matrix covers 168 combinations across levels 0–8, root and
negative/out-of-range declarations, zero/earlier/self/later restart values, and
abstract versus replacement levels. Before extracting the fix, 74 of the 178
restart tests failed on main `4ec19de`; all 178 pass with the resolver change.
Eight representative source/PDF pairs are included here; all 168 boundary test
expectations were rechecked against the retained Word PDFs during extraction.

Observed Word import behavior:

- Root restart declarations and values below zero or above seven become zero
  (never restart).
- Zero and accepted earlier-level values retain their meaning.
- Other self/later-level declarations suppress an abstract level's label.
- Invalid replacement declarations discard the entire replacement override,
  including its start override, so the abstract definition applies again.

These are Word compatibility observations for the tested inputs, not a claim
that every accepted Word value is prescribed by OOXML. The resolver applies
them to a private numbering copy. Tests also verify source immutability and
the counter progress of invisible levels through plain HTML, Markdown and text.
No PDF visitor or layout code is changed by this PR.

## Evidence locations

DOCX sources, manifests and supporting data stay together in this group.
Captured Word PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/Restart/).
Each manifest records repository-relative source and capture paths.
