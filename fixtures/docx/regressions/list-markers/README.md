# List-marker regression evidence

Each group contains synthetic DOCX sources, a README, and a `word-evidence.json`
manifest with repository-relative paths and exact source/capture hashes.
Supporting data stays in the relevant group. Microsoft Word PDF captures live
under the same group and basename in
[`truth/pdf-word/regressions/list-markers`](../../../../truth/pdf-word/regressions/list-markers/).

| Group | Behavior |
| --- | --- |
| Alphabetic | Word alphabetic rollover and large-label goldens |
| GappedHeading | Heading/body counter continuity across missing levels |
| NegativeStarts | Negative level starts versus standalone overrides |
| Resolution | Independent numbering-property and linked-style resolution |
| Restart | Restart-declaration import normalization |
| Suffix | List suffix nothing |
| Text | Level-text expansion and missing-level start rules |

The tests recreate these structures in memory; normal test runs do not require
Word. Capture generation uses the repository's standard fixture-to-truth path
mapping. Review generated captures before replacing the recorded Word evidence
and update manifest hashes only after reviewing the new bytes.
