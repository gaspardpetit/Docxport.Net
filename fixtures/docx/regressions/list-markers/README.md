# List-marker regression evidence

Each group contains synthetic DOCX sources, a README, and a `word-evidence.json`
manifest with repository-relative paths and exact source/capture hashes.
Supporting data stays in the relevant group. Microsoft Word PDF captures live
under the same group and basename in
[`truth/pdf-word/regressions/list-markers`](../../../../truth/pdf-word/regressions/list-markers/).

| Group | Behavior |
| --- | --- |
| Alphabetic | Word alphabetic rollover and large-label goldens |
| AlphabeticBoundaries | Alphabetic zero, negatives and signed overflow |
| ChicagoBoundaries | Chicago symbol repetition, rollover and nonpositive counters |
| DecimalBoundaries | Signed decimal counter boundaries and zero-padding |
| EnclosedBoundaries | Circled, parenthesized and full-stop marker ranges |
| EnclosedIdeographs | Parenthesized ideograph range and signed boundaries |
| GappedHeading | Heading/body counter continuity across missing levels |
| HexBoundaries | Hexadecimal marker range, signed boundaries and overflow |
| IdeographRanges | Traditional and zodiac symbol ranges with decimal fallback |
| Kana | Japanese Aiueo/Iroha width variants and cyclic rollover |
| NegativeStarts | Negative level starts versus standalone overrides |
| Resolution | Independent numbering-property and linked-style resolution |
| Restart | Restart-declaration import normalization |
| RomanBoundaries | Roman marker rollover, zero and signed boundaries |
| RussianBoundaries | Russian alphabetic repetition, rollover and nonpositive counters |
| Suffix | List suffix nothing |
| Text | Level-text expansion and missing-level start rules |

The tests recreate these structures in memory; normal test runs do not require
Word. Capture generation uses the repository's standard fixture-to-truth path
mapping. Review generated captures before replacing the recorded Word evidence
and update manifest hashes only after reviewing the new bytes.
