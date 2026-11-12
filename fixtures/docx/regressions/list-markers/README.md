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
| ArabicAlphabets | Arabic alphabetical/Abjad orders, non-joiners and rollover |
| ChineseCounting | Chinese/Taiwanese counting-to-digits transition |
| ChineseLegal | Simplified/traditional Chinese legal numerals and limits |
| ChineseUnitCounting | Chinese/Taiwanese unit counting, internal zeros and limits |
| CustomAlphabetAliases | Exact regional alphabet and Hebrew numeral aliases |
| CustomEnclosedAliases | Exact circled, full-stop and parenthesized number aliases |
| CustomDigitAliases | Exact Hindi, Thai and Taiwanese digit aliases |
| CustomFullWidthAlias | Exact full-width decimal alias and direct boundary controls |
| CustomHexChicagoAliases | Exact hexadecimal and Chicago symbol aliases |
| CustomIdeographAliases | Exact traditional, zodiac and enclosed ideograph aliases |
| CustomJapaneseAliases | Exact Japanese numeric aliases and direct boundary controls |
| CustomChineseCountingAliases | Exact Chinese and Taiwanese counting aliases and direct boundary controls |
| CustomChineseLegalAliases | Exact Chinese legal numeral aliases and direct boundary controls |
| CustomIdeographDigitalAlias | Exact Ideographic digit alias and direct boundary controls |
| CustomKanaAliases | Exact Aiueo/Iroha half-width and full-width aliases |
| CustomKoreanAlphabetAliases | Exact Ganada and Chosung aliases |
| CustomKoreanNumericAliases | Korean counting and digit aliases and boundary controls |
| CustomLatinAliases | Custom Latin letter and Roman aliases |
| CustomNumeric | Direct custom numeric patterns and decimal fallback |
| CustomWordAliases | Exact Hindi, Thai and Vietnamese counting-word aliases |
| DecimalBoundaries | Signed decimal counter boundaries and zero-padding |
| EnclosedBoundaries | Circled, parenthesized and full-stop marker ranges |
| EnclosedIdeographs | Parenthesized ideograph range and signed boundaries |
| GappedHeading | Heading/body counter continuity across missing levels |
| HebrewAlphabet | Hebrew alphabet progression, direction mark and rollover |
| HebrewNumerals | Hebrew additive numerals, special forms and rollover |
| HexBoundaries | Hexadecimal marker range, signed boundaries and overflow |
| HindiCounting | Hindi counting words, lexical forms and rollover |
| HindiLetters | Hindi letter sets, repetitions and rollover |
| HindiThaiDigits | Hindi/Thai digit substitution and rollover |
| IdeographRanges | Traditional and zodiac symbol ranges with decimal fallback |
| JapaneseCounting | Japanese unit counting and the 999999 limit |
| JapaneseDigits | Japanese digit transitions and the 9999 limit |
| JapaneseLegal | Japanese legal numerals and unit transitions |
| Kana | Japanese Aiueo/Iroha width variants and cyclic rollover |
| KoreanDigital2 | Korean ideographic digits, distinct zero and upper limit |
| KoreanCycles | Ganada/Chosung cycles and signed boundaries |
| KoreanUnitCounting | Korean counting/legal units and range limits |
| NegativeStarts | Negative level starts versus standalone overrides |
| Resolution | Independent numbering-property and linked-style resolution |
| Restart | Restart-declaration import normalization |
| RomanBoundaries | Roman marker rollover, zero and signed boundaries |
| RussianBoundaries | Russian alphabetic repetition, rollover and nonpositive counters |
| Suffix | List suffix nothing |
| VietnameseCounting | Vietnamese counting words, lexical forms and rollover |
| ZodiacCycle | Paired traditional zodiac cycles and signed boundaries |
| TaiwaneseKoreanDigits | Taiwanese/Korean digit substitution and signed boundaries |
| ThaiCounting | Thai counting words, contextual forms and range boundaries |
| ThaiLetters | Thai alphabet repetitions, rollover and signed boundaries |
| Text | Level-text expansion and missing-level start rules |

The tests recreate these structures in memory; normal test runs do not require
Word. Capture generation uses the repository's standard fixture-to-truth path
mapping. Review generated captures before replacing the recorded Word evidence
and update manifest hashes only after reviewing the new bytes.
