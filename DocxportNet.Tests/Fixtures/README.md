# Document fixtures

Keep real-world documents in `samples/`. Keep synthetic and targeted regression
documents here, grouped by the feature under test:

- `Doc/`: existing Word and legacy DOC compatibility cases.
- `DocKnownGaps/`: focused cases for known rendering gaps.
- `DocNativeOracle/`: documents compared with Word's native rendering.
- `Features/`: cross-cutting document feature cases.
- `Probes/`: promoted exploratory probes.
- `TextBoxes/`: text-box and shape placement cases.
- `Omml/`: equation source and expected-expression fixtures.

Fixture inputs belong in source control; generated exports, rasterizations,
and comparison results do not. Word's temporary `~$` lock files are not fixtures.
Preserve existing paths because tests and comparison reports can refer to them.
