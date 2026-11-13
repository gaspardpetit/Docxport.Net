# Root-only definitions with higher-level context

Eighteen cases retain only abstract level 0, starting at 2. Eight cases use
paragraph levels 0/1/1/0/1/1 or 0/8/8/0/8/8, with no override, start only,
formatting only or both. Six cases keep all paragraphs at level 0 but define
a higher-level start/formatting override. Four ordinary single-level controls
retain their existing one-slot representation.

Word leaves undefined child markers blank. Child formatting replacements
activate at 7 and reset after the next parent. A higher-level override alone
also requires DOC's multilevel representation, even when no paragraph uses it.
The private importer now materializes nine positional slots for these used
instances. Missing slots retain blank labels and decimal counters starting at
zero. An additional test guards against expansion by an unused instance.

`source` contains DOCX inputs; `projected` contains API projections of the
editable DOCs under `fixtures/doc/regressions/list-markers/RootHigher`.
Fresh Word captures of all three formats are grouped under
`truth/pdf-word/regressions/list-markers/DocRootHigher`. The hash manifest
records 54 sequences / 324 editable list positions, including blank labels;
independent PDF extraction checks complete labels and body text.

Tests verify imported slots, source numbering immutability, all five text
visitors, projected schema validity, root-only controls and unused instances.
The baseline failed fourteen slot expectations; four ordinary controls passed.
Two pre-fix DOCs and their Word PDFs are retained in `before` subgroups, with
a separate hash manifest. Word renders all six DOC items as root markers 2–7,
losing the blank or replacement child markers and shifting later parent items.
This group covers explicit numId/ilvl pairs and higher overrides. Split
paragraph/style inheritance is covered separately by `DocStyleContext`.
