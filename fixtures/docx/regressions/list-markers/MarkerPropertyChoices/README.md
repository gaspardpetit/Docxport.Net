# Numbering symbol property choices

Six documents cover direct caps on/off, supported choices on/off, unsupported-namespace fallback, and nested choices inside numbering-level run properties.

Fresh Microsoft Word PDF captures on 2026-11-12 confirm all 36 visible labels and unchanged Item0-Item5 body text. Word ListString remains lowercase even when caps is visibly applied; the PDF is the presentation oracle. See `word-evidence.json` for both logical and visible labels, hashes and paired source/PDF paths.

Tests exercise plain/rich Markdown, HTML and plain text, and verify unchanged source numbering, document and styles XML. This group covers property branch selection; style inheritance, logical reference fields and DOC projection remain separate scopes.
