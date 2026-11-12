# Style definition compatibility for numbering

Thirteen documents cover wrapped paragraph/default/linked numbering styles, paragraph and numbering properties, IDs, basedOn inheritance, supported choices, fallback selection, and ignored IDs. The basedon-wrapper control remains unnumbered under Word's forward style/list association import.

Fresh Microsoft Word ListString and PDF captures on 2026-11-12 confirm 36 visible labels and three suppressed markers, with Item0-Item2 preserved. See `word-evidence.json` for exact logical/visible values and paired source/capture hashes. Tests also compare SDK preprocessing and verify unchanged source styles XML across plain/rich text visitors.

Compatibility normalization precedes association import on a private numbering style view. General font/layout resolution uses original styles. DOC projection remains a separate scope.
