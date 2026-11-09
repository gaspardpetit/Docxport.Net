# Custom counting-word aliases

Five retained DOCX/Word PDF pairs cover 85 labels: canonical Hindi, Thai and
Vietnamese counting-word aliases, plus Thai/Vietnamese altered-whitespace
controls that Word treats as decimal fallback. Matching uses exact strings.
The manifest records source/capture hashes, starts, logical Word ListString
labels and raw PDF extraction. Ignoring extraction whitespace makes all PDF
labels agree. Tests preserve exact logical text in Markdown, HTML and plain
text and verify the caller's numbering XML remains unchanged.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/CustomWordAliases/)
