# Roman marker boundary evidence

Nine synthetic DOCX sources and retained Microsoft Word PDF captures cover
the first two 32767-value rollover boundaries, zero, negative overrides and
signed integer overflow. Sources and manifests stay in this group; captures
are in the matching [truth group](../../../../../truth/pdf-word/regressions/list-markers/RomanBoundaries/).
The manifest records exact hashes, paths, numeric expansions and label delimiters.

Word restarts positive Roman list labels at I after 32767, without resetting
the numeric counter. Zero expands to a space; negatives expand to empty text.
Tests retain literal label text and body content and assert source immutability
across plain HTML, Markdown and text. A lowercase zero control covers i–v as
well. Uppercase large-counter controls also show signed overflow and the later
return from empty negative labels to visible labels.

The shared Roman formatter used by fields retains its previous behavior.
Other numeric formats and PDF/layout algorithms are outside this PR.
