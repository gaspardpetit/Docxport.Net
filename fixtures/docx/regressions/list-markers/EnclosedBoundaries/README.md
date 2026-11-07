# Enclosed-number marker evidence

Five synthetic DOCX sources and retained Microsoft Word PDF captures contain
111 labels across four enclosed-number formats. The manifest records their
paths, hashes and visible expansions. Captures are in the matching
[truth group](../../../../../truth/pdf-word/regressions/list-markers/EnclosedBoundaries/).

The supplied Word captures use circled symbols through 20 for
decimalEnclosedCircle, through 10 for decimalEnclosedCircleChinese, and through
20 for decimalEnclosedFullstop and decimalEnclosedParen. Zero and values above
these ranges use decimal text. The separate 34–52 sequence confirms that Word
does not use the additional Unicode circled-number blocks for list markers.

In-memory tests reproduce the sequences in plain Markdown, HTML and text,
preserving literal label text, body content and source XML. Shared field-number
formatting retains its existing circled-number behavior. Other numeric formats,
unverified signed limits and PDF/layout algorithms are outside this change.
