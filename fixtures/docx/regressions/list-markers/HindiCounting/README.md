# Hindi counting list markers

Retained source DOCX and Word PDF cover 271 labels: lexical spellings,
combining marks, zero, negative starts, thousands and rollover after 9999.
The manifest records logical Word ListString labels, raw PDF extraction,
starts and artifact hashes. Four extracted labels insert a space inside
one word; ignoring whitespace makes all PDF labels agree with Word ListString.
Export assertions preserve the exact logical spelling and spaces.

The [Word capture](../../../../../truth/pdf-word/regressions/list-markers/HindiCounting/)
uses the standard grouped truth tree. Tests recreate numbering in memory
for Markdown, HTML and plain text. Shared field-number formatting is unchanged.
