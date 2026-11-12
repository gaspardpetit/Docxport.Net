# Bullet label placeholders

Four DOCX/Word PDF pairs cover 24 labels with bullet, decimal and Roman
ancestors, plus an ordinary literal bullet control. Word expands numeric ancestors in a bullet label while leaving
references to a bullet's own numeric value empty. Ancestor counters continue
across visits to the bullet level. Fresh Word captures on 2026-11-11 agree with all retained labels.
The manifest records source/capture hashes,
levels, logical Word labels and raw PDF extraction. Ignoring extraction
whitespace makes every PDF label agree.
Tests verify plain/rich Markdown and HTML, plain text and equivalent source
numbering XML. Literal escaping and marker caps inheritance are separate work.

[Word captures](../../../../../truth/pdf-word/regressions/list-markers/BulletPlaceholders/)
