# Alphabetic list-marker Word evidence

The synthetic sources start uppercase/lowercase lists at 26. Captured Word
PDFs show Z, AA, BB (and z, aa, bb). Additional lowercase sources and Word
PDFs at 780 and 781 preserve the repetition-count rollover evidence. The hash manifest identifies both sources
and PDF captures. The large-label goldens come from the audit's Word list
reference captures, with their visible PDF comparison recorded in the audit.

Word repeats a single Latin letter, wrapping the repetition count after 30:
780 is thirty z characters and 781 is a. Regression tests check the captured
values in Markdown, HTML and text. Generic field formatting retains positional
letters such as AB at 28. This change targets positive Latin list counters;
other numbering formats and field formatting are outside its scope.

## Evidence locations

DOCX sources, manifests and supporting data stay together in this group.
Captured Word PDFs are in [the matching truth group](../../../../../truth/pdf-word/regressions/list-markers/Alphabetic/).
Each manifest records repository-relative source and capture paths.
