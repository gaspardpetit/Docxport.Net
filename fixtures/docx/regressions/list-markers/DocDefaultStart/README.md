# Omitted abstract-level starts through DOC

Twelve cases cover levels 0, 1 and 8 with omitted starts and explicit starts of 0, 1 and 7. A thirteenth control preserves a root replacement with an omitted start that currently projects correctly from 1. Source inputs are under `source/`; API projections are under `projected/`; editable DOC exports are under `fixtures/doc/regressions/list-markers/DefaultStart`.

Fresh Word ListString and PDF captures on 2026-11-12 confirm all 39 source/DOC/projected sequences (234 editable labels). Independent PDF extraction confirms the visible labels and six body items in each file. The manifest retains source/output/capture paths and hashes. Word PDFs are grouped under `truth/pdf-word/regressions/list-markers/DocDefaultStart/{source,doc,projected}`.

An omitted abstract-level start must import as zero, matching the existing text renderer, initial Word labels and labels after restarts. The three omitted-start cases previously imported as one. Nine explicit-start controls remain unchanged. The replacement control confirms that broadening this default change to replacement levels would regress an already correct root sequence; replacement activation/projection requires its own follow-up.

The zero default agrees with the [Open XML starting-value specification](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.startnumberingvalue?view=openxml-3.0.1). Word-visible output is the acceptance oracle. PDF/layout code is untouched.
