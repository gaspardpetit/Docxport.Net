# Word PDF reference generation

Run `make truth` on Windows with Microsoft Word installed to export every DOCX under
`fixtures/docx` to a PDF under `truth/pdf-word`. The fixture-relative directory and
filename are preserved. For example, `fixtures/docx/features/Example.docx` becomes
`truth/pdf-word/features/Example.pdf`. This target replaces existing PDFs.

Run `make truth-missing` to export only fixtures whose PDF does not exist. It does not
start Word when every selected PDF is already present.

Use `TRUTH_FILTER` to select fixture **filenames** with a wildcard, and `TRUTH_DIR`
to write to another location while inspecting a change:

```sh
make truth-missing TRUTH_FILTER='StyledRepeatedSpaces' TRUTH_DIR=output/pdf/truth-smoke
```

Exports use Word's clean view, without revision balloons. Review generated references
before committing them; they are the comparison truth for PDF fidelity checks.
If Word rejects a fixture on normal open, the script retries with Word's Open and Repair.

Run `make truth-image` to create local PNGs from the Word PDFs under
`output/pdf/word-truth-images`. It renders each page at 400×520 and 1200×1560,
matching the current scorer's page pixel sizes and white-canvas padding. The target
refreshes an image set when its PDF timestamp, PDF size, renderer version, or rendering
recipe changes; otherwise it leaves that image set untouched. `TRUTH_FILTER` and
`TRUTH_IMAGE_DIR` can limit the run or redirect the output. The images are generated
locally and are not committed.
Install the pinned renderer dependencies with
`python -m pip install -r tools/doc-docx/requirements-truth-image.txt`.

Run `make score` for the complete Word comparison workflow. In order, it checks
that every selected DOCX has a readable Word PDF, refreshes local Word images,
exports Flow PDFs through the Docxport CLI, checks those PDFs, and produces document
and page scorecards. The export uses the CLI's default installed-font provider with
cached fields and requires no Aspose license. Use `TRUTH_FILTER` for a focused run.
Outputs are local under `output/pdf/current-flow` and
`output/pdf/word-image-scores` by default.

Each step can run independently: `make check-truth`, `make truth-image`,
`make render-flow`, `make check-flow`, `make score-current`, and `make score-diff`.
`render-flow` builds the CLI once, then invokes it for each fixture. It writes
`flow-results.csv` alongside the generated PDFs, including failed exports. Use
`CLI_PROJECT` and `CLI_DLL` only when testing another CLI build.
Set `FLOW_DIR` and `SCORE_DIR` to use other output directories. The scorecard
reports visual metrics and page counts without requiring a baseline. When a reviewed
baseline is available, set `BASELINE` and run `make score-diff` separately. It marks
smoothed RMSE changes over `SCORE_TOLERANCE` (default 0.001) as improvements or
regressions, and fails on a regression or changed page count. After reviewing a
complete scorecard, `make baseline-update` can copy it to the baseline path; this
target rejects filtered or incomplete scorecards. No baseline is included here.
