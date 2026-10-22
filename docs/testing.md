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
