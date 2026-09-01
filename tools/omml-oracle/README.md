# OMML reference oracle

This developer-only tool converts the pinned upstream OMML corpus through the
Ruby Plurimath implementation. It does not become a runtime dependency of
Docxport.Net.

Requirements: Ruby 3.0 or newer, Bundler, Git, and network access for the first
dependency restore.

From the repository root:

```powershell
pwsh tools/omml-oracle/Generate-Oracle.ps1
```

`Gemfile` pins both behavioral implementations by commit. The generated manifest
records fixture hashes, conversion failures, and oracle revisions. Generated
expectations remain distinct from reviewed files under `Normative`.
