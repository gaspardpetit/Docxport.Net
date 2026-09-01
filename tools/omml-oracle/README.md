# OMML reference oracle

This developer-only tool converts the pinned upstream OMML corpus through the
Ruby Plurimath implementation. It does not become a runtime dependency of
Docxport.Net.

Requirements: Ruby 3.0 or newer, Bundler, Git, and network access for the first
dependency restore. On Windows, native gems also require an MSYS2 UCRT toolchain.
The PowerShell wrapper automatically discovers Scoop-managed Ruby and MSYS2
installations even when their shims are not yet visible in the current session.

From the repository root:

```powershell
pwsh tools/omml-oracle/Generate-Oracle.ps1
```

`Gemfile` pins both behavioral implementations by commit. The generated manifest
records fixture hashes, conversion failures, and oracle revisions. Generated
expectations remain distinct from reviewed files under `Normative`.
