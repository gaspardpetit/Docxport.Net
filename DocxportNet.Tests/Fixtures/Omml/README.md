# OMML test data

This directory deliberately keeps two kinds of expectations separate:

- `Upstream/Plurimath` is an unmodified, pinned third-party input corpus.
- `OracleGenerated` contains disposable output produced by the pinned Ruby oracle.
- `Normative` contains expectations authored and reviewed for Docxport.Net. These
  are authoritative even when they differ from the oracle.

Run `pwsh tools/omml-oracle/Generate-Oracle.ps1` from the repository root to
regenerate oracle output. Generated output is not a specification and should not
be edited by hand.
