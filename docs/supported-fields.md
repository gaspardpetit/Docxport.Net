# Supported Fields Compliance Checklist

This document tracks per-field compliance objectives for DocxportNet. Use it as a living checklist while iterating field-by-field. Mark items as done only when we have tests that demonstrate the behavior in a Word document (or a minimal synthetic document that mirrors Word output).

Legend:
- [ ] Not done
- [x] Done
- [~] Partial / in progress

## Global objectives (apply to all fields)

- [x] End-to-end test in a Word document (or synthetic OpenXML body) that matches expected output.
- [x] Cached result handling (cache mode) covered by test when applicable.
- [x] CHARFORMAT and MERGEFORMAT behavior covered where relevant.
- [x] Nested-field behavior covered where relevant.
- [x] Regression test present (explicitly tracked).

---

## IF

Status: feature-complete

- [x] Structured branch replay (runs + formatting).
- [x] Nested field evaluation inside expressions.
- [x] Wildcards and numeric/string comparisons.
- [x] Error propagation in comparisons (uses nested field text).
- [x] Regression test present.

## SET

Status: feature-complete

- [x] Sets bookmark value; emits no output.
- [x] Nested-field value for SET.
- [x] Cache mode suppresses SET output.
- [x] Regression test present.

## REF

Status: feature-complete (review multi-run bookmarks)

- [x] Resolves bookmark value.
- [x] REF switches: \d \f \h \n \p \r \t \w (via resolver).
- [x] CHARFORMAT/MERGEFORMAT for rendered output.
- [x] Multi-run/structured bookmark replay.
- [x] Regression test present.

## DOCVARIABLE

Status: feature-complete (edge-case hardening optional)

- [x] Resolver path (delegate / value resolver).
- [x] Error text when missing.
- [x] CHARFORMAT/MERGEFORMAT for rendered output.
- [x] Regression test present.

## DATE / TIME / CREATEDATE / SAVEDATE / PRINTDATE

Status: feature-complete

- [x] Uses NowProvider for DATE/TIME.
- [x] Uses document properties for CREATEDATE/SAVEDATE/PRINTDATE (with fallback).
- [x] \@ formatting coverage.
- [x] Regression test present.

## DOCPROPERTY

Status: feature-complete (custom props optional)

- [x] Built-in core properties resolve.
- [x] Custom properties resolve (includeCustomProperties = true).
- [x] \* formatting coverage.
- [x] Regression test present.

## MERGEFIELD

Status: feature-complete (mail-merge semantics implemented)

- [x] \b and \f behavior when result is non-blank.
- [x] \m alias mapping.
- [x] \v vertical rendering.
- [x] General format switches (\*, \#, \@) applied.
- [x] Mail-merge record semantics (record cursor + NEXT/NEXTIF/SKIPIF integration).
- [x] Regression test present.

## SEQ

Status: partial

- [x] Increment behavior.
- [x] \c repeat.
- [x] \r reset.
- [x] \s heading level reset.
- [x] \h hide (and interaction with \*).
- [x] General format switches (\*, \#).
- [x] Regression test present.

## COMPARE

Status: feature-complete

- [x] Returns 1/0 and applies formatting.
- [x] Nested field evaluation.
- [x] Regression test present.

## ASK

Status: partial (delegate-backed)

- [x] Delegate prompt + default.
- [x] \o behavior (only-once semantics vs true merge behavior).
- [x] No output emitted.
- [x] Regression test present.

## SKIPIF / NEXTIF

Status: feature-complete

- [x] Evaluates comparison correctly.
- [x] Suppresses output.
- [x] Record skipping / control-flow integration (merge cursor).
- [x] Regression test present.

## Formula (=)

Status: partial

- [x] Core operators and comparisons.
- [x] Supported functions (current registry).
- [x] Table range resolution (A1, ABOVE/LEFT/RIGHT/BELOW) with TableResolver.
- [x] Error handling parity with Word (not yet implemented).
- [ ] Additional Word functions (not yet implemented).
- [x] Regression test present.

---

## NEXT

Status: feature-complete

- [x] Advances merge record cursor unconditionally.
- [x] No output emitted.
- [x] Regression test present.

## MERGEREC / MERGESEQ

Status: feature-complete

- [x] MERGEREC returns current record number from data source.
- [x] MERGESEQ returns sequence number of merged records.
- [x] Regression test present.

## Document Metrics (NUMPAGES / NUMWORDS / NUMCHARS)

Status: not tracked

- [x] Resolves from document stats.
- [ ] Compute on the fly when stats missing.
- [x] Formatting switches apply where relevant.
- [x] Regression test present.

## PAGEREF / NOTEREF

Status: out of scope (requires pagination/layout engine)

- [ ] Resolves page number for bookmark (PAGEREF).
- [ ] Resolves note reference mark for bookmark (NOTEREF).
- [ ] Regression test present.

## GREETINGLINE

Status: partial (locale-aware providers; simple templates)

- [x] Provider-based macro resolution (locale-aware registry).
- [x] Uses merge cursor values.
- [x] Regression test present.
- [~] Locale-specific templates (may differ from Word labels; documented intentional divergence).

## ADDRESSBLOCK

Status: partial (locale-aware providers; simple templates)

- [x] Provider-based macro resolution (locale-aware registry).
- [x] Uses merge cursor values.
- [x] Regression test present.
- [~] Locale-specific templates (may differ from Word labels; documented intentional divergence).

## DATABASE

Status: partial (pluggable provider; basic TSV rendering)

- [x] Pluggable database provider interface (optional external providers).
- [ ] Default provider (e.g., SqlClient / T-SQL).
- [ ] Optional providers (ODBC, PostgreSQL, MySQL).
- [ ] External data query support (document-level or provider-driven).
- [ ] Mapping to output rows/records.
- [ ] Configurable rendering (table/HTML/Markdown).
- [ ] Regression test present.

## INCLUDETEXT / INCLUDEPICTURE

Status: not tracked

- [ ] External content resolution (file/URL).
- [ ] Security/sandbox policy for external loads.
- [ ] Regression test present.

## TOC / TC / INDEX / XE / RD

Status: out of scope (requires indexing + layout engine).
Note: Full support needs a two-pass indexer (collect entries first, render later) plus pagination to compute page numbers. A possible interim “TOC-lite” would link to headings/bookmarks without page numbers.

- [ ] TOC aggregation and rendering (with page numbers).
- [ ] TC entry capture for TOC.
- [ ] INDEX aggregation and rendering (with page numbers).
- [ ] XE entry capture for index.
- [ ] RD external document inclusion.
- [ ] Optional TOC-lite (hyperlinks only, no page numbers).
- [ ] Regression test present.

## Field Error Semantics

Status: partial

- [x] Word-style error strings for REF/DOCVARIABLE/DOCPROPERTY missing values.
- [x] Formula error strings for divide-by-zero and syntax/unknown functions.
- [ ] Word-style error strings for other field types (e.g., SEQ, MERGEFIELD, ASK).
- [~] Nested field error propagation behavior (known divergence: missing REF inside IF).
- [~] Regression tests for error strings across all fields.
