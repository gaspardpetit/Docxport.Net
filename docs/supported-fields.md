# Supported Fields Compliance Checklist

This document tracks per-field compliance objectives for DocxportNet. Use it as a living checklist while iterating field-by-field. Mark items as done only when we have tests that demonstrate the behavior in a Word document (or a minimal synthetic document that mirrors Word output).

Legend:
- [ ] Not done
- [x] Done
- [~] Partial / in progress

## Global objectives (apply to all fields)

- [ ] End-to-end test in a Word document (or synthetic OpenXML body) that matches expected output.
- [ ] Cached result handling (cache mode) covered by test when applicable.
- [ ] CHARFORMAT and MERGEFORMAT behavior covered where relevant.
- [ ] Nested-field behavior covered where relevant.
- [ ] Regression test present (explicitly tracked).

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

Status: partial (mail-merge semantics incomplete)

- [x] \b and \f behavior when result is non-blank.
- [x] \m alias mapping.
- [x] \v vertical rendering.
- [x] General format switches (\*, \#, \@) applied.
- [ ] Mail-merge record semantics (not yet implemented).
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

- [ ] Delegate prompt + default.
- [ ] \o behavior (only-once semantics vs true merge behavior).
- [ ] No output emitted.
- [ ] Regression test present.

## SKIPIF / NEXTIF

Status: partial (no record skipping)

- [ ] Evaluates comparison correctly.
- [ ] Suppresses output.
- [ ] Record skipping / control-flow integration (not yet implemented).
- [ ] Regression test present.

## Formula (=)

Status: partial

- [ ] Core operators and comparisons.
- [ ] Supported functions (current registry).
- [ ] Table range resolution (A1, ABOVE/LEFT/RIGHT/BELOW) with TableResolver.
- [ ] Error handling parity with Word (not yet implemented).
- [ ] Additional Word functions (not yet implemented).
- [ ] Regression test present.
