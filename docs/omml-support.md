# OMML Conversion Compliance Checklist

This document tracks the requirements for a professional-grade, standalone Office Math Markup Language (OMML) converter. It is intentionally independent from the Docxport document-walking and visitor pipeline. Pipeline integration should consume this utility later, without moving conversion rules into individual visitors.

The converter is one-way for this phase:

```text
OMML -> semantic math model -> MathML / LaTeX / UnicodeMath / basic readable text
```

Mark an item complete only when its behavior is covered by a focused unit test. Structural features should also have at least one nested/composition test. Visual properties that cannot be represented exactly in an output format must have a documented deterministic approximation.

Legend:

- [ ] Not implemented
- [x] Implemented and tested
- [~] Partial or intentionally approximate

Priority:

- P0: required for a useful first release
- P1: required for broad professional use
- P2: advanced layout, interoperability, or hardening

## Sources reviewed

- ISO/IEC 29500 / ECMA-376 Office Math, Part 1, section 22.1 (normative OMML model).
- [Microsoft Office implementation notes for `m:oMath`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/1d77457b-2884-4749-9b4a-c150ca13cc19).
- [Open XML SDK `DocumentFormat.OpenXml.Math` API](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.math).
- [MathML Core](https://www.w3.org/TR/mathml-core/) and [MathML 4](https://www.w3.org/TR/mathml4/).
- [Plurimath](https://github.com/plurimath/plurimath) at commit `00c52783877b38f6b8e6e109f1803f96bb34fc62`.
- [Plurimath OMML model and fixtures](https://github.com/plurimath/omml) at commit `51d4abe5df58fe33a92df094971c5828c3459ffb`.

Plurimath is a behavioral oracle and source of test ideas, not the definition of correctness. Its current importer deliberately drops or approximates several OMML layout properties. The standard and explicit Docxport output contracts take precedence.

## Scope and output contracts

### Standalone public surface

- [x] P0 Accept an OMML XML string containing `m:oMath` or `m:oMathPara`.
- [x] P0 Accept Open XML SDK `OfficeMath` and `DocumentFormat.OpenXml.Math.Paragraph` instances without reparsing XML.
- [x] P0 Provide expression-only LaTeX output; callers add Markdown `$` or `$$` delimiters.
- [x] P0 Provide a standalone MathML `<math>` element with the correct inline/display mode.
- [x] P0 Provide HTML-ready output using native MathML, with no document or visitor dependency.
- [x] P0 Provide UnicodeMath output as the structure-preserving plain-text representation.
- [x] P0 Provide a separate basic readable-text output for accessibility and low-fidelity consumers.
- [x] P0 Provide throwing methods with consistent exception types for malformed input.
- [x] P0 Provide `Try...` methods for non-throwing conversion.
- [x] P0 Distinguish malformed XML, unsupported valid OMML, and lossy-but-successful conversion.
- [x] P1 Return optional diagnostics identifying unsupported elements/properties and applied approximations.
- [x] P1 Support caller-selected fallback policy: throw, extract descendant text, placeholder, or omit.
- [x] P1 Expose inline/display override while defaulting from `oMath` versus `oMathPara`.
- [ ] P2 Allow custom symbol and unsupported-node handlers without exposing the internal AST.

### Architectural boundaries

- [x] P0 Keep parsing and rendering independent of `DxpWalker`, document context, and visitor types.
- [x] P0 Use one parsed semantic model for all writers.
- [x] P0 Keep output writers deterministic and culture-invariant.
- [x] P0 Preserve source child order, including repeated children of the same type.
- [x] P1 Keep the semantic model internal until a stable customization use case requires public exposure.
- [x] P1 Ensure the implementation is safe for concurrent use.
- [ ] P1 Support all Docxport target frameworks, including browser/WASM.
- [ ] P2 Permit streaming to a `TextWriter` in addition to string-returning convenience methods.

## Parsing and common semantics

### Roots, arguments, and sequences

- [ ] P0 Parse inline `m:oMath`.
- [ ] P0 Parse display `m:oMathPara` containing one or more `m:oMath` children.
- [ ] P0 Preserve adjacent expression order and intentional empty arguments.
- [ ] P0 Parse `m:e`, `m:num`, `m:den`, `m:deg`, `m:sub`, `m:sup`, `m:lim`, and `m:fName` as ordered math arguments.
- [ ] P1 Parse `m:argPr/m:argSz` and retain argument-size intent.
- [ ] P1 Apply schema defaults when property elements or `m:val` attributes are absent.
- [ ] P1 Recognize on/off lexical forms `on`, `off`, `true`, `false`, `1`, and `0` where Open XML permits them.
- [ ] P1 Handle multiple adjacent `oMath` elements consistently with Word behavior.
- [ ] P2 Detect invalid nested `oMath` and math content outside `oMath`; apply the configured recovery policy.

### Runs and tokens

- [ ] P0 Parse `m:r`, optional `m:rPr`, and `m:t` in document order.
- [ ] P0 Preserve significant spaces, non-breaking spaces, tabs, newlines, and empty text nodes.
- [ ] P0 Handle BMP and supplementary Unicode scalars without splitting surrogate pairs.
- [ ] P0 Escape XML, HTML, and LaTeX metacharacters correctly.
- [ ] P0 Classify tokens as identifiers, numbers, operators, or text for MathML.
- [ ] P0 Avoid classifying a mixed run such as `x+1` as a single MathML identifier.
- [ ] P1 Support `m:lit` literal-text behavior.
- [ ] P1 Support `m:nor` normal-text behavior.
- [ ] P1 Support math styles `m:sty`: plain (`p`), bold (`b`), italic (`i`), and bold-italic (`bi`).
- [ ] P1 Support math scripts `m:scr`: roman, script, fraktur, double-struck, sans-serif, and monospace.
- [ ] P1 Combine `m:scr` and `m:sty` into the appropriate MathML `mathvariant`, LaTeX alphabet command, and text approximation.
- [ ] P1 Support `m:aln` alignment markers.
- [ ] P1 Treat U+200B and other invisible control characters deliberately rather than accidentally emitting them.
- [ ] P2 Resolve Word symbol-font runs through `DxpFontSymbols` when sufficient font/code information exists.
- [ ] P2 Preserve language and bidirectional direction where representable.

### Manual breaks and alignment

- [ ] P1 Parse `m:brk` and its optional `m:alnAt` alignment index.
- [ ] P1 Preserve a break as a semantic line boundary in MathML, LaTeX, and text.
- [ ] P1 Support multiple breaks in a single equation.
- [ ] P1 Support breaks nested in every structure where Word emits them.
- [ ] P2 Apply document math settings `m:brkBin` (`before`, `after`, `repeat`).
- [ ] P2 Apply `m:brkBinSub` (`--`, `-+`, `+-`) when breaking at subtraction.

## OMML structures

### Fractions: `m:f` / `m:fPr`

- [ ] P0 Render numerator and denominator recursively.
- [ ] P0 Support default/bar fraction (`m:type="bar"`).
- [ ] P1 Support skewed fraction (`skw`).
- [ ] P1 Support linear fraction (`lin`).
- [ ] P1 Support no-bar stacked fraction (`noBar`).
- [ ] P1 Handle empty numerator or denominator deterministically.
- [ ] P1 Retain `m:ctrlPr` formatting intent.
- [ ] P2 Respect document-level `m:smallFrac` in inline output where representable.

### Radicals: `m:rad` / `m:radPr`

- [ ] P0 Render square root when the degree is absent or hidden.
- [ ] P0 Render an indexed root when `m:deg` is present.
- [ ] P1 Apply `m:degHide`, including missing-value defaults.
- [ ] P1 Preserve an explicitly empty degree distinctly from a missing degree.
- [ ] P1 Retain `m:ctrlPr` formatting intent.

### Scripts

- [ ] P0 Render superscript `m:sSup`.
- [ ] P0 Render subscript `m:sSub`.
- [ ] P0 Render combined subscript/superscript `m:sSubSup`.
- [ ] P1 Render pre-subscript/pre-superscript `m:sPre` using MathML multiscripts and the closest LaTeX equivalent.
- [ ] P1 Support empty base, subscript, or superscript arguments.
- [ ] P1 Apply `m:alnScr` for aligned scripts where representable.
- [ ] P1 Preserve nesting of scripts around structured bases.
- [ ] P2 Distinguish semantic operator limits from ordinary scripts.

### Delimiters: `m:d` / `m:dPr`

- [ ] P0 Render default parentheses when delimiter properties are absent.
- [ ] P0 Render explicit `m:begChr` and `m:endChr`.
- [ ] P0 Support an intentionally empty opening or closing delimiter.
- [ ] P1 Render multiple `m:e` arguments separated by `m:sepChr`.
- [ ] P1 Distinguish a missing separator from an explicitly empty separator.
- [ ] P1 Support `m:grow` stretchy delimiters.
- [ ] P1 Support delimiter shape `m:shp`: centered and match.
- [ ] P1 Map common parentheses, brackets, braces, bars, double bars, angle brackets, floors, ceilings, and white brackets.
- [ ] P1 Preserve delimiters around matrices and equation arrays.
- [ ] P2 Provide safe fallback for arbitrary Unicode delimiter characters.

### N-ary operators: `m:nary` / `m:naryPr`

- [ ] P0 Render operator, lower limit, upper limit, and operand.
- [ ] P0 Support summation, product, coproduct, integrals, contour integrals, intersections, unions, wedges, and vees.
- [ ] P1 Support arbitrary Unicode `m:chr`, with integral as the schema/Word default when absent.
- [ ] P1 Support `m:limLoc`: under/over and sub/sup.
- [ ] P1 Apply `m:subHide` and `m:supHide`.
- [ ] P1 Apply `m:grow`.
- [ ] P1 Distinguish display and inline placement.
- [ ] P2 Apply document defaults `m:intLim` and `m:naryLim` when local `m:limLoc` is absent.

### Functions: `m:func` / `m:funcPr`

- [ ] P0 Preserve function name and argument as distinct semantic values.
- [ ] P0 Render common named functions with conventional LaTeX commands where available.
- [ ] P1 Avoid treating arbitrary multi-letter identifiers as known functions without evidence.
- [ ] P1 Support structured and styled function names.
- [ ] P1 Preserve function application when the argument is empty or begins with a delimiter.
- [ ] P1 Retain `m:ctrlPr` formatting intent.

### Limits: `m:limLow`, `m:limUpp`

- [ ] P0 Render lower limits with `m:e` as the base and `m:lim` as the limit.
- [ ] P0 Render upper limits with `m:e` as the base and `m:lim` as the limit.
- [ ] P1 Recognize conventional limit/operator bases without rewriting arbitrary content.
- [ ] P1 Preserve nested accents, functions, and scripts in bases and limits.
- [ ] P1 Retain `m:ctrlPr` formatting intent.

### Accents: `m:acc` / `m:accPr`

- [ ] P0 Render the default hat when `m:chr` is absent.
- [ ] P0 Render an explicit accent character above the argument.
- [ ] P1 Map common acute, grave, hat, check, tilde, macron, breve, dot, diaeresis, and vector accents to idiomatic LaTeX.
- [ ] P1 Support arbitrary Unicode combining/accent characters via generic over-accent output.
- [ ] P1 Distinguish an accent from an ordinary overset.
- [ ] P1 Retain `m:ctrlPr` formatting intent.

### Bars: `m:bar` / `m:barPr`

- [ ] P0 Render an overbar.
- [ ] P1 Apply `m:pos="top"` and `m:pos="bot"`.
- [ ] P1 Distinguish bars from accent characters in MathML.
- [ ] P1 Retain `m:ctrlPr` formatting intent.

### Group characters: `m:groupChr` / `m:groupChrPr`

- [ ] P1 Render group characters above and below an expression.
- [ ] P1 Support explicit `m:chr` and the default group character.
- [ ] P1 Apply `m:pos` and `m:vertJc` independently.
- [ ] P1 Map overbrace, underbrace, overparen, underparen, and other common group characters to idiomatic LaTeX.
- [ ] P1 Retain `m:ctrlPr` formatting intent.

### Matrices: `m:m`, `m:mr`, `m:mPr`

- [ ] P0 Render rows and cells while preserving rectangular and ragged input.
- [ ] P0 Render nested expressions inside cells.
- [ ] P1 Apply matrix column groups `m:mcs/m:mc/m:mcPr`.
- [ ] P1 Apply column repetition `m:count`.
- [ ] P1 Apply column justification `m:mcJc`: left, center, and right.
- [ ] P1 Apply row/base justification `m:baseJc`: top, center, and bottom.
- [ ] P1 Apply placeholder visibility `m:plcHide`.
- [ ] P2 Retain row spacing `m:rSp` and `m:rSpRule`.
- [ ] P2 Retain column spacing/gap `m:cSp`, `m:cGp`, and `m:cGpRule`.
- [ ] P2 Handle inconsistent column definitions without data loss.

### Equation arrays: `m:eqArr` / `m:eqArrPr`

- [ ] P0 Render each `m:e` as a separate row.
- [ ] P0 Preserve alignment markers within rows.
- [ ] P1 Apply `m:baseJc`.
- [ ] P1 Apply `m:maxDist` and `m:objDist` where representable.
- [ ] P2 Retain `m:rSp` and `m:rSpRule`.
- [ ] P1 Produce idiomatic MathML tables and LaTeX aligned/gathered output.

### Boxes: `m:box` / `m:boxPr`

- [ ] P1 Preserve the boxed expression even when no visual box is requested.
- [ ] P1 Apply operator emulation `m:opEmu`.
- [ ] P1 Apply `m:noBreak`, `m:diff`, `m:brk`, and `m:aln` semantics.
- [ ] P1 Retain `m:ctrlPr` formatting intent.
- [ ] P2 Document output-specific approximations for operator emulation and differential spacing.

### Border boxes: `m:borderBox` / `m:borderBoxPr`

- [ ] P1 Render a four-sided box by default.
- [ ] P1 Apply `m:hideTop`, `m:hideBot`, `m:hideLeft`, and `m:hideRight` independently.
- [ ] P1 Apply horizontal and vertical strikes.
- [ ] P1 Apply bottom-left-to-top-right and top-left-to-bottom-right diagonal strikes.
- [ ] P1 Use MathML `menclose` notation values where available.
- [ ] P1 Provide deterministic LaTeX/text approximations when a border combination has no native representation.
- [ ] P1 Retain `m:ctrlPr` formatting intent.

### Phantoms: `m:phant` / `m:phantPr`

- [ ] P1 Render hidden layout content using MathML phantom/padded constructs.
- [ ] P1 Apply `m:show` and `m:transp` without silently exposing content that should be invisible.
- [ ] P1 Apply `m:zeroWid`, `m:zeroAsc`, and `m:zeroDesc` independently.
- [ ] P1 Define whether readable text includes, annotates, or omits phantom content.
- [ ] P1 Provide documented LaTeX approximations (`\phantom`, `\hphantom`, `\vphantom`, or equivalent).
- [ ] P1 Retain `m:ctrlPr` formatting intent.

## Paragraph and document math properties

### Math paragraph: `m:oMathParaPr`

- [ ] P0 Mark `oMathPara` as display math.
- [ ] P1 Support justification `m:jc`: left, right, center, and centerGroup.
- [ ] P1 Preserve multiple equations in one math paragraph.
- [ ] P1 Preserve relative alignment points across equations.

### Document math settings: `m:mathPr`

These settings normally come from the DOCX settings part. The standalone API should accept them through an optional context/options object rather than requiring a document.

- [ ] P2 Support `m:mathFont` as a formatting hint.
- [ ] P2 Support `m:brkBin` and `m:brkBinSub`.
- [ ] P2 Support `m:smallFrac` and `m:dispDef`.
- [ ] P2 Support left/right margins and default justification.
- [ ] P2 Support pre-, post-, inter-, and intra-equation spacing.
- [ ] P2 Support wrap indent and wrap-right behavior.
- [ ] P2 Support default integral and n-ary limit placement.

## Embedded WordprocessingML

Word accepts more WordprocessingML inside `m:oMath` than the base schema clearly documents. A standalone converter must define these semantics rather than inheriting pipeline behavior accidentally.

- [ ] P1 Parse Word run properties inside `m:ctrlPr` and math runs.
- [ ] P1 Support bold, italic, color, size, font, vertical alignment, and language when meaningful to the target.
- [ ] P1 Preserve ordinary `w:t`, `w:tab`, `w:br`, `w:cr`, `w:noBreakHyphen`, and `w:softHyphen` content.
- [ ] P1 Define handling for `w:sym` and symbol fonts.
- [ ] P2 Define standalone handling for hyperlinks: preserve visible math content and optionally expose the target.
- [ ] P2 Define handling for simple and complex fields: cached result by default; evaluation remains outside this utility.
- [ ] P2 Define handling for content controls, smart tags, and custom XML: unwrap visible content by default.
- [ ] P2 Define tracked-change policy: accept, reject, preserve/annotate, or caller-selected.
- [ ] P2 Preserve visible content inside move ranges and revision containers according to that policy.
- [ ] P2 Ignore non-content range markers safely: bookmarks, comments, permissions, proofing, and custom XML ranges.
- [ ] P2 Define fallback for drawings, objects, pictures, ruby, and other unexpected run content.

## MathML writer

- [ ] P0 Emit namespace-correct XML rooted at `<math xmlns="http://www.w3.org/1998/Math/MathML">`.
- [ ] P0 Set inline/block display semantics correctly.
- [ ] P0 Ensure the HTML-ready surface is safe to embed without double-escaping text or admitting source markup.
- [ ] P0 Emit `mi`, `mn`, `mo`, and `mtext` using deterministic tokenization.
- [ ] P0 Emit `mrow` only where grouping is semantically required.
- [ ] P0 Support fractions, roots, scripts, multiscripts, fenced expressions, limits, and tables.
- [ ] P1 Prefer MathML Core-compatible constructs for browser rendering.
- [ ] P1 Use `stretchy`, `accent`, `accentunder`, `movablelimits`, and `mathvariant` correctly.
- [ ] P1 Support `menclose`, `mphantom`, and `mpadded` for advanced layout.
- [ ] P1 Preserve meaningful spacing without copying Word layout measurements blindly.
- [ ] P1 Produce XML that parses without DTDs or external entities.
- [ ] P1 Add optional semantics/annotation output only behind an explicit option.
- [ ] P2 Validate representative output in current Chromium, Firefox, and WebKit/Safari engines.

## LaTeX writer

- [ ] P0 Emit an expression without `$`, `$$`, `\(`, or `\[` wrappers.
- [ ] P0 Escape text-mode and math-mode reserved characters correctly.
- [ ] P0 Use braces conservatively so nested output is unambiguous.
- [ ] P0 Support standard fractions, roots, scripts, operators, limits, and matrices.
- [ ] P1 Emit common named functions and symbols idiomatically.
- [ ] P1 Choose deterministic environments for matrices and equation arrays.
- [ ] P1 Support accents, braces, bars, boxes, and phantoms where LaTeX has a standard equivalent.
- [ ] P1 Report required non-core packages, or restrict default output to a documented package baseline.
- [ ] P1 Never inject raw control sequences from untrusted OMML text.
- [ ] P2 Offer a compatibility profile for MathJax/KaTeX-supported LaTeX.

## Readable Unicode text writer

- [ ] P0 Preserve every visible literal, identifier, number, and operator.
- [ ] P0 Use explicit grouping to avoid ambiguous flattening.
- [ ] P0 Render fractions, roots, scripts, limits, and matrices in stable readable notation.
- [ ] P1 Use Unicode super/subscript characters only when the complete value is representable; otherwise use `^(...)` and `_(...)`.
- [ ] P1 Render multi-line structures with a caller-selected single-line or multi-line policy.
- [ ] P1 Define accessible names for invisible or purely visual constructs.
- [ ] P1 Avoid dependence on terminal width or current culture.
- [ ] P2 Define the supported UnicodeMath version/profile and document intentional deviations.

## Error handling, security, and quality

- [ ] P0 Reject malformed XML with a clear input exception.
- [ ] P0 Never silently discard an unknown semantic node.
- [ ] P0 Preserve descendant visible text in the default unsupported-node fallback.
- [ ] P0 Include the unsupported element name/path in diagnostics.
- [ ] P1 Disable DTD processing and external entity resolution.
- [ ] P1 Bound input size, nesting depth, and output growth, or document caller-enforced limits.
- [ ] P1 Avoid recursive stack exhaustion on deeply nested input.
- [ ] P1 Handle missing required children with explicit recovery rules.
- [ ] P1 Handle duplicate singleton properties deterministically.
- [ ] P1 Handle unknown attributes and future-version extension elements without losing visible content.
- [ ] P1 Preserve namespace correctness regardless of the source prefix.
- [ ] P1 Test alternate prefixes and default namespaces.
- [ ] P1 Test null, empty, whitespace-only, and non-OMML XML inputs.
- [ ] P1 Test deterministic output across cultures and line-ending conventions.
- [ ] P1 Test concurrent conversions.
- [ ] P2 Benchmark large equations and documents containing thousands of equations.
- [ ] P2 Fuzz malformed and adversarial XML inputs.

## Plurimath fixture audit

At OMML commit `51d4abe5df58fe33a92df094971c5828c3459ffb`, the repository contains 279 `.omml` fixtures:

- 189 general fixtures in `spec/fixtures/omml`.
- 90 fixtures in `spec/fixtures/omml/line_break`.

The corpus is valuable for complex composition and real-world regression coverage. It is not a complete property matrix.

### Structures represented in the corpus

- [x] Import/recreate and attribute the general fixture corpus.
- [x] Import/recreate and attribute the line-break fixture corpus.
- [ ] Fractions: fixtures 001-004 and complex compositions.
- [ ] Scripts: fixtures 005-008 plus later nested cases.
- [ ] Radicals: fixtures 009-012.
- [ ] N-ary operators: fixtures 013-046 plus complex compositions.
- [ ] Delimiters and piecewise expressions: fixtures 045-072.
- [ ] Functions and powers: fixtures 073-099.
- [ ] Limits, accents, bars, and border boxes: fixtures 100-155.
- [ ] Matrices and equation arrays: fixtures 156-177.
- [ ] Mixed complex expressions: fixtures 178 onward and `issue-158.omml`.
- [ ] Group characters: fixtures 185-186.
- [ ] Manual breaks across many parent structures: `line_break/*`.
- [ ] Math styles and scripts: plain/bold/italic/bold-italic plus script, fraktur, double-struck, sans-serif, and monospace examples.
- [ ] Delimiter variants: parentheses, brackets, braces, bars, double bars, angles, floors, ceilings, and white brackets.
- [ ] Matrix column counts and centered columns.
- [ ] Hidden n-ary subscript/superscript flags and both limit locations.

### Important corpus gaps requiring Docxport fixtures

- [ ] `m:box` and every `m:boxPr` behavior.
- [ ] Fraction types `skw`, `lin`, and `noBar` (the reviewed fixtures do not exercise `m:type`).
- [ ] Delimiter shape `m:shp` and explicit grow-off behavior.
- [ ] Equation-array spacing, distance, and base-justification properties.
- [ ] Matrix left/right column justification, base justification, placeholders, and spacing/gap properties.
- [ ] Border-box horizontal, vertical, and diagonal strike combinations.
- [ ] Phantom zero-width/ascent/descent and transparency combinations.
- [ ] `m:argSz`, `m:lit`, `m:nor`, and `m:aln` semantics.
- [ ] Math paragraph justification and multiple `oMath` children.
- [ ] Document-level `m:mathPr` defaults.
- [ ] Embedded hyperlinks, fields, content controls, custom XML, and tracked revisions.
- [ ] Alternate namespace prefixes and default math namespace.
- [ ] Malformed XML, missing required arguments, duplicate properties, and unknown elements.
- [ ] Supplementary-plane Unicode and combining sequences.
- [ ] All output-injection and resource-limit security cases.

### Oracle workflow

- [x] Record upstream repository, commit, path, and license with reused fixtures.
- [x] Generate MathML, LaTeX, and UnicodeMath oracle outputs using a pinned Plurimath version.
- [x] Store generated oracle outputs separately from hand-authored normative expectations.
- [x] Canonicalize XML before comparing MathML; do not compare prefixes or insignificant whitespace.
- [ ] Treat an oracle disagreement as a review prompt, not automatic proof that Docxport is wrong.
- [ ] Add focused named fixtures for every discovered regression; numbered corpus files alone are difficult to diagnose.

## Findings from the Plurimath implementation review

The reviewed importer provides useful implementation lessons that should become explicit requirements:

- [ ] Preserve original element order rather than iterating model properties.
- [ ] Make unsupported typed nodes fail or diagnose explicitly instead of falling through.
- [ ] Distinguish empty content from absent content.
- [ ] Normalize non-breaking spaces and entities without accepting malformed XML silently.
- [ ] Keep token resolution contextual: text, operator, function name, accent, and delimiter are not interchangeable.
- [ ] Do not bind a unary/ternary function to adjacent content without clear syntactic evidence.
- [ ] Handle default characters explicitly: parentheses, integral, hat, and group characters.
- [ ] Treat styled runs as content plus style, not as reordered sibling values.
- [ ] Avoid quadratic behavior when consuming repeated sibling elements.

The following Plurimath behaviors are intentionally not sufficient as our target:

- [ ] Preserve fraction type; the reviewed importer maps every fraction to a conventional fraction.
- [ ] Preserve bar position; the reviewed importer maps every bar to the same bar function.
- [ ] Preserve border-side and strike properties; the reviewed importer maps every border box to one `menclose` form.
- [ ] Preserve box semantics; the reviewed importer unwraps boxes.
- [ ] Preserve phantom layout semantics; the reviewed importer unwraps phantoms.
- [ ] Preserve matrix/equation-array layout properties; the reviewed importer primarily retains rows and cells.
- [ ] Respect `degHide`, rather than deciding square root solely from an empty degree.
- [ ] Apply delimiter grow/shape behavior, not only boundary and separator characters.
- [ ] Carry diagnostics for approximations rather than silently simplifying them.

## Goal-by-goal implementation sequence

Complete these goals in order. Goals 3-10 are vertical feature slices: each includes parsing/model work, MathML, LaTeX, UnicodeMath and basic-text output, focused tests, nested/composition tests, and all applicable upstream corpus tests. Do not defer a feature's writers to a later goal.

### Goal 1: Oracle and test harness

- [x] Import or recreate the pinned Plurimath fixtures with attribution.
- [x] Pin the oracle version and record a reproducible generation command.
- [x] Canonicalize MathML and normalize only insignificant output differences.
- [x] Separate generated oracle expectations from normative hand-authored expectations.
- [x] Establish helpers for focused OMML fragments, nested expressions, and malformed input.

### Goal 2: Standalone API and parser foundation

- [x] Implement the standalone public conversion surface and options.
- [x] Implement secure XML parsing, OMML root validation, ordered sequences, and the initial semantic model.
- [x] Implement consistent exceptions, `Try...` methods, diagnostics, and fallback policies.
- [x] Establish deterministic, culture-invariant MathML, LaTeX, UnicodeMath, basic-text, and HTML-ready writers.
- [x] Confirm the architecture has no dependency on the document walker or visitors.

### Goal 3: Runs, tokens, symbols, and styling

- [ ] Implement math runs, text, whitespace, Unicode scalar handling, and token classification.
- [ ] Implement literal/normal text, math style, math script, escaping, and applicable Word run properties.
- [ ] Integrate symbol-font translation where sufficient information exists.
- [ ] Cover simple-run, mixed-token, styled-run, invisible-character, and supplementary-Unicode cases.

### Goal 4: Fractions, radicals, and scripts

- [ ] Implement every fraction type, including bar, skewed, linear, and no-bar.
- [ ] Implement square and indexed roots, degree hiding, and empty-degree distinctions.
- [ ] Implement subscript, superscript, combined scripts, pre-scripts, and script alignment.
- [ ] Cover structured and deeply nested bases, degrees, numerators, denominators, and scripts.

### Goal 5: Delimiters and decorations

- [ ] Implement delimiters, repeated arguments, separators, empty boundaries, growth, and shape.
- [ ] Implement accents, bars above/below, and group characters.
- [ ] Cover common and arbitrary Unicode delimiters and accents.
- [ ] Cover decorated matrices and nested decorated expressions.

### Goal 6: Functions, limits, and n-ary operators

- [ ] Implement named and arbitrary functions without unsafe semantic guessing.
- [ ] Implement upper and lower limits.
- [ ] Implement standard and arbitrary n-ary operators, hidden limits, growth, and limit placement.
- [ ] Cover differences between ordinary scripts and operator limits in inline and display math.

### Goal 7: Matrices and equation arrays

- [ ] Implement matrix rows/cells, column groups/counts, alignment, placeholders, spacing, and gaps.
- [ ] Implement equation-array rows, alignment points, justification, distances, and spacing.
- [ ] Handle ragged and inconsistent input without losing cell content.
- [ ] Produce idiomatic MathML tables, LaTeX environments, and readable text.

### Goal 8: Boxes, border boxes, and phantoms

- [ ] Implement box operator, break, alignment, differential, and no-break semantics.
- [ ] Implement every border-side and horizontal/vertical/diagonal strike combination.
- [ ] Implement phantom visibility, transparency, zero-width, zero-ascent, and zero-descent behavior.
- [ ] Document and diagnose every output-specific approximation.

### Goal 9: Breaks and equation layout

- [ ] Implement manual breaks and alignment indices across every supported parent structure.
- [ ] Implement math-paragraph justification, multiple equations, and relative alignment.
- [ ] Implement optional document math settings and their local-property precedence.
- [ ] Cover all dedicated upstream line-break fixtures.

### Goal 10: Embedded WordprocessingML

- [ ] Implement visible Word run content and formatting inside math.
- [ ] Define and implement standalone policies for fields, hyperlinks, content controls, smart tags, and custom XML.
- [ ] Define and implement tracked-change handling independently from the document pipeline.
- [ ] Safely unwrap or diagnose bookmarks, comments, range markers, drawings, objects, and unexpected content.

### Goal 11: Corpus conformance and gap closure

- [ ] Run the complete pinned general and line-break corpora.
- [ ] Resolve every crash, silent-content loss, and unexplained oracle difference.
- [ ] Add named Docxport-owned fixtures for every corpus gap listed above.
- [ ] Verify every normative OMML structure and property has a test or documented intentional limitation.

### Goal 12: Production hardening

- [ ] Complete malformed, adversarial, unknown-extension, namespace, and resource-limit tests.
- [ ] Complete concurrency, culture, line-ending, performance, and fuzz testing.
- [ ] Validate all target frameworks, including browser/WASM.
- [ ] Validate representative MathML in supported browser engines and LaTeX under the documented profile.
- [ ] Complete public API and compatibility documentation.

### Goal 13: Pipeline integration (later and separate)

- [ ] Add thin integration into the HTML visitor using the standalone HTML-ready/MathML surface.
- [ ] Add thin integration into the Markdown visitor using expression-only LaTeX plus visitor-owned delimiters.
- [ ] Add thin integration into the plain-text visitor using the standalone readable-text surface.
- [ ] Confirm standalone and pipeline outputs use the same conversion behavior.

## Workflow for each goal

Use the following workflow for each goal above:

1. [ ] Select the next incomplete goal; do not mix work from later goals unless it is a necessary shared prerequisite.
2. [ ] Review the applicable ECMA-376/ISO 29500 requirements, Microsoft implementation notes, Open XML SDK model, and target-format specifications.
3. [ ] Review the pinned Plurimath OMML model, importer, writers, and related fixtures for guidance and known edge cases.
4. [ ] Record any newly discovered requirement or corpus gap in this checklist before or alongside implementation.
5. [ ] Implement the parser/model behavior and all applicable output writers as one vertical slice.
6. [ ] Add attributed upstream fixtures and focused synthetic tests for uncovered defaults, properties, malformed input, and compositions.
7. [ ] Run focused tests while developing, then the complete test suite for all supported target frameworks applicable to the change.
8. [ ] Review the goal as a package: public API, parser/model, every writer, diagnostics, tests, documentation, and interaction with previously completed goals.
9. [ ] Confirm no visible content is silently lost and every approximation is deterministic and diagnosable.
10. [ ] Update this checklist only for behavior demonstrated by passing tests.
11. [ ] Work on a dedicated development branch and review the final diff for unrelated or generated changes.
12. [ ] Commit only the completed goal's files, preserving unrelated working-tree changes; use a commit message that identifies the OMML goal.

### Completion gate for goals 3-10

A feature goal is complete only when all of the following are true:

- [ ] Its normative elements, properties, defaults, and invalid-input behavior have been reviewed.
- [ ] Parsing and the semantic model preserve the information required by every output.
- [ ] MathML, LaTeX, UnicodeMath, basic text, and HTML-ready behavior where distinct are implemented.
- [ ] Focused, nested/composition, upstream-corpus, and identified-gap tests pass.
- [ ] The complete repository test suite passes.
- [ ] Lossy mappings are documented and produce the agreed diagnostics.
- [ ] The integrated result has been reviewed against all previously completed goals.
- [ ] The checklist and public documentation reflect the tested behavior.

## Definition of professional-grade completion

- [ ] Every normative OMML structure has a tested semantic conversion or a documented intentional limitation.
- [ ] Every property above has a test proving preservation, approximation, or an explicit diagnostic.
- [ ] All imported corpus fixtures convert without crashes or silent visible-content loss.
- [ ] Every corpus gap above has a focused Docxport-owned fixture.
- [ ] MathML is namespace-valid and renders acceptably in supported browsers.
- [ ] LaTeX compiles under the documented baseline and/or renders under the selected MathJax/KaTeX profile.
- [ ] Text output remains understandable when copied without styling.
- [ ] Malformed, adversarial, and future-version input fails safely.
- [ ] Public API behavior, defaults, diagnostics, and compatibility guarantees are documented.
- [ ] Pipeline integration can call the standalone API without special conversion branches.
