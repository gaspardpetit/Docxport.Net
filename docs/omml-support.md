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
- [x] P1 Support all Docxport target frameworks, including browser/WASM.
- [ ] P2 Permit streaming to a `TextWriter` in addition to string-returning convenience methods.

## Parsing and common semantics

### Roots, arguments, and sequences

- [x] P0 Parse inline `m:oMath`.
- [x] P0 Parse display `m:oMathPara` containing one or more `m:oMath` children.
- [x] P0 Preserve adjacent expression order and intentional empty arguments.
- [x] P0 Parse `m:e`, `m:num`, `m:den`, `m:deg`, `m:sub`, `m:sup`, `m:lim`, and `m:fName` as ordered math arguments.
- [x] P1 Parse `m:argPr/m:argSz` and retain argument-size intent.
- [x] P1 Apply schema defaults when property elements or `m:val` attributes are absent.
- [x] P1 Recognize on/off lexical forms `on`, `off`, `true`, `false`, `1`, and `0` where Open XML permits them.
- [x] P1 Handle multiple adjacent `oMath` elements consistently with Word behavior.
- [x] P2 Detect invalid nested `oMath` and math content outside `oMath`; apply the configured recovery policy.

`m:argSz` is retained only on the Word-supported argument pairs: box and
group-character bases; lower/upper limits; n-ary sub/superscripts; radical
degrees; and ordinary/pre-script sub/superscripts. Its `-2` through `2` value is
relative, with absent or valueless properties defaulting to zero. MathML maps it
to relative `scriptlevel`; LaTeX uses its nearest standard math style and emits
`OMML002`; UnicodeMath and readable text preserve content and diagnose sizing.

### Runs and tokens

- [x] P0 Parse `m:r`, optional `m:rPr`, and `m:t` in document order.
- [x] P0 Preserve significant spaces, non-breaking spaces, tabs, newlines, and empty text nodes.
- [x] P0 Handle BMP and supplementary Unicode scalars without splitting surrogate pairs.
- [x] P0 Escape XML, HTML, and LaTeX metacharacters correctly.
- [x] P0 Classify tokens as identifiers, numbers, operators, or text for MathML.
- [x] P0 Avoid classifying a mixed run such as `x+1` as a single MathML identifier.
- [x] P1 Support `m:lit` literal-text behavior.
- [x] P1 Support `m:nor` normal-text behavior.
- [x] P1 Support math styles `m:sty`: plain (`p`), bold (`b`), italic (`i`), and bold-italic (`bi`).
- [x] P1 Support math scripts `m:scr`: roman, script, fraktur, double-struck, sans-serif, and monospace.
- [x] P1 Combine `m:scr` and `m:sty` into the appropriate MathML `mathvariant`, LaTeX alphabet command, and text approximation.
- [x] P1 Support `m:aln` alignment markers.
- [x] P1 Treat U+200B and other invisible control characters deliberately rather than accidentally emitting them.
- [x] P2 Resolve Word symbol-font runs through `DxpFontSymbols` when sufficient font/code information exists.
- [x] P2 Preserve language and bidirectional direction where representable.

### Manual breaks and alignment

- [x] P1 Parse `m:brk` and its optional `m:alnAt` alignment index.
- [x] P1 Preserve a break as a semantic line boundary in MathML, LaTeX, and text.
- [x] P1 Support multiple breaks in a single equation.
- [x] P1 Support breaks nested in every structure where Word emits them.

A run or box `m:brk` is represented by a MathML `mspace` line-break marker;
its numeric `alnAt` target is retained as `data-omml-align-at`. LaTeX places
the smallest containing sequence in an `aligned` environment, while
UnicodeMath and readable text use a newline. Numeric operator targeting has no
portable LaTeX or plain-text equivalent and produces an `OMML002` warning.
Paragraph-level `w:br`/`w:cr` boundaries between `m:oMath` children become
MathML `mtable` rows and LaTeX `aligned` rows.
- [x] P2 Apply document math settings `m:brkBin` (`before`, `after`, `repeat`).
- [x] P2 Apply `m:brkBinSub` (`--`, `-+`, `+-`) when breaking at subtraction.

## OMML structures

### Fractions: `m:f` / `m:fPr`

- [x] P0 Render numerator and denominator recursively.
- [x] P0 Support default/bar fraction (`m:type="bar"`).
- [x] P1 Support skewed fraction (`skw`).
- [x] P1 Support linear fraction (`lin`).
- [x] P1 Support no-bar stacked fraction (`noBar`).
- [x] P1 Handle empty numerator or denominator deterministically.
- [x] P1 Retain `m:ctrlPr` formatting intent.
- [x] P2 Respect document-level `m:smallFrac` in inline output where representable.

### Radicals: `m:rad` / `m:radPr`

- [x] P0 Render square root when the degree is absent or hidden.
- [x] P0 Render an indexed root when `m:deg` is present.
- [x] P1 Apply `m:degHide`, including missing-value defaults.
- [x] P1 Preserve an explicitly empty degree distinctly from a missing degree.
- [x] P1 Retain `m:ctrlPr` formatting intent.

### Scripts

- [x] P0 Render superscript `m:sSup`.
- [x] P0 Render subscript `m:sSub`.
- [x] P0 Render combined subscript/superscript `m:sSubSup`.
- [x] P1 Render pre-subscript/pre-superscript `m:sPre` using MathML multiscripts and the closest LaTeX equivalent.
- [x] P1 Support empty base, subscript, or superscript arguments.
- [x] P1 Apply `m:alnScr` for aligned scripts where representable.
- [x] P1 Preserve nesting of scripts around structured bases.
- [x] P2 Distinguish semantic operator limits from ordinary scripts.

### Delimiters: `m:d` / `m:dPr`

- [x] P0 Render default parentheses when delimiter properties are absent.
- [x] P0 Render explicit `m:begChr` and `m:endChr`.
- [x] P0 Support an intentionally empty opening or closing delimiter.
- [x] P1 Render multiple `m:e` arguments separated by `m:sepChr`.
- [x] P1 Distinguish a missing separator from an explicitly empty separator.
- [x] P1 Support `m:grow` stretchy delimiters.
- [x] P1 Support delimiter shape `m:shp`: centered and match.
- [x] P1 Map common parentheses, brackets, braces, bars, double bars, angle brackets, floors, ceilings, and white brackets.
- [x] P1 Preserve delimiters around matrices and equation arrays.
- [x] P2 Provide safe fallback for arbitrary Unicode delimiter characters.

### N-ary operators: `m:nary` / `m:naryPr`

- [x] P0 Render operator, lower limit, upper limit, and operand.
- [x] P0 Support summation, product, coproduct, integrals, contour integrals, intersections, unions, wedges, and vees.
- [x] P1 Support arbitrary Unicode `m:chr`, with integral as the schema/Word default when absent.
- [x] P1 Support `m:limLoc`: under/over and sub/sup.
- [x] P1 Apply `m:subHide` and `m:supHide`.
- [x] P1 Apply `m:grow`.
- [x] P1 Distinguish display and inline placement.
- [x] P2 Apply document defaults `m:intLim` and `m:naryLim` when local `m:limLoc` is absent.

### Functions: `m:func` / `m:funcPr`

- [x] P0 Preserve function name and argument as distinct semantic values.
- [x] P0 Render common named functions with conventional LaTeX commands where available.
- [x] P1 Avoid treating arbitrary multi-letter identifiers as known functions without evidence.
- [x] P1 Support structured and styled function names.
- [x] P1 Preserve function application when the argument is empty or begins with a delimiter.
- [x] P1 Retain `m:ctrlPr` formatting intent.

### Limits: `m:limLow`, `m:limUpp`

- [x] P0 Render lower limits with `m:e` as the base and `m:lim` as the limit.
- [x] P0 Render upper limits with `m:e` as the base and `m:lim` as the limit.
- [x] P1 Recognize conventional limit/operator bases without rewriting arbitrary content.
- [x] P1 Preserve nested accents, functions, and scripts in bases and limits.
- [x] P1 Retain `m:ctrlPr` formatting intent.

### Accents: `m:acc` / `m:accPr`

- [x] P0 Render the default hat when `m:chr` is absent.
- [x] P0 Render an explicit accent character above the argument.
- [x] P1 Map common acute, grave, hat, check, tilde, macron, breve, dot, diaeresis, and vector accents to idiomatic LaTeX.
- [x] P1 Support arbitrary Unicode combining/accent characters via generic over-accent output.
- [x] P1 Distinguish an accent from an ordinary overset.
- [x] P1 Retain `m:ctrlPr` formatting intent.

### Bars: `m:bar` / `m:barPr`

- [x] P0 Render an overbar.
- [x] P1 Apply `m:pos="top"` and `m:pos="bot"`.
- [x] P1 Distinguish bars from accent characters in MathML.
- [x] P1 Retain `m:ctrlPr` formatting intent.

### Group characters: `m:groupChr` / `m:groupChrPr`

- [x] P1 Render group characters above and below an expression.
- [x] P1 Support explicit `m:chr` and the default group character.
- [x] P1 Apply `m:pos` and `m:vertJc` independently.
- [x] P1 Map overbrace, underbrace, overparen, underparen, and other common group characters to idiomatic LaTeX.
- [x] P1 Retain `m:ctrlPr` formatting intent.

### Matrices: `m:m`, `m:mr`, `m:mPr`

- [x] P0 Render rows and cells while preserving rectangular and ragged input.
- [x] P0 Render nested expressions inside cells.
- [x] P1 Apply matrix column groups `m:mcs/m:mc/m:mcPr`.
- [x] P1 Apply column repetition `m:count`.
- [x] P1 Apply column justification `m:mcJc`: left, center, and right.
- [x] P1 Apply row/base justification `m:baseJc`: top, center, and bottom.
- [x] P1 Apply placeholder visibility `m:plcHide`.
- [x] P2 Retain row spacing `m:rSp` and `m:rSpRule`.
- [x] P2 Retain column spacing/gap `m:cSp`, `m:cGp`, and `m:cGpRule`.
- [x] P2 Handle inconsistent column definitions without data loss.

### Equation arrays: `m:eqArr` / `m:eqArrPr`

- [x] P0 Render each `m:e` as a separate row.
- [x] P0 Preserve alignment markers within rows.
- [x] P1 Apply `m:baseJc`.
- [x] P1 Apply `m:maxDist` and `m:objDist` where representable.
- [x] P2 Retain `m:rSp` and `m:rSpRule`.
- [x] P1 Produce idiomatic MathML tables and LaTeX aligned/gathered output.

MathML uses `mtable`, retains OMML-only layout values in `data-omml-*`
attributes, and represents a visible empty-cell placeholder with a zero-width
`mspace` marker rather than inventing mathematical content. `maxDist` maps to a
full-width table; `objDist` is retained as metadata because MathML has no exact
equivalent. LaTeX uses `array` for matrices and `aligned` or `gathered` for
equation arrays. Word's numeric row/column spacing rules have no exact portable
LaTeX equivalent and therefore remain in the semantic model and MathML metadata.

### Boxes: `m:box` / `m:boxPr`

- [x] P1 Preserve the boxed expression even when no visual box is requested.
- [x] P1 Apply operator emulation `m:opEmu`.
- [x] P1 Apply `m:noBreak`, `m:diff`, `m:brk`, and `m:aln` semantics.
- [x] P1 Retain `m:ctrlPr` formatting intent.
- [x] P2 Document output-specific approximations for operator emulation and differential spacing.

### Border boxes: `m:borderBox` / `m:borderBoxPr`

- [x] P1 Render a four-sided box by default.
- [x] P1 Apply `m:hideTop`, `m:hideBot`, `m:hideLeft`, and `m:hideRight` independently.
- [x] P1 Apply horizontal and vertical strikes.
- [x] P1 Apply bottom-left-to-top-right and top-left-to-bottom-right diagonal strikes.
- [x] P1 Use MathML `menclose` notation values where available.
- [x] P1 Provide deterministic LaTeX/text approximations when a border combination has no native representation.
- [x] P1 Retain `m:ctrlPr` formatting intent.

### Phantoms: `m:phant` / `m:phantPr`

- [x] P1 Render hidden layout content using MathML phantom/padded constructs.
- [x] P1 Apply `m:show` and `m:transp` without silently exposing content that should be invisible.
- [x] P1 Apply `m:zeroWid`, `m:zeroAsc`, and `m:zeroDesc` independently.
- [x] P1 Define whether readable text includes, annotates, or omits phantom content.
- [x] P1 Provide documented LaTeX approximations (`\phantom`, `\hphantom`, `\vphantom`, or equivalent).
- [x] P1 Retain `m:ctrlPr` formatting intent.

MathML represents border boxes with `menclose`, hidden phantom content with
`mphantom`, and independently suppressed dimensions with `mpadded`. OMML
operator-emulation, no-break, differential, and phantom-transparency intent is
also retained in `data-omml-*` attributes when MathML has no exact equivalent.

LaTeX uses `\boxed` for a plain four-sided border and the MathJax/KaTeX
`\enclose` extension for arbitrary side and strike combinations. A shown,
zero-width phantom uses `\mathrlap`, which requires `mathtools`; other phantom
forms use `\phantom`, `\hphantom`, `\vphantom`, and `\smash`. A box break is
placed in the smallest containing LaTeX `aligned` environment.
UnicodeMath uses the standard rectangle mask and phantom/smash operators used
by Plurimath. Readable text describes border notation, includes shown phantom
content, and omits hidden phantom content. Every non-portable approximation
produces an `OMML002` warning.

## Paragraph and document math properties

### Math paragraph: `m:oMathParaPr`

- [x] P0 Mark `oMathPara` as display math.
- [x] P1 Support justification `m:jc`: left, right, center, and centerGroup.
- [x] P1 Preserve multiple equations in one math paragraph.
- [x] P1 Preserve relative alignment points across equations.

### Document math settings: `m:mathPr`

These settings normally come from the DOCX settings part. The standalone API should accept them through an optional context/options object rather than requiring a document.

- [x] P2 Support `m:mathFont` as a formatting hint.
- [x] P2 Support `m:brkBin` and `m:brkBinSub`.
- [x] P2 Support `m:smallFrac` and `m:dispDef`.
- [x] P2 Support left/right margins and default justification.
- [x] P2 Support pre-, post-, inter-, and intra-equation spacing.
- [x] P2 Support wrap indent and wrap-right behavior.
- [x] P2 Support default integral and n-ary limit placement.

Standalone callers provide document math settings through
`DxpOmmlConversionOptions`: font and binary-break hints; compact fractions and
display defaults; justification; margins and spacing in twips; wrapping; and
integral/n-ary limit placement. A local `m:oMathParaPr/m:jc` overrides
`DefaultJustification`, and `WrapRight` overrides the schema-alternative
`WrapIndentTwips`. MathML retains settings without native equivalents as
`data-omml-*` metadata; textual formats report those layout approximations.

## Embedded WordprocessingML

Word accepts more WordprocessingML inside `m:oMath` than the base schema clearly documents. A standalone converter must define these semantics rather than inheriting pipeline behavior accidentally.

- [x] P1 Provide an injectable embedded-content resolver while retaining lightweight standalone fallback behavior.
- [x] P1 Provide a walker-backed visible-text resolver for LaTeX pipeline integrations.
- [x] P1 Parse Word run properties inside `m:ctrlPr` and math runs.
- [x] P1 Support bold, italic, color, size, font, vertical alignment, and language when meaningful to the target.
- [x] P1 Preserve ordinary `w:t`, `w:tab`, `w:br`, `w:cr`, `w:noBreakHyphen`, and `w:softHyphen` content.
- [x] P1 Define handling for `w:sym` and symbol fonts.
- [x] P2 Define standalone handling for hyperlinks: preserve visible math content and optionally expose the target.
- [x] P2 Define handling for simple and complex fields: cached result by default; evaluation remains outside this utility.
- [x] P2 Define handling for content controls, smart tags, and custom XML: unwrap visible content by default.
- [x] P2 Define tracked-change policy: accept, reject, preserve/annotate, or caller-selected.
- [x] P2 Preserve visible content inside move ranges and revision containers according to that policy.
- [x] P2 Ignore non-content range markers safely: bookmarks, comments, permissions, proofing, and custom XML ranges.
- [x] P2 Define fallback for drawings, objects, pictures, ruby, and other unexpected run content.

Math runs retain Word bold/italic, color, half-point size, font, vertical alignment,
language, and direction. MathML represents applicable run presentation directly;
LaTeX emits portable style, color, size, and vertical-alignment constructs;
UnicodeMath and basic text retain visible content and diagnose presentation that
their linear formats cannot carry. `m:ctrlPr` is parsed independently because it
formats a structure's non-selectable control character. MathML retains those
properties as `data-omml-control-*` metadata, while every target diagnoses the
control-character-only styling approximation rather than incorrectly applying it
to the complete numerator, denominator, base, or limit.

`DxpOmmlConversionOptions.RevisionMode` selects accepted, rejected, or annotated
revision and move content. `FieldMode` defaults to cached simple/complex field
results and can omit fields; evaluation remains a document-pipeline concern.
Hyperlinks always preserve visible content, and `IncludeHyperlinkTargets` plus
`HyperlinkTargetResolver` optionally appends a package-resolved target. Content
controls, smart tags, and custom XML are transparent. Non-content ranges are
ignored, while drawings, objects, pictures, ruby, content parts, and unknown Word
content use the standard fallback policy and emit `OMML011` diagnostics.

## MathML writer

- [x] P0 Emit namespace-correct XML rooted at `<math xmlns="http://www.w3.org/1998/Math/MathML">`.
- [x] P0 Set inline/block display semantics correctly.
- [x] P0 Ensure the HTML-ready surface is safe to embed without double-escaping text or admitting source markup.
- [x] P0 Emit `mi`, `mn`, `mo`, and `mtext` using deterministic tokenization.
- [ ] P0 Emit `mrow` only where grouping is semantically required.
- [x] P0 Support fractions, roots, scripts, multiscripts, fenced expressions, limits, and tables.
- [x] P1 Prefer MathML Core-compatible constructs for browser rendering.
- [x] P1 Use `stretchy`, `accent`, `accentunder`, `movablelimits`, and `mathvariant` correctly.
- [x] P1 Support `menclose`, `mphantom`, and `mpadded` for advanced layout.
- [x] P1 Preserve meaningful spacing without copying Word layout measurements blindly.
- [x] P1 Produce XML that parses without DTDs or external entities.
- [ ] P1 Add optional semantics/annotation output only behind an explicit option.
- [x] P2 Validate representative output in current Chromium, Firefox, and WebKit/Safari engines.

## LaTeX writer

- [x] P0 Emit an expression without `$`, `$$`, `\(`, or `\[` wrappers.
- [x] P0 Escape text-mode and math-mode reserved characters correctly.
- [x] P0 Use braces conservatively so nested output is unambiguous.
- [x] P0 Support standard fractions, roots, scripts, operators, limits, and matrices.
- [x] P1 Emit common named functions and symbols idiomatically.
- [x] P1 Choose deterministic environments for matrices and equation arrays.
- [x] P1 Support accents, braces, bars, boxes, and phantoms where LaTeX has a standard equivalent.
- [x] P1 Report required non-core packages, or restrict default output to a documented package baseline.
- [x] P1 Never inject raw control sequences from untrusted OMML text.
- [ ] P2 Offer a compatibility profile for MathJax/KaTeX-supported LaTeX.

## Readable Unicode text writer

- [x] P0 Preserve every visible literal, identifier, number, and operator.
- [x] P0 Use explicit grouping to avoid ambiguous flattening.
- [x] P0 Render fractions, roots, scripts, limits, and matrices in stable readable notation.
- [x] P1 Use Unicode super/subscript characters only when the complete value is representable; otherwise use `^(...)` and `_(...)`.
- [ ] P1 Render multi-line structures with a caller-selected single-line or multi-line policy.
- [x] P1 Define accessible names for invisible or purely visual constructs.
- [x] P1 Avoid dependence on terminal width or current culture.
- [x] P2 Define the supported UnicodeMath version/profile and document intentional deviations.

## Error handling, security, and quality

- [x] P0 Reject malformed XML with a clear input exception.
- [x] P0 Never silently discard an unknown semantic node.
- [x] P0 Preserve descendant visible text in the default unsupported-node fallback.
- [x] P0 Include the unsupported element name/path in diagnostics.
- [x] P1 Disable DTD processing and external entity resolution.
- [x] P1 Bound input size, nesting depth, and output growth, or document caller-enforced limits.
- [x] P1 Avoid recursive stack exhaustion on deeply nested input.
- [x] P1 Handle missing required children with explicit recovery rules.
- [x] P1 Handle duplicate singleton properties deterministically.
- [x] P1 Handle unknown attributes and future-version extension elements without losing visible content.
- [x] P1 Preserve namespace correctness regardless of the source prefix.
- [x] P1 Test alternate prefixes and default namespaces.
- [x] P1 Test null, empty, whitespace-only, and non-OMML XML inputs.
- [x] P1 Test deterministic output across cultures and line-ending conventions.
- [x] P1 Test concurrent conversions.
- [x] P2 Benchmark large equations and documents containing thousands of equations.
- [x] P2 Fuzz malformed and adversarial XML inputs.

UnicodeMath output targets the linear notation accepted by current Microsoft
Office equation input: explicit `_(...)`/`^(...)`, `▒〖...〗` n-ary operands,
`■(...)` matrices, and `█(...)` equation arrays. It intentionally favors stable,
unambiguous grouping over typographic Unicode superscript substitution. OMML-only
layout is retained where that notation has a defined operator and otherwise
reported through `OMML002`.

`MaxInputCharacters`, `MaxNestingDepth`, `MaxElementCount`, and
`MaxOutputCharacters` bound standalone conversion. XML and Open XML SDK trees
are checked iteratively before semantic recursion. Deterministic mutation tests
exercise malformed/adversarial input, and a 2,000-equation case guards repeated
sibling performance.

## Plurimath fixture audit

At OMML commit `51d4abe5df58fe33a92df094971c5828c3459ffb`, the repository contains 279 `.omml` fixtures:

- 189 general fixtures in `spec/fixtures/omml`.
- 90 fixtures in `spec/fixtures/omml/line_break`.

The corpus is valuable for complex composition and real-world regression coverage. It is not a complete property matrix.

The complete conformance gate converts all 277 well-formed fixtures through
MathML, LaTeX, UnicodeMath, and readable text, verifies nonempty output,
namespace-valid MathML, and absence of unsupported semantic nodes. Fixture
`line_break/line-break-064.omml` contains a non-schema direct `m:t` argument;
Docxport recovers its visible content. `187.omml` contains the undeclared HTML
entity `&nbsp;`, while `issue-158.omml` contains an unescaped ampersand. Those two
inputs are deliberately and explicitly rejected as malformed XML.

### Structures represented in the corpus

- [x] Import/recreate and attribute the general fixture corpus.
- [x] Import/recreate and attribute the line-break fixture corpus.
- [x] Fractions: fixtures 001-004 and complex compositions.
- [x] Scripts: fixtures 005-008 plus later nested cases.
- [x] Radicals: fixtures 009-012.
- [x] N-ary operators: fixtures 013-046 plus complex compositions.
- [x] Delimiters and piecewise expressions: fixtures 045-072.
- [x] Functions and powers: fixtures 073-099.
- [x] Limits, accents, bars, and border boxes: fixtures 100-155.
- [x] Matrices and equation arrays: fixtures 156-177.
- [x] Mixed complex expressions: fixtures 178 onward; malformed `issue-158.omml` is explicitly rejected.
- [x] Group characters: fixtures 185-186.
- [x] Manual breaks across many parent structures: `line_break/*`.
- [x] Math styles and scripts: plain/bold/italic/bold-italic plus script, fraktur, double-struck, sans-serif, and monospace examples.
- [x] Delimiter variants: parentheses, brackets, braces, bars, double bars, angles, floors, ceilings, and white brackets.
- [x] Matrix column counts and centered columns.
- [x] Hidden n-ary subscript/superscript flags and both limit locations.

### Important corpus gaps requiring Docxport fixtures

- [x] `m:box` and every `m:boxPr` behavior.
- [x] Fraction types `skw`, `lin`, and `noBar` (the reviewed fixtures do not exercise `m:type`).
- [x] Delimiter shape `m:shp` and explicit grow-off behavior.
- [x] Equation-array spacing, distance, and base-justification properties.
- [x] Matrix left/right column justification, base justification, placeholders, and spacing/gap properties.
- [x] Border-box horizontal, vertical, and diagonal strike combinations.
- [x] Phantom zero-width/ascent/descent and transparency combinations.
- [x] `m:argSz`, `m:lit`, `m:nor`, and `m:aln` semantics.
- [x] Math paragraph justification and multiple `oMath` children.
- [x] Document-level `m:mathPr` defaults.
- [x] Embedded hyperlinks, fields, content controls, custom XML, and tracked revisions.
- [x] Alternate namespace prefixes and default math namespace.
- [x] Malformed XML, missing required arguments, duplicate properties, and unknown elements.
- [x] Supplementary-plane Unicode and combining sequences.
- [x] All output-injection and resource-limit security cases.

Named Docxport-owned corpus-gap evidence is kept separately from the upstream
oracle. Valid reusable inputs live in `Fixtures/Omml/Normative` with reviewed
readable-text expectations; invalid XML inputs use `.invalid.xml`. Properties
that are supplied through options, and combinatorial cases that are clearer as
focused builders, remain named unit cases:

| Gap | Named evidence |
| --- | --- |
| Fraction variants | `fraction-types.omml`; `SupportsEveryFractionType` |
| Delimiter growth/shape and variants | `delimiter-shape-and-growth.omml`; `MapsCommonAndArbitraryUnicodeDelimiters` |
| Matrix/equation-array layout | `matrix-and-equation-array-layout.omml`; `RetainsMatrixLayoutPropertiesAndPlaceholderVisibility`; `RetainsEquationArrayPropertiesAndRows` |
| Border boxes and phantoms | `border-box-and-phantom-layout.omml`; independent border/strike/dimension theories |
| Argument size and run semantics | `argument-size.omml`; `run-semantics.omml`; `PreservesEveryApplicableRelativeArgumentSize` |
| Paragraph and document settings | `paragraph-multiple-equations.omml`; `ExposesEveryDocumentMathSettingThroughOptions` |
| Embedded WordprocessingML | `embedded-wordprocessing.omml`; `DxpOmmlEmbeddedWordprocessingTests` |
| Namespace independence | `alternate-namespace.omml`; `AcceptsAlternateAndDefaultOmmlNamespacePrefixes` |
| Malformed and recovery | `undeclared-entity.invalid.xml`; `unescaped-ampersand.invalid.xml`; `direct-math-text-recovery.omml` |
| Unicode scalars/sequences | `supplementary-and-combining-unicode.omml`; `ClassifiesMixedTokensAndPreservesSupplementaryScalars` |
| Output injection | `latex-injection-is-text.omml`; `LatexControlSequencesFromOmmlTextAreAlwaysEscapedAsText` |
| Resource limits | `RejectsInputBeyondConfiguredLimit`; `RejectsXmlBeyondConfiguredDepthBeforeSemanticRecursion`; `RejectsOpenXmlSdkTreeBeyondConfiguredElementCount`; `RejectsOutputBeyondConfiguredLimitAndTryConvertReportsIt` |

### Oracle workflow

- [x] Record upstream repository, commit, path, and license with reused fixtures.
- [x] Generate MathML, LaTeX, and UnicodeMath oracle outputs using a pinned Plurimath version.
- [x] Store generated oracle outputs separately from hand-authored normative expectations.
- [x] Canonicalize XML before comparing MathML; do not compare prefixes or insignificant whitespace.
- [x] Treat an oracle disagreement as a review prompt, not automatic proof that Docxport is wrong.
- [x] Add focused named fixtures for every discovered regression; numbered corpus files alone are difficult to diagnose.

The conformance test compares every available generated oracle artifact on each
run: 276 MathML, 277 LaTeX, and 259 UnicodeMath outputs. Exact equality is not a
conformance requirement. The pinned oracle emits structurally incomplete MathML
for missing arguments (for example, a one-child `mfrac`), drops properties that
Docxport preserves, makes different but equivalent grouping/style choices, and
contains one non-well-formed MathML artifact. LaTeX and UnicodeMath exact matches
are retained as a regression floor (currently 11 and 61 respectively), while
focused normative tests decide disagreements for fractions, radicals, scripts,
delimiters, functions, operators, decorations, matrices, arrays, boxes,
phantoms, breaks, and embedded WordprocessingML.

## Findings from the Plurimath implementation review

The reviewed importer provides useful implementation lessons that should become explicit requirements:

- [x] Preserve original element order rather than iterating model properties.
- [x] Make unsupported typed nodes fail or diagnose explicitly instead of falling through.
- [x] Distinguish empty content from absent content.
- [x] Normalize non-breaking spaces and entities without accepting malformed XML silently.
- [x] Keep token resolution contextual: text, operator, function name, accent, and delimiter are not interchangeable.
- [x] Do not bind a unary/ternary function to adjacent content without clear syntactic evidence.
- [x] Handle default characters explicitly: parentheses, integral, hat, and group characters.
- [x] Treat styled runs as content plus style, not as reordered sibling values.
- [x] Avoid quadratic behavior when consuming repeated sibling elements.

The following Plurimath behaviors are intentionally not sufficient as our target:

- [x] Preserve fraction type; the reviewed importer maps every fraction to a conventional fraction.
- [x] Preserve bar position; the reviewed importer maps every bar to the same bar function.
- [x] Preserve border-side and strike properties; the reviewed importer maps every border box to one `menclose` form.
- [x] Preserve box semantics; the reviewed importer unwraps boxes.
- [x] Preserve phantom layout semantics; the reviewed importer unwraps phantoms.
- [x] Preserve matrix/equation-array layout properties; the reviewed importer primarily retains rows and cells.
- [x] Respect `degHide`, rather than deciding square root solely from an empty degree.
- [x] Apply delimiter grow/shape behavior, not only boundary and separator characters.
- [x] Carry diagnostics for approximations rather than silently simplifying them.

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

- [x] Implement math runs, text, whitespace, Unicode scalar handling, and token classification.
- [x] Implement literal/normal text, math style, math script, escaping, and applicable Word run properties.
- [x] Integrate symbol-font translation where sufficient information exists.
- [x] Cover simple-run, mixed-token, styled-run, invisible-character, and supplementary-Unicode cases.

### Goal 4: Fractions, radicals, and scripts

- [x] Implement every fraction type, including bar, skewed, linear, and no-bar.
- [x] Implement square and indexed roots, degree hiding, and empty-degree distinctions.
- [x] Implement subscript, superscript, combined scripts, pre-scripts, and script alignment.
- [x] Cover structured and deeply nested bases, degrees, numerators, denominators, and scripts.

### Goal 5: Delimiters and decorations

- [x] Implement delimiters, repeated arguments, separators, empty boundaries, growth, and shape.
- [x] Implement accents, bars above/below, and group characters.
- [x] Cover common and arbitrary Unicode delimiters and accents.
- [x] Cover decorated matrices and nested decorated expressions.

### Goal 6: Functions, limits, and n-ary operators

- [x] Implement named and arbitrary functions without unsafe semantic guessing.
- [x] Implement upper and lower limits.
- [x] Implement standard and arbitrary n-ary operators, hidden limits, growth, and limit placement.
- [x] Cover differences between ordinary scripts and operator limits in inline and display math.

### Goal 7: Matrices and equation arrays

- [x] Implement matrix rows/cells, column groups/counts, alignment, placeholders, spacing, and gaps.
- [x] Implement equation-array rows, alignment points, justification, distances, and spacing.
- [x] Handle ragged and inconsistent input without losing cell content.
- [x] Produce idiomatic MathML tables, LaTeX environments, UnicodeMath, and readable text.

### Goal 8: Boxes, border boxes, and phantoms

- [x] Implement box operator, break, alignment, differential, and no-break semantics.
- [x] Implement every border-side and horizontal/vertical/diagonal strike combination.
- [x] Implement phantom visibility, transparency, zero-width, zero-ascent, and zero-descent behavior.
- [x] Document and diagnose every output-specific approximation.

### Goal 9: Breaks and equation layout

- [x] Implement manual breaks and alignment indices across every supported parent structure.
- [x] Implement math-paragraph justification, multiple equations, and relative alignment.
- [x] Implement optional document math settings and their local-property precedence.
- [x] Cover all dedicated upstream line-break fixtures.

### Goal 10: Embedded WordprocessingML

- [x] Add the embedded-content resolver boundary and walker-backed LaTeX text adapter.
- [x] Implement visible Word run content and formatting inside math.
- [x] Define and implement standalone policies for fields, hyperlinks, content controls, smart tags, and custom XML.
- [x] Define and implement tracked-change handling independently from the document pipeline.
- [x] Safely unwrap or diagnose bookmarks, comments, range markers, drawings, objects, and unexpected content.

### Goal 11: Corpus conformance and gap closure

- [x] Run the complete pinned general and line-break corpora.
- [x] Resolve every crash, silent-content loss, and unexplained oracle difference.
- [x] Add named Docxport-owned fixtures for every corpus gap listed above.
- [x] Verify every normative OMML structure and property has a test or documented intentional limitation.

Goal 11's corpus guarantees are enforced, not observational: all valid inputs run
through every writer; MathML is parsed again; unsupported semantic diagnostics
fail the gate; and every source text literal must survive readable output unless
OMML explicitly hides it. Oracle artifacts are compared on every run and exact
matches cannot regress below the recorded LaTeX/UnicodeMath floors.

The remaining unchecked items elsewhere in this checklist are explicit later
scope, not corpus ambiguities: customization/streaming APIs, minimal `mrow`
optimization, optional MathML annotations, selectable compatibility and
multiline profiles, browser rendering, and production depth/performance/fuzz
hardening. Conservative `mrow` grouping is semantically valid but not claimed
minimal. These limitations do not discard a normative OMML structure or
property; all lossy conversions identified by Goal 11 retain visible content and
emit a diagnostic.

### Goal 12: Production hardening

- [x] Complete malformed, adversarial, unknown-extension, namespace, and resource-limit tests.
- [x] Complete concurrency, culture, line-ending, performance, and fuzz testing.
- [x] Validate all target frameworks, including browser/WASM.
- [x] Validate representative MathML in supported browser engines and LaTeX under the documented profile.
- [x] Complete public API and compatibility documentation.

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

- [x] Every normative OMML structure has a tested semantic conversion or a documented intentional limitation.
- [x] Every property above has a test proving preservation, approximation, or an explicit diagnostic.
- [x] All imported corpus fixtures convert without crashes or silent visible-content loss.
- [x] Every corpus gap above has a focused Docxport-owned fixture.
- [x] MathML is namespace-valid and renders acceptably in supported browsers.
- [x] LaTeX compiles under the documented baseline and/or renders under the selected MathJax/KaTeX profile.
- [x] Text output remains understandable when copied without styling.
- [x] Malformed, adversarial, and future-version input fails safely.
- [x] Public API behavior, defaults, diagnostics, and compatibility guarantees are documented.
- [x] Pipeline integration can call the standalone API without special conversion branches.
