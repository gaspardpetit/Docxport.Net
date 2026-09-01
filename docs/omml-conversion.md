# Standalone OMML conversion

`DocxportNet.Omml.DxpOmmlConverter` converts an `m:oMath` or `m:oMathPara`
fragment independently of the DOCX export pipeline. It accepts XML strings and
Open XML SDK `OfficeMath` or math `Paragraph` objects. SDK objects are traversed
directly and are not serialized and reparsed.

The supported outputs are MathML, expression-only LaTeX, UnicodeMath, and basic
readable text. `ToHtml` returns the same native MathML as `ToMathMl`; document
visitors remain responsible for surrounding HTML or Markdown delimiters.

```csharp
using DocxportNet.Omml;

string mathml = DxpOmmlConverter.ToMathMl(omml);
string latex = DxpOmmlConverter.ToLatex(omml);
string unicodeMath = DxpOmmlConverter.ToUnicodeMath(omml);
string text = DxpOmmlConverter.ToText(omml);
```

Use `Convert` to receive diagnostics and inferred display mode. `TryConvert`
returns malformed-input and unsupported-input failures without throwing.

Valid structures that do not yet have semantic support produce `OMML001`
diagnostics. The default policy preserves their visible descendant text. Callers
can instead request a placeholder, omission, or an exception through
`DxpOmmlConversionOptions.FallbackPolicy`. This fallback behavior is the Goal 2
foundation; later feature goals replace it with structural conversions.

Parsing prohibits DTDs and external entities, validates the OMML namespace and
root, and applies a configurable input-character limit. Output is deterministic,
culture-invariant, and contains no document-walker or visitor dependency.

## Runs and tokens

Math runs are parsed semantically rather than through the unsupported-element
fallback. Text, significant whitespace, empty text, tabs, line breaks, BMP
characters, and supplementary Unicode scalars retain their input order. Mixed
runs are split into identifier, number, operator, and text tokens for MathML;
for example, `x+12` becomes `mi`, `mo`, and `mn` tokens.

`m:lit`, `m:nor`, every `m:sty` value, and the roman, script, fraktur,
double-struck, sans-serif, and monospace `m:scr` families are supported. MathML
uses `mathvariant`; LaTeX uses nested alphabet commands; UnicodeMath uses its
linear alphabet controls. Basic text intentionally drops styling while retaining
readable content. Applicable Word bold, italic, language, RTL, run-font, and
`w:sym` information is honored. Known legacy symbol fonts use
`DxpFontSymbols`.

An OMML alignment marker becomes `malignmark` in MathML and `&` in LaTeX and
UnicodeMath. U+200B becomes a zero-width MathML space, `{}` in LaTeX, remains
available in UnicodeMath for fidelity, and is omitted from basic readable text.

## Fractions, radicals, and scripts

Fractions support the default/bar, skewed, linear, and no-bar forms. MathML uses
`mfrac` (including `bevelled` and zero-line-thickness variants) or an explicit
linear slash; LaTeX uses `\frac`, `\genfrac`, or a linear form. UnicodeMath and
basic text use unambiguous parenthesized numerator/denominator notation. Set
`DxpOmmlConversionOptions.SmallFractions` when a caller has obtained Word's
document-level `smallFrac` setting; this applies compact MathML styling only to
inline expressions.

Radicals distinguish a missing degree, an explicitly empty degree, a visible
degree, and a degree suppressed by `degHide`. Scripts support subscript,
superscript, combined subscript/superscript, and prescripts. Empty arguments and
arbitrarily nested supported structures are preserved. `alnScr` intent is
retained as `data-omml-align-scripts` on MathML; textual formats retain the
script structure but have no separate alignment-style mechanism. Control-run
property presence is retained in the internal semantic model for later pipeline
integration. Ordinary scripts are not guessed to be operator limits; explicit
limit structures are handled separately by the later functions/limits goal.
