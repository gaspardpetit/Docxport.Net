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
