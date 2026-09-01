using System.Text;
using System.Xml.Linq;

namespace DocxportNet.Omml;

internal static class OmmlWriter
{
    private const string MathMlNamespace = "http://www.w3.org/1998/Math/MathML";

    public static DxpOmmlConversionResult Write(
        OmmlDocument document,
        DxpOmmlOutputFormat format,
        DxpOmmlConversionOptions options)
    {
        bool isDisplay = options.Display ?? document.IsDisplay;
        List<DxpOmmlDiagnostic> diagnostics = new();
        string output = format switch
        {
            DxpOmmlOutputFormat.MathMl => WriteMathMl(document, isDisplay, options, diagnostics),
            DxpOmmlOutputFormat.Latex => WriteTextual(document, format, options, diagnostics, EscapeLatex),
            DxpOmmlOutputFormat.UnicodeMath => WriteTextual(document, format, options, diagnostics, static value => value),
            DxpOmmlOutputFormat.Text => WriteTextual(document, format, options, diagnostics, static value => value),
            _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unknown OMML output format."),
        };
        return new DxpOmmlConversionResult(output, format, isDisplay, diagnostics.AsReadOnly());
    }

    private static string WriteMathMl(
        OmmlDocument document,
        bool isDisplay,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics)
    {
        XNamespace math = MathMlNamespace;
        XElement row = new(math + "mrow");
        foreach (OmmlNode node in document.Children)
            AppendMathMl(row, node, options.SmallFractions && !isDisplay, options, diagnostics);

        XElement root = new(
            math + "math",
            new XAttribute("display", isDisplay ? "block" : "inline"),
            row);
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static void AppendMathMl(
        XElement parent,
        OmmlNode node,
        bool compactFractions,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics)
    {
        XNamespace math = MathMlNamespace;
        if (node is OmmlSequence sequence)
        {
            XElement row = new(math + "mrow");
            foreach (OmmlNode child in sequence.Children)
                AppendMathMl(row, child, compactFractions, options, diagnostics);
            parent.Add(row);
            return;
        }

        if (node is OmmlRun run)
        {
            XElement container = parent;
            string? variant = MathVariant(run);
            if (variant != null || run.Language != null || run.RightToLeft)
            {
                container = new XElement(math + "mstyle");
                if (variant != null) container.SetAttributeValue("mathvariant", variant);
                if (run.Language != null) container.SetAttributeValue(XNamespace.Xml + "lang", run.Language);
                if (run.RightToLeft) container.SetAttributeValue("dir", "rtl");
                parent.Add(container);
            }
            if (run.Alignment) container.Add(new XElement(math + "malignmark"));
            foreach (OmmlToken token in run.Tokens)
            {
                if (token.Value == "\u200B") container.Add(new XElement(math + "mspace", new XAttribute("width", "0")));
                else container.Add(new XElement(math + (token.Kind switch { OmmlTokenKind.Identifier => "mi", OmmlTokenKind.Number => "mn", OmmlTokenKind.Operator => "mo", _ => "mtext" }), token.Value));
            }
            return;
        }

        if (node is OmmlFraction fraction)
        {
            XElement rendered;
            if (fraction.Type == OmmlFractionType.Linear)
            {
                rendered = new XElement(math + "mrow");
                AppendMathMl(rendered, fraction.Numerator, compactFractions, options, diagnostics);
                rendered.Add(new XElement(math + "mo", "/"));
                AppendMathMl(rendered, fraction.Denominator, compactFractions, options, diagnostics);
            }
            else
            {
                rendered = new XElement(math + "mfrac");
                if (fraction.Type == OmmlFractionType.Skewed) rendered.SetAttributeValue("bevelled", "true");
                if (fraction.Type == OmmlFractionType.NoBar) rendered.SetAttributeValue("linethickness", "0");
                AppendMathMl(rendered, fraction.Numerator, compactFractions, options, diagnostics);
                AppendMathMl(rendered, fraction.Denominator, compactFractions, options, diagnostics);
            }
            if (compactFractions)
                parent.Add(new XElement(math + "mstyle", new XAttribute("displaystyle", "false"), new XAttribute("scriptlevel", "1"), rendered));
            else parent.Add(rendered);
            return;
        }

        if (node is OmmlRadical radical)
        {
            bool indexed = radical.HasDegree && !radical.DegreeHidden;
            XElement rendered = new(math + (indexed ? "mroot" : "msqrt"));
            AppendMathMl(rendered, radical.Radicand, compactFractions, options, diagnostics);
            if (indexed) AppendMathMl(rendered, radical.Degree, compactFractions, options, diagnostics);
            parent.Add(rendered);
            return;
        }

        if (node is OmmlScript script)
        {
            XElement rendered = new(math + (script.Type switch
            {
                OmmlScriptType.Subscript => "msub",
                OmmlScriptType.Superscript => "msup",
                OmmlScriptType.SubSup => "msubsup",
                _ => "mmultiscripts",
            }));
            AppendMathMl(rendered, script.Base, compactFractions, options, diagnostics);
            if (script.Type == OmmlScriptType.PreSubSup)
            {
                rendered.Add(new XElement(math + "mprescripts"));
                AppendMathMl(rendered, script.Subscript, compactFractions, options, diagnostics);
                AppendMathMl(rendered, script.Superscript, compactFractions, options, diagnostics);
            }
            else
            {
                if (script.Type != OmmlScriptType.Superscript) AppendMathMl(rendered, script.Subscript, compactFractions, options, diagnostics);
                if (script.Type != OmmlScriptType.Subscript) AppendMathMl(rendered, script.Superscript, compactFractions, options, diagnostics);
            }
            if (script.AlignScripts) rendered.SetAttributeValue("data-omml-align-scripts", "true");
            parent.Add(rendered);
            return;
        }

        string fallback = ResolveFallback((OmmlUnsupported)node, options, diagnostics);
        if (fallback.Length != 0)
            parent.Add(new XElement(math + "mtext", fallback));
    }

    private static string WriteTextual(
        OmmlDocument document,
        DxpOmmlOutputFormat format,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics,
        Func<string, string> escape)
    {
        StringBuilder output = new();
        foreach (OmmlNode node in document.Children)
            AppendTextual(output, node, format, options, diagnostics, escape);
        return output.ToString();
    }

    private static void AppendTextual(
        StringBuilder output,
        OmmlNode node,
        DxpOmmlOutputFormat format,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics,
        Func<string, string> escape)
    {
        if (node is OmmlSequence sequence)
        {
            foreach (OmmlNode child in sequence.Children)
                AppendTextual(output, child, format, options, diagnostics, escape);
            return;
        }

        if (node is OmmlRun run)
        {
            string value = string.Concat(run.Tokens.Select(token => token.Value));
            if (format == DxpOmmlOutputFormat.Latex)
            {
                value = EscapeLatex(value);
                value = ApplyLatexStyle(run, value);
                if (run.Alignment) output.Append('&');
                output.Append(value.Replace("\u200B", "{}"));
            }
            else
            {
                if (run.Alignment && format == DxpOmmlOutputFormat.UnicodeMath) output.Append('&');
                if (format == DxpOmmlOutputFormat.UnicodeMath && UnicodeMathStyle(run) is string command)
                    output.Append($"\\{command}\"{value.Replace("\"", "\\\"")}\"");
                else output.Append(escape(format == DxpOmmlOutputFormat.Text ? value.Replace("\u200B", string.Empty) : value));
            }
            return;
        }

        if (node is OmmlFraction fraction)
        {
            string numerator = RenderTextual(fraction.Numerator, format, options, diagnostics, escape);
            string denominator = RenderTextual(fraction.Denominator, format, options, diagnostics, escape);
            output.Append(format switch
            {
                DxpOmmlOutputFormat.Latex when fraction.Type == OmmlFractionType.Bar => $"\\frac{{{numerator}}}{{{denominator}}}",
                DxpOmmlOutputFormat.Latex when fraction.Type == OmmlFractionType.NoBar => $"\\genfrac{{}}{{}}{{0pt}}{{}}{{{numerator}}}{{{denominator}}}",
                DxpOmmlOutputFormat.Latex => $"{{{numerator}}}/{{{denominator}}}",
                DxpOmmlOutputFormat.UnicodeMath => $"({numerator})/({denominator})",
                _ => $"({numerator})/({denominator})",
            });
            return;
        }

        if (node is OmmlRadical radical)
        {
            string radicand = RenderTextual(radical.Radicand, format, options, diagnostics, escape);
            string degree = RenderTextual(radical.Degree, format, options, diagnostics, escape);
            bool indexed = radical.HasDegree && !radical.DegreeHidden;
            output.Append(format switch
            {
                DxpOmmlOutputFormat.Latex when indexed => $"\\sqrt[{degree}]{{{radicand}}}",
                DxpOmmlOutputFormat.Latex => $"\\sqrt{{{radicand}}}",
                DxpOmmlOutputFormat.UnicodeMath when indexed => $"√({degree}&{radicand})",
                DxpOmmlOutputFormat.UnicodeMath => $"√({radicand})",
                _ when indexed => $"root({degree}, {radicand})",
                _ => $"sqrt({radicand})",
            });
            return;
        }

        if (node is OmmlScript script)
        {
            string @base = RenderTextual(script.Base, format, options, diagnostics, escape);
            string sub = RenderTextual(script.Subscript, format, options, diagnostics, escape);
            string sup = RenderTextual(script.Superscript, format, options, diagnostics, escape);
            if (format == DxpOmmlOutputFormat.Latex)
            {
                string latexBase = @base.Length == 0 ? "{}" : @base;
                output.Append(script.Type switch { OmmlScriptType.Subscript => $"{latexBase}_{{{sub}}}", OmmlScriptType.Superscript => $"{latexBase}^{{{sup}}}", OmmlScriptType.SubSup => $"{latexBase}_{{{sub}}}^{{{sup}}}", _ => $"{{}}_{{{sub}}}^{{{sup}}}{@base}" });
            }
            else if (format == DxpOmmlOutputFormat.UnicodeMath)
                output.Append(script.Type switch { OmmlScriptType.Subscript => $"{@base}_({sub})", OmmlScriptType.Superscript => $"{@base}^({sup})", OmmlScriptType.SubSup => $"{@base}_({sub})^({sup})", _ => $"_({sub})^({sup}) {@base}" });
            else
                output.Append(script.Type switch { OmmlScriptType.Subscript => $"{@base}_({sub})", OmmlScriptType.Superscript => $"{@base}^({sup})", OmmlScriptType.SubSup => $"{@base}_({sub})^({sup})", _ => $"[{sub},{sup}]{@base}" });
            return;
        }

        output.Append(escape(ResolveFallback((OmmlUnsupported)node, options, diagnostics)));
    }

    private static string RenderTextual(OmmlNode node, DxpOmmlOutputFormat format,
        DxpOmmlConversionOptions options, List<DxpOmmlDiagnostic> diagnostics, Func<string, string> escape)
    {
        StringBuilder result = new();
        AppendTextual(result, node, format, options, diagnostics, escape);
        return result.ToString();
    }

    private static string ResolveFallback(
        OmmlUnsupported unsupported,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics)
    {
        DxpOmmlDiagnostic diagnostic = new(
            "OMML001",
            DxpOmmlDiagnosticSeverity.Warning,
            $"OMML element '{unsupported.ElementName}' is not yet semantically supported; fallback policy '{options.FallbackPolicy}' was applied at {unsupported.Path}.",
            unsupported.Path,
            unsupported.ElementName);
        diagnostics.Add(diagnostic);

        return options.FallbackPolicy switch
        {
            DxpOmmlFallbackPolicy.Throw => throw new DxpOmmlUnsupportedException(diagnostic),
            DxpOmmlFallbackPolicy.ExtractText => unsupported.VisibleText,
            DxpOmmlFallbackPolicy.Placeholder => options.Placeholder ?? string.Empty,
            DxpOmmlFallbackPolicy.Omit => string.Empty,
            _ => throw new ArgumentOutOfRangeException(nameof(options.FallbackPolicy)),
        };
    }

    private static string EscapeLatex(string value)
    {
        StringBuilder escaped = new(value.Length);
        foreach (char character in value)
        {
            escaped.Append(character switch
            {
                '\\' => @"\textbackslash{}",
                '{' => @"\{",
                '}' => @"\}",
                '$' => @"\$",
                '&' => @"\&",
                '#' => @"\#",
                '%' => @"\%",
                '_' => @"\_",
                '^' => @"\^{}",
                '~' => @"\~{}",
                _ => character.ToString(),
            });
        }
        return escaped.ToString();
    }

    private static string? MathVariant(OmmlRun run)
    {
        if (run.Normal || (run.Style == OmmlMathStyle.Plain &&
            run.Script is OmmlMathScript.Default or OmmlMathScript.Roman))
            return "normal";
        string weight = run.Style switch
        {
            OmmlMathStyle.Bold => "bold",
            OmmlMathStyle.Italic => "italic",
            OmmlMathStyle.BoldItalic => "bold-italic",
            _ => "",
        };
        return run.Script switch
        {
            OmmlMathScript.Roman => weight.Length == 0 ? "normal" : weight,
            OmmlMathScript.Script => weight.StartsWith("bold", StringComparison.Ordinal) ? "bold-script" : "script",
            OmmlMathScript.Fraktur => weight.StartsWith("bold", StringComparison.Ordinal) ? "bold-fraktur" : "fraktur",
            OmmlMathScript.DoubleStruck => "double-struck",
            OmmlMathScript.SansSerif => weight switch
            {
                "bold" => "bold-sans-serif",
                "italic" => "sans-serif-italic",
                "bold-italic" => "sans-serif-bold-italic",
                _ => "sans-serif",
            },
            OmmlMathScript.Monospace => "monospace",
            _ => weight.Length == 0 ? null : weight,
        };
    }

    private static string ApplyLatexStyle(OmmlRun run, string value)
    {
        string? script = run.Script switch { OmmlMathScript.Roman => "mathrm", OmmlMathScript.Script => "mathcal", OmmlMathScript.Fraktur => "mathfrak", OmmlMathScript.DoubleStruck => "mathbb", OmmlMathScript.SansSerif => "mathsf", OmmlMathScript.Monospace => "mathtt", _ => null };
        if (script != null) value = $"\\{script}{{{value}}}";
        string? weight = run.Style switch { OmmlMathStyle.Bold => "mathbf", OmmlMathStyle.Italic => "mathit", OmmlMathStyle.BoldItalic => "boldsymbol", OmmlMathStyle.Plain when script == null => "mathrm", _ => null };
        return weight == null ? value : $"\\{weight}{{{value}}}";
    }

    private static string? UnicodeMathStyle(OmmlRun run)
    {
        return run.Script switch { OmmlMathScript.Roman => run.Style switch { OmmlMathStyle.Bold => "mbf", OmmlMathStyle.Italic => "mit", OmmlMathStyle.BoldItalic => "mbfit", _ => "mup" }, OmmlMathScript.Script => run.Style is OmmlMathStyle.Bold or OmmlMathStyle.BoldItalic ? "mbfscr" : "mscr", OmmlMathScript.Fraktur => run.Style is OmmlMathStyle.Bold or OmmlMathStyle.BoldItalic ? "mbffrak" : "mfrak", OmmlMathScript.DoubleStruck => "Bbb", OmmlMathScript.SansSerif => run.Style switch { OmmlMathStyle.Bold => "mbfsans", OmmlMathStyle.Italic => "mitsans", OmmlMathStyle.BoldItalic => "mbfitsans", _ => "msans" }, OmmlMathScript.Monospace => "mtt", _ => run.Style switch { OmmlMathStyle.Bold => "mbf", OmmlMathStyle.Italic => "mit", OmmlMathStyle.BoldItalic => "mbfit", OmmlMathStyle.Plain => "mup", _ => null } };
    }
}
