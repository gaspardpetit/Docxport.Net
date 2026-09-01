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
            AppendMathMl(row, node, options, diagnostics);

        XElement root = new(
            math + "math",
            new XAttribute("display", isDisplay ? "block" : "inline"),
            row);
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private static void AppendMathMl(
        XElement parent,
        OmmlNode node,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics)
    {
        XNamespace math = MathMlNamespace;
        if (node is OmmlSequence sequence)
        {
            XElement row = new(math + "mrow");
            foreach (OmmlNode child in sequence.Children)
                AppendMathMl(row, child, options, diagnostics);
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

        output.Append(escape(ResolveFallback((OmmlUnsupported)node, options, diagnostics)));
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
