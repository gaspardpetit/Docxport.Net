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
            DxpOmmlOutputFormat.Latex => WriteTextual(document, options, diagnostics, EscapeLatex),
            DxpOmmlOutputFormat.UnicodeMath => WriteTextual(document, options, diagnostics, static value => value),
            DxpOmmlOutputFormat.Text => WriteTextual(document, options, diagnostics, static value => value),
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

        string fallback = ResolveFallback((OmmlUnsupported)node, options, diagnostics);
        if (fallback.Length != 0)
            parent.Add(new XElement(math + "mtext", fallback));
    }

    private static string WriteTextual(
        OmmlDocument document,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics,
        Func<string, string> escape)
    {
        StringBuilder output = new();
        foreach (OmmlNode node in document.Children)
            AppendTextual(output, node, options, diagnostics, escape);
        return output.ToString();
    }

    private static void AppendTextual(
        StringBuilder output,
        OmmlNode node,
        DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics,
        Func<string, string> escape)
    {
        if (node is OmmlSequence sequence)
        {
            foreach (OmmlNode child in sequence.Children)
                AppendTextual(output, child, options, diagnostics, escape);
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
}
