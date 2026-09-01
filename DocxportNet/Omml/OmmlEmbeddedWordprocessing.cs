using DocumentFormat.OpenXml;
using System.Text;
using System.Xml.Linq;

namespace DocxportNet.Omml;

internal static class OmmlEmbeddedWordprocessing
{
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Word2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    public static bool TryResolve(OmmlUnsupported unsupported, DxpOmmlConversionOptions options,
        List<DxpOmmlDiagnostic> diagnostics, out string text)
    {
        IReadOnlyList<Node> roots = unsupported.XmlElements.Count != 0
            ? unsupported.XmlElements.Select(FromXElement).ToArray()
            : unsupported.OpenXmlElements.Select(FromOpenXml).ToArray();
        if (roots.Count == 0 || roots.Any(root => root.Namespace != WordNamespace && root.Namespace != Word2010Namespace))
        {
            text = string.Empty;
            return false;
        }

        StringBuilder output = new();
        RenderState state = new(options.FieldMode != DxpOmmlFieldMode.Omit);
        foreach (Node root in roots)
            Render(root, state, output, unsupported, options, diagnostics);
        text = output.ToString();
        return true;
    }

    private static void Render(Node node, RenderState state, StringBuilder output,
        OmmlUnsupported source, DxpOmmlConversionOptions options, List<DxpOmmlDiagnostic> diagnostics)
    {
        if (node.Namespace != WordNamespace && node.Namespace != Word2010Namespace)
        {
            if (node.Namespace == MathNamespace && node.Name == "t" && state.Emit)
                output.Append(node.Text);
            else
                RenderChildren(node, state, output, source, options, diagnostics);
            return;
        }

        switch (node.Name)
        {
            case "t":
            case "delText":
                if (state.Emit) output.Append(node.Text);
                return;
            case "tab":
                if (state.Emit) output.Append('\t');
                return;
            case "br":
            case "cr":
                if (state.Emit) output.Append('\n');
                return;
            case "noBreakHyphen":
                if (state.Emit) output.Append('-');
                return;
            case "softHyphen":
                if (state.Emit) output.Append('\u00AD');
                return;
            case "sym":
                if (state.Emit)
                    output.Append(global::DocxportNet.DxpFontSymbols.TranslateWordSymbol(
                        node.Attribute("font"), node.Attribute("char")));
                return;
            case "instrText":
            case "rPr":
            case "sdtPr":
            case "sdtEndPr":
            case "smartTagPr":
            case "customXmlPr":
                return;
            case "fldChar":
                UpdateFieldState(node.Attribute("fldCharType") ?? node.Attribute("type"), state, options.FieldMode);
                return;
            case "fldSimple":
                if (options.FieldMode == DxpOmmlFieldMode.CachedResult)
                    RenderChildren(node, state, output, source, options, diagnostics);
                return;
            case "ins":
            case "moveTo":
            case "conflictIns":
                RenderRevision(node, inserted: true, state, output, source, options, diagnostics);
                return;
            case "del":
            case "moveFrom":
            case "conflictDel":
                RenderRevision(node, inserted: false, state, output, source, options, diagnostics);
                return;
            case "hyperlink":
                RenderHyperlink(node, state, output, source, options, diagnostics);
                return;
            case "bookmarkStart":
            case "bookmarkEnd":
            case "commentRangeStart":
            case "commentRangeEnd":
            case "permStart":
            case "permEnd":
            case "proofErr":
            case "customXmlInsRangeStart":
            case "customXmlInsRangeEnd":
            case "customXmlDelRangeStart":
            case "customXmlDelRangeEnd":
            case "customXmlMoveFromRangeStart":
            case "customXmlMoveFromRangeEnd":
            case "customXmlMoveToRangeStart":
            case "customXmlMoveToRangeEnd":
            case "moveFromRangeStart":
            case "moveFromRangeEnd":
            case "moveToRangeStart":
            case "moveToRangeEnd":
                return;
            case "drawing":
            case "object":
            case "pict":
            case "ruby":
            case "contentPart":
                AppendUnexpected(node, output, source, options, diagnostics);
                return;
            case "r":
            case "sdt":
            case "sdtContent":
            case "smartTag":
            case "customXml":
            case "dir":
            case "bdo":
                RenderChildren(node, state, output, source, options, diagnostics);
                return;
            default:
                AppendUnexpected(node, output, source, options, diagnostics);
                return;
        }
    }

    private static void RenderChildren(Node node, RenderState state, StringBuilder output,
        OmmlUnsupported source, DxpOmmlConversionOptions options, List<DxpOmmlDiagnostic> diagnostics)
    {
        foreach (Node child in node.Children)
            Render(child, state, output, source, options, diagnostics);
    }

    private static void RenderRevision(Node node, bool inserted, RenderState state, StringBuilder output,
        OmmlUnsupported source, DxpOmmlConversionOptions options, List<DxpOmmlDiagnostic> diagnostics)
    {
        bool include = options.RevisionMode switch
        {
            DxpOmmlRevisionMode.Accept => inserted,
            DxpOmmlRevisionMode.Reject => !inserted,
            _ => true,
        };
        if (!include) return;

        if (options.RevisionMode == DxpOmmlRevisionMode.Preserve && state.Emit)
            output.Append(inserted ? "[inserted:" : "[deleted:");
        RenderChildren(node, state, output, source, options, diagnostics);
        if (options.RevisionMode == DxpOmmlRevisionMode.Preserve && state.Emit)
            output.Append(']');
    }

    private static void RenderHyperlink(Node node, RenderState state, StringBuilder output,
        OmmlUnsupported source, DxpOmmlConversionOptions options, List<DxpOmmlDiagnostic> diagnostics)
    {
        RenderChildren(node, state, output, source, options, diagnostics);
        if (!state.Emit || !options.IncludeHyperlinkTargets) return;

        string? relationshipId = node.Attribute("id");
        string? anchor = node.Attribute("anchor");
        string? target = options.HyperlinkTargetResolver?.Invoke(relationshipId, anchor) ?? anchor;
        if (!string.IsNullOrEmpty(target)) output.Append(" (").Append(target).Append(')');
        else
            diagnostics.Add(new DxpOmmlDiagnostic("OMML012", DxpOmmlDiagnosticSeverity.Warning,
                "The embedded hyperlink target was requested but could not be resolved without package context.",
                source.Path, source.ElementName));
    }

    private static void UpdateFieldState(string? type, RenderState state, DxpOmmlFieldMode mode)
    {
        switch (type)
        {
            case "begin":
                state.FieldEmitStack.Push(state.Emit);
                state.Emit = false;
                break;
            case "separate":
                if (state.FieldEmitStack.Count != 0)
                    state.Emit = mode == DxpOmmlFieldMode.CachedResult && state.FieldEmitStack.Peek();
                break;
            case "end":
                if (state.FieldEmitStack.Count != 0)
                    state.Emit = state.FieldEmitStack.Pop();
                break;
        }
    }

    private static void AppendUnexpected(Node node, StringBuilder output, OmmlUnsupported source,
        DxpOmmlConversionOptions options, List<DxpOmmlDiagnostic> diagnostics)
    {
        DxpOmmlDiagnostic diagnostic = new("OMML011", DxpOmmlDiagnosticSeverity.Warning,
            $"Embedded WordprocessingML element 'w:{node.Name}' cannot be rendered faithfully; fallback policy '{options.FallbackPolicy}' was applied.",
            source.Path, $"w:{node.Name}");
        diagnostics.Add(diagnostic);
        switch (options.FallbackPolicy)
        {
            case DxpOmmlFallbackPolicy.Throw:
                throw new DxpOmmlUnsupportedException(diagnostic);
            case DxpOmmlFallbackPolicy.ExtractText:
                output.Append(VisibleText(node));
                break;
            case DxpOmmlFallbackPolicy.Placeholder:
                output.Append(options.Placeholder);
                break;
        }
    }

    private static string VisibleText(Node node)
    {
        StringBuilder result = new();
        AppendVisibleText(node, result);
        return result.ToString();
    }

    private static void AppendVisibleText(Node node, StringBuilder output)
    {
        if ((node.Namespace == WordNamespace || node.Namespace == MathNamespace) &&
            (node.Name is "t" or "delText")) output.Append(node.Text);
        else if (node.Namespace == WordNamespace && node.Name == "tab") output.Append('\t');
        else if (node.Namespace == WordNamespace && node.Name is "br" or "cr") output.Append('\n');
        else foreach (Node child in node.Children) AppendVisibleText(child, output);
    }

    private static Node FromXElement(XElement element) => new(
        element.Name.NamespaceName,
        element.Name.LocalName,
        element.HasElements ? string.Empty : element.Value,
        element.Attributes().ToDictionary(a => a.Name.LocalName, a => a.Value, StringComparer.Ordinal),
        element.Elements().Select(FromXElement).ToArray());

    private static Node FromOpenXml(OpenXmlElement element) => new(
        element.NamespaceUri,
        element.LocalName,
        element.HasChildren ? string.Empty : element.InnerText,
        element.GetAttributes().GroupBy(a => a.LocalName, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.First().Value ?? string.Empty, StringComparer.Ordinal),
        element.ChildElements.Select(FromOpenXml).ToArray());

    private sealed record Node(string Namespace, string Name, string Text,
        IReadOnlyDictionary<string, string> Attributes, IReadOnlyList<Node> Children)
    {
        public string? Attribute(string name) => Attributes.TryGetValue(name, out string? value) ? value : null;
    }

    private sealed class RenderState(bool emit)
    {
        public bool Emit { get; set; } = emit;
        public Stack<bool> FieldEmitStack { get; } = new();
    }
}
