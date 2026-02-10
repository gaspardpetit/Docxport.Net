using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Fields.Formatting;
using System.Linq;

namespace DocxportNet.Fields.Eval;

internal static class DxpFieldEvalRules
{
    internal static bool HasMergeFormat(DxpFieldParser parser, string instruction)
    {
        var parse = parser.Parse(instruction);
        if (!TryGetCharOrMergeFormat(parse.Ast.FormatSpecs, out _, out var hasMergeFormat))
            return false;
        return hasMergeFormat;
    }

    internal static string GetEvaluationErrorText(DxpFieldParser parser, string instruction)
    {
        var parse = parser.Parse(instruction);
        var fieldType = parse.Ast.FieldType;
        if (string.IsNullOrWhiteSpace(fieldType))
            return "Error! Invalid field code.";

        var normalizedFieldType = fieldType!.Trim().ToUpperInvariant();
        switch (normalizedFieldType)
        {
            case "REF":
                return "Error! Reference source not found.";
            case "DOCVARIABLE":
                return "Error! No document variable supplied.";
            case "DOCPROPERTY":
                return "Error! Unknown document property name.";
            case "IF":
                return "Error! Invalid field code.";
            case "=":
                return "Error! Invalid formula.";
            default:
                return "Error! Invalid field code.";
        }
    }

    internal static bool TryGetCharOrMergeFormat(
        IReadOnlyList<IDxpFieldFormatSpec> specs,
        out bool hasCharFormat,
        out bool hasMergeFormat)
    {
        hasCharFormat = false;
        hasMergeFormat = false;
        foreach (var spec in specs)
        {
            if (spec is not DxpTextTransformFormatSpec transform)
                continue;
            if (transform.Kind == DxpTextTransformKind.Charformat)
                hasCharFormat = true;
            else if (transform.Kind == DxpTextTransformKind.MergeFormat)
                hasMergeFormat = true;
        }
        return hasCharFormat || hasMergeFormat;
    }

    internal static bool HasRenderableContent(Run r)
    {
        return r.ChildElements.Any(child =>
            child is Text or DeletedText or NoBreakHyphen or TabChar or Break or CarriageReturn or Drawing);
    }
}
