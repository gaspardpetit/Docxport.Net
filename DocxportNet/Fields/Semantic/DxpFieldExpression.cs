using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Fields.Semantic;

internal sealed record DxpFieldExpression(
    IReadOnlyList<DxpFieldExpressionPart> Parts,
    DxpSemanticSourceProvenance? Source = null)
{
    public static DxpFieldExpression FromText(string text)
        => new(new DxpFieldExpressionPart[] { new DxpFieldExpressionText(text) });

    public string FieldType
        => DxpFieldExpressionTokenizer.Tokenize(this).FirstOrDefault()?.LiteralText ?? string.Empty;
}

internal static class DxpFieldExpressionSource
{
    private const string Word2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";

    public static DxpSemanticSourceProvenance Capture(
        Run? sourceRun,
        DxpFieldEvalContext context)
    {
        Paragraph? paragraph = sourceRun?.Ancestors<Paragraph>().FirstOrDefault();
        string? paragraphId = paragraph?.ExtendedAttributes
            .FirstOrDefault(attribute =>
                attribute.LocalName == "paraId" && attribute.NamespaceUri == Word2010Namespace)
            .Value;
        return new DxpSemanticSourceProvenance(
            context.CurrentStoryKeyProvider?.Invoke(),
            context.CurrentDocumentOrder,
            string.IsNullOrEmpty(paragraphId) ? null : paragraphId);
    }
}

internal abstract record DxpFieldExpressionPart;
internal sealed record DxpFieldExpressionText(string Text) : DxpFieldExpressionPart;
internal sealed record DxpFieldExpressionChild(DxpFieldExpression Expression) : DxpFieldExpressionPart;
internal sealed record DxpFieldExpressionParagraph : DxpFieldExpressionPart;

internal sealed record DxpFieldExpressionToken(IReadOnlyList<DxpFieldTemplatePart> Parts)
{
    public string? LiteralText
        => Parts.All(static part => part is DxpFieldTemplateText)
            ? string.Concat(Parts.Cast<DxpFieldTemplateText>().Select(static part => part.Text))
            : null;
}

internal abstract record DxpFieldTemplatePart;
internal sealed record DxpFieldTemplateText(string Text) : DxpFieldTemplatePart;
internal sealed record DxpFieldTemplateChild(DxpFieldExpression Expression) : DxpFieldTemplatePart;
internal sealed record DxpFieldTemplateParagraph : DxpFieldTemplatePart;
