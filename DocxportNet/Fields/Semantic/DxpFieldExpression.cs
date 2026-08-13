using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Fields.Semantic;

internal sealed record DxpFieldExpression(
    IReadOnlyList<DxpFieldExpressionPart> Parts,
    DxpSemanticSourceProvenance? Source = null,
    string? CachedResult = null,
    DxpSemanticParagraphFormat? SourceParagraphFormat = null)
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

    public static DxpSemanticRunFormat? CaptureRunFormat(Run? run)
    {
        RunProperties? properties = run?.RunProperties;
        if (properties == null)
            return null;
        return new DxpSemanticRunFormat(
            ReadOnOff(properties.Bold),
            ReadOnOff(properties.Italic),
            ReadOnOff(properties.Strike),
            properties.Underline?.Val?.InnerText,
            properties.Color?.Val?.Value,
            properties.FontSize?.Val?.Value,
            properties.RunStyle?.Val?.Value,
            properties.Languages?.Val?.Value);
    }

    public static DxpSemanticParagraphFormat? CaptureParagraphFormat(Paragraph? paragraph)
    {
        ParagraphProperties? properties = paragraph?.ParagraphProperties;
        if (properties == null)
            return null;
        return new DxpSemanticParagraphFormat(
            properties.ParagraphStyleId?.Val?.Value,
            properties.Justification?.Val?.InnerText,
            properties.OutlineLevel?.Val?.Value,
            properties.NumberingProperties?.NumberingId?.Val?.Value,
            properties.NumberingProperties?.NumberingLevelReference?.Val?.Value);
    }

    private static bool? ReadOnOff(DocumentFormat.OpenXml.Wordprocessing.OnOffType? value)
        => value == null ? null : value.Val?.Value ?? true;
}

internal abstract record DxpFieldExpressionPart;
internal sealed record DxpFieldExpressionText(
    string Text,
    DxpSemanticRunFormat? Format = null) : DxpFieldExpressionPart;
internal sealed record DxpFieldExpressionChild(DxpFieldExpression Expression) : DxpFieldExpressionPart;
internal sealed record DxpFieldExpressionParagraph(
    DxpSemanticParagraphFormat? Format = null) : DxpFieldExpressionPart;

internal sealed record DxpFieldExpressionToken(IReadOnlyList<DxpFieldTemplatePart> Parts)
{
    public string? LiteralText
        => Parts.All(static part => part is DxpFieldTemplateText)
            ? string.Concat(Parts.Cast<DxpFieldTemplateText>().Select(static part => part.Text))
            : null;
}

internal abstract record DxpFieldTemplatePart;
internal sealed record DxpFieldTemplateText(
    string Text,
    DxpSemanticRunFormat? Format = null) : DxpFieldTemplatePart;
internal sealed record DxpFieldTemplateChild(DxpFieldExpression Expression) : DxpFieldTemplatePart;
internal sealed record DxpFieldTemplateParagraph(
    DxpSemanticParagraphFormat? Format = null) : DxpFieldTemplatePart;
