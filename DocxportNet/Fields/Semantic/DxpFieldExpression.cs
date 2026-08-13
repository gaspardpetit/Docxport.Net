namespace DocxportNet.Fields.Semantic;

internal sealed record DxpFieldExpression(IReadOnlyList<DxpFieldExpressionPart> Parts)
{
    public static DxpFieldExpression FromText(string text)
        => new(new DxpFieldExpressionPart[] { new DxpFieldExpressionText(text) });

    public string FieldType
        => DxpFieldExpressionTokenizer.Tokenize(this).FirstOrDefault()?.LiteralText ?? string.Empty;
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
