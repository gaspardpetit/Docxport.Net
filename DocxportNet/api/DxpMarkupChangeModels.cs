using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.API;

public enum DxpMarkupChangeKind
{
    Unchanged,
    Inserted,
    Deleted
}

public sealed record DxpMarkupChangeContext(
    Run Run,
    Paragraph Paragraph,
    DxpStyleEffectiveRunStyle ResolvedStyle,
    RunProperties? RunProperties,
    string? RunStyleId,
    IReadOnlyList<DxpStyleInfo> ParagraphStyleChain,
    DxpIDocumentContext DocumentContext
);

public sealed record DxpMarkupChangeDecision(
    DxpMarkupChangeKind ChangeKind,
    DxpStyleEffectiveRunStyle RenderStyle
);

public interface DxpIMarkupChangeClassifierProvider
{
    Func<DxpMarkupChangeContext, DxpMarkupChangeDecision?>? MarkupChangeClassifier { get; }
}

public static class DxpMarkupChangeClassifiers
{
    private static readonly Func<DxpMarkupChangeContext, DxpMarkupChangeDecision?> UnderlineInsertedStrikeDeletedClassifier = ClassifyUnderlineInsertedStrikeDeleted;

    public static Func<DxpMarkupChangeContext, DxpMarkupChangeDecision?> UnderlineInsertedStrikeDeleted() =>
        UnderlineInsertedStrikeDeletedClassifier;

    private static DxpMarkupChangeDecision? ClassifyUnderlineInsertedStrikeDeleted(DxpMarkupChangeContext context)
    {
        var style = context.ResolvedStyle;
        if (style.Strike || style.DoubleStrike)
        {
            return new DxpMarkupChangeDecision(
                DxpMarkupChangeKind.Deleted,
                style with {
                    Strike = false,
                    DoubleStrike = false
                });
        }

        if (style.Underline)
        {
            return new DxpMarkupChangeDecision(
                DxpMarkupChangeKind.Inserted,
                style with { Underline = false });
        }

        return null;
    }
}
