using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Walker.Context;

internal static class DxpSectionBreakListRule
{
    internal static bool ShouldAdvance(Paragraph paragraph, bool acceptRevisions)
    {
        if (paragraph.ParagraphProperties?.SectionProperties == null)
            return true;

        return paragraph.Descendants().Any(element =>
            IsRenderable(element) && IsInRevisionView(element, paragraph, acceptRevisions));
    }

    private static bool IsRenderable(OpenXmlElement element) => element switch
    {
        Text text => !string.IsNullOrEmpty(text.Text),
        DeletedText text => !string.IsNullOrEmpty(text.Text),
        Drawing or Picture or SymbolChar or FootnoteReference or EndnoteReference or
            DocumentFormat.OpenXml.Math.OfficeMath or Break or TabChar or CarriageReturn or
            NoBreakHyphen or SoftHyphen or PositionalTab or EmbeddedObject or ContentPart or
            PageNumber or DayShort or DayLong or MonthShort or MonthLong or YearShort or YearLong => true,
        _ => false
    };

    private static bool IsInRevisionView(OpenXmlElement element, Paragraph paragraph, bool acceptRevisions)
    {
        for (var ancestor = element; ancestor != null && !ReferenceEquals(ancestor, paragraph); ancestor = ancestor.Parent)
        {
            if (acceptRevisions && ancestor is DeletedRun)
                return false;
            if (!acceptRevisions && ancestor is InsertedRun)
                return false;
        }
        return true;
    }
}
