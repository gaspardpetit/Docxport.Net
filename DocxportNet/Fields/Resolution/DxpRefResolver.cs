using DocxportNet.API;

namespace DocxportNet.Fields.Resolution;

public sealed record DxpRefRequest(
    string Bookmark,
    bool Separator,
    bool Footnote,
    bool Hyperlink,
    bool ParagraphNumber,
    bool AboveBelow,
    bool RelativeParagraphNumber,
    bool SuppressNonNumeric,
    bool FullContextParagraphNumber,
    string? SeparatorText);

public static class DxpRefRequests
{
    public static DxpRefRequest Simple(string bookmark) => new(
        Bookmark: bookmark,
        Separator: false,
        Footnote: false,
        Hyperlink: false,
        ParagraphNumber: false,
        AboveBelow: false,
        RelativeParagraphNumber: false,
        SuppressNonNumeric: false,
        FullContextParagraphNumber: false,
        SeparatorText: null);
}

public static class DxpRefRecords
{
    public static DxpRefRecord FromIndex(
        string bookmark,
        DxpFieldNodeBuffer? nodes,
        DxpRefBookmark bm,
        DxpRefParagraphNumber? para,
        DxpRefFootnote? footnote,
        DxpRefEndnote? endnote,
        DxpRefHyperlink? hyperlink)
        => new(
            Bookmark: bookmark,
            Nodes: nodes,
            DocumentText: bm.Text,
            DocumentOrder: bm.DocumentOrder,
            ParagraphNumber: para,
            Footnote: footnote,
            Endnote: endnote,
            Hyperlink: hyperlink);

    public static DxpRefRecord FromNodes(string bookmark, DxpFieldNodeBuffer nodes)
        => new(
            Bookmark: bookmark,
            Nodes: nodes,
            DocumentText: null,
            DocumentOrder: null,
            ParagraphNumber: null,
            Footnote: null,
            Endnote: null,
            Hyperlink: null);
}

public sealed record DxpRefResult(
    string? Text,
    string? HyperlinkTarget = null,
    string? FootnoteText = null,
    string? FootnoteMark = null);

public sealed record DxpRefRecord(
    string Bookmark,
    DxpFieldNodeBuffer? Nodes,
    string? DocumentText,
    int? DocumentOrder,
    DxpRefParagraphNumber? ParagraphNumber,
    DxpRefFootnote? Footnote,
    DxpRefEndnote? Endnote,
    DxpRefHyperlink? Hyperlink);

public static class DxpRefRecordExtensions
{
    public static DxpRefResult? Format(
        this DxpRefRecord record,
        DxpRefRequest request,
        DxpFieldEvalContext context)
    {
        if (record == null)
            return null;

        bool wantsParagraphNumber = request.FullContextParagraphNumber ||
            request.RelativeParagraphNumber ||
            request.ParagraphNumber;

        string? text = null;
        if (wantsParagraphNumber && record.ParagraphNumber != null)
        {
            if (request.FullContextParagraphNumber)
                text = record.ParagraphNumber.FullNumber;
            else if (request.RelativeParagraphNumber)
                text = record.ParagraphNumber.CurrentLevelNumber;
            else if (request.ParagraphNumber)
                text = record.ParagraphNumber.CurrentLevelNumber;

            if (request.SuppressNonNumeric)
                text = record.ParagraphNumber.NumericOnly;
        }
        else
        {
            text = record.Nodes != null ? record.Nodes.ToPlainText() : record.DocumentText;
        }

        if (!string.IsNullOrEmpty(request.SeparatorText) && !string.IsNullOrEmpty(text))
            text = text!.Replace(".", request.SeparatorText);

        if (request.AboveBelow && wantsParagraphNumber && record.DocumentOrder.HasValue)
        {
            var current = context.CurrentDocumentOrder;
            if (current.HasValue)
                text = AppendAboveBelow(text, current.Value, record.DocumentOrder.Value);
        }

        string? hyperlinkTarget = null;
        if (request.Hyperlink)
            hyperlinkTarget = record.Hyperlink?.Target ?? record.Bookmark;

        string? footnoteText = null;
        string? footnoteMark = null;
        if (request.Footnote)
        {
            if (record.Footnote != null)
            {
                footnoteText = record.Footnote.Text;
                footnoteMark = record.Footnote.Mark;
            }
            else if (record.Endnote != null)
            {
                footnoteText = record.Endnote.Text;
                footnoteMark = record.Endnote.Mark;
            }

            if (!string.IsNullOrEmpty(footnoteText))
                text = footnoteText;
        }

        return new DxpRefResult(text, hyperlinkTarget, footnoteText, footnoteMark);
    }

    private static string? AppendAboveBelow(string? text, int current, int target)
    {
        if (string.IsNullOrEmpty(text))
            text = string.Empty;
        var label = current < target ? "below" : current > target ? "above" : null;
        if (label == null)
            return text;
        return string.IsNullOrEmpty(text) ? label : $"{text} {label}";
    }
}

public interface IDxpRefResolver
{
    Task<DxpRefRecord?> ResolveAsync(
        DxpRefRequest request,
        DxpFieldEvalContext context,
        DxpIDocumentContext? documentContext);
}
