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

public sealed record DxpRefResult(
    string? Text,
    string? HyperlinkTarget = null,
    string? FootnoteText = null,
    string? FootnoteMark = null);

public sealed record DxpRefHyperlink(string Bookmark, string Target, string Text);

public sealed record DxpRefFootnote(string Bookmark, string Text, string? Mark);

public interface IDxpRefResolver
{
    Task<DxpRefResult?> ResolveAsync(DxpRefRequest request, DxpFieldEvalContext context);
}
