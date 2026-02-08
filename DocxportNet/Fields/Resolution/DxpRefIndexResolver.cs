using DocxportNet.API;

namespace DocxportNet.Fields.Resolution;

public sealed class DxpRefIndexResolver : IDxpRefResolver
{
    private readonly DxpRefIndex _index;
    private readonly Func<int?> _currentOrderProvider;

    public DxpRefIndexResolver(DxpRefIndex index, Func<int?> currentOrderProvider)
    {
        _index = index ?? throw new ArgumentNullException(nameof(index));
        _currentOrderProvider = currentOrderProvider ?? throw new ArgumentNullException(nameof(currentOrderProvider));
    }

    public Task<DxpRefResult?> ResolveAsync(DxpRefRequest request, DxpFieldEvalContext context)
    {
        if (!_index.Bookmarks.TryGetValue(request.Bookmark, out var bm))
            return Task.FromResult<DxpRefResult?>(null);

        string? text = bm.Text;

        bool wantsParagraphNumber = request.FullContextParagraphNumber ||
            request.RelativeParagraphNumber ||
            request.ParagraphNumber;

        if (wantsParagraphNumber)
        {
            text = null;
            if (_index.ParagraphNumbers.TryGetValue(request.Bookmark, out var para))
            {
                if (request.FullContextParagraphNumber)
                    text = para.FullNumber;
                else if (request.RelativeParagraphNumber)
                    text = para.CurrentLevelNumber;
                else if (request.ParagraphNumber)
                    text = para.CurrentLevelNumber;

                if (request.SuppressNonNumeric)
                    text = para.NumericOnly;
            }
        }

        if (!string.IsNullOrEmpty(request.SeparatorText) && !string.IsNullOrEmpty(text))
            text = text!.Replace(".", request.SeparatorText);

        if (request.AboveBelow && wantsParagraphNumber)
        {
            var current = _currentOrderProvider();
            if (current.HasValue)
            {
                text = AppendAboveBelow(text, current.Value, bm.DocumentOrder);
            }
        }

        string? hyperlinkTarget = request.Hyperlink ? request.Bookmark : null;
        string? footnoteText = null;
        string? footnoteMark = null;
        if (request.Footnote && _index.Footnotes.TryGetValue(request.Bookmark, out var fn))
        {
            footnoteText = fn.Text;
            footnoteMark = fn.Mark;
        }
        if (request.Footnote && _index.Endnotes.TryGetValue(request.Bookmark, out var en))
        {
            footnoteText = en.Text;
            footnoteMark = en.Mark;
        }
        if (request.Hyperlink && _index.Hyperlinks.TryGetValue(request.Bookmark, out var link))
        {
            hyperlinkTarget = link.Target;
        }

        if (request.Footnote && !string.IsNullOrEmpty(footnoteText))
            text = footnoteText;

        return Task.FromResult<DxpRefResult?>(new DxpRefResult(text, hyperlinkTarget, footnoteText, footnoteMark));
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
