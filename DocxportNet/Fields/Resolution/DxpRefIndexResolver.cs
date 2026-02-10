using DocxportNet.API;

namespace DocxportNet.Fields.Resolution;

public sealed class DxpRefIndexResolver : IDxpRefResolver
{
    public Task<DxpRefRecord?> ResolveAsync(
        DxpRefRequest request,
        DxpFieldEvalContext context,
        DxpIDocumentContext? documentContext)
    {
        var index = documentContext?.DocumentIndex.RefIndex;
        var documentNodes = documentContext?.DocumentIndex.BookmarkNodes;
        DxpFieldNodeBuffer? nodes = null;
        if (context.TryGetBookmarkNodes(request.Bookmark, out var evalNodes))
            nodes = evalNodes;
        else if (documentNodes != null && documentNodes.TryGetValue(request.Bookmark, out var docNodes))
            nodes = docNodes;

        if (index != null && index.Bookmarks.TryGetValue(request.Bookmark, out var bm))
        {
            index.ParagraphNumbers.TryGetValue(request.Bookmark, out var para);
            index.Footnotes.TryGetValue(request.Bookmark, out var footnote);
            index.Endnotes.TryGetValue(request.Bookmark, out var endnote);
            index.Hyperlinks.TryGetValue(request.Bookmark, out var hyperlink);
            return Task.FromResult<DxpRefRecord?>(DxpRefRecords.FromIndex(
                request.Bookmark,
                nodes,
                bm,
                para,
                footnote,
                endnote,
                hyperlink));
        }

        if (nodes == null)
            return Task.FromResult<DxpRefRecord?>(null);

        return Task.FromResult<DxpRefRecord?>(DxpRefRecords.FromNodes(request.Bookmark, nodes));
    }
}
