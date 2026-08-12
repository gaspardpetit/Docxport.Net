using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Runtime.CompilerServices;

namespace DocxportNet.Fields.Resolution;

internal static class DxpBookmarkRangeProjector
{
    public static bool TryProject(
        Body body,
        string bookmarkName,
        out IReadOnlyList<OpenXmlElement> blocks,
        out string? error)
    {
        blocks = Array.Empty<OpenXmlElement>();
        error = null;

        var positions = new Dictionary<OpenXmlElement, ElementRange>(OpenXmlElementReferenceComparer.Instance);
        int nextPosition = 0;
        Index(body, positions, ref nextPosition);

        foreach (var start in body.Descendants<BookmarkStart>()
                     .Where(candidate => string.Equals(candidate.Name?.Value, bookmarkName, StringComparison.OrdinalIgnoreCase)))
        {
            string? id = start.Id?.Value;
            if (string.IsNullOrWhiteSpace(id))
                continue;

            var startRange = positions[start];
            var end = body.Descendants<BookmarkEnd>()
                .Where(candidate => string.Equals(candidate.Id?.Value, id, StringComparison.Ordinal))
                .FirstOrDefault(candidate => positions[candidate].Start > startRange.Start);
            if (end == null)
                continue;

            int endPosition = positions[end].Start;
            var projected = new List<OpenXmlElement>();
            foreach (var block in body.ChildElements)
            {
                var clone = Project(block, startRange.Start, endPosition, positions, out bool hasContent);
                if (clone != null && hasContent)
                    projected.Add(clone);
            }

            blocks = projected;
            return true;
        }

        error = $"Bookmark '{bookmarkName}' was not found or does not have a valid matching end marker.";
        return false;
    }

    private static ElementRange Index(
        OpenXmlElement element,
        IDictionary<OpenXmlElement, ElementRange> positions,
        ref int nextPosition)
    {
        int start = nextPosition++;
        int end = start;
        foreach (var child in element.ChildElements)
            end = Index(child, positions, ref nextPosition).End;
        var range = new ElementRange(start, end);
        positions[element] = range;
        return range;
    }

    private static OpenXmlElement? Project(
        OpenXmlElement element,
        int startPosition,
        int endPosition,
        IReadOnlyDictionary<OpenXmlElement, ElementRange> positions,
        out bool hasContent)
    {
        hasContent = false;
        var range = positions[element];
        if (range.End <= startPosition || range.Start >= endPosition)
            return null;

        if (range.Start > startPosition && range.End < endPosition)
        {
            hasContent = !IsStructuralProperty(element);
            return element.CloneNode(true);
        }

        if (element is BookmarkStart or BookmarkEnd)
            return null;

        if (!element.HasChildren)
            return null;

        var clone = element.CloneNode(false);
        foreach (var child in element.ChildElements)
        {
            if (IsStructuralProperty(child))
            {
                clone.AppendChild(child.CloneNode(true));
                continue;
            }

            var childClone = Project(child, startPosition, endPosition, positions, out bool childHasContent);
            if (childClone != null)
                clone.AppendChild(childClone);
            hasContent |= childHasContent;
        }

        return hasContent ? clone : null;
    }

    private static bool IsStructuralProperty(OpenXmlElement element)
        => element is ParagraphProperties
            or RunProperties
            or TableProperties
            or TableGrid
            or TableRowProperties
            or TableCellProperties
            or SectionProperties
            || element.LocalName.EndsWith("Pr", StringComparison.Ordinal);

    private readonly record struct ElementRange(int Start, int End);

    private sealed class OpenXmlElementReferenceComparer : IEqualityComparer<OpenXmlElement>
    {
        public static OpenXmlElementReferenceComparer Instance { get; } = new();
        public bool Equals(OpenXmlElement? x, OpenXmlElement? y) => ReferenceEquals(x, y);
        public int GetHashCode(OpenXmlElement obj) => RuntimeHelpers.GetHashCode(obj);
    }
}
