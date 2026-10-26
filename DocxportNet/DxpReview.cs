using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet;

/// <summary>Accepts or rejects selected tracked changes in a DOCX package.</summary>
public static class DxpReview
{
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// An empty set selects no changes. A null accept set accepts every change not
    /// selected for rejection; a null reject set rejects every change not selected
    /// for acceptance. Two null sets leave the document unchanged.
    /// </summary>
    public static byte[] Apply(byte[] docxBytes, IReadOnlyCollection<string>? accept = null,
        IReadOnlyCollection<string>? reject = null)
    {
        if (docxBytes == null) throw new ArgumentNullException(nameof(docxBytes));
        var acceptSet = accept == null ? null : new HashSet<string>(accept, StringComparer.Ordinal);
        var rejectSet = reject == null ? null : new HashSet<string>(reject, StringComparer.Ordinal);
        if ((accept == null && reject == null) || (acceptSet?.Count == 0 && rejectSet?.Count == 0))
            return (byte[])docxBytes.Clone();
        if (acceptSet != null && rejectSet != null && acceptSet.Overlaps(rejectSet))
            throw new ArgumentException("A revision cannot be both accepted and rejected.");

        using var stream = new MemoryStream();
        stream.Write(docxBytes, 0, docxBytes.Length);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            if (document.MainDocumentPart == null)
                throw new InvalidDataException("The DOCX has no main document part.");

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var part in EnumerateParts(document))
            {
                var root = part.RootElement;
                if (root == null) continue;
                foreach (var change in root.Descendants().Reverse().ToArray())
                {
                    if (change.Parent == null || change.NamespaceUri != WordNamespace || !IsRevision(change))
                        continue;
                    string? rawId = change.GetAttribute("id", WordNamespace).Value;
                    if (string.IsNullOrEmpty(rawId)) continue;
                    string id = DxpReviewId.Create(part, rawId);
                    bool explicitlyAccept = acceptSet?.Contains(id) == true;
                    bool explicitlyReject = rejectSet?.Contains(id) == true;
                    if (explicitlyAccept || explicitlyReject) seen.Add(id);
                    bool? decision = explicitlyAccept ? true : explicitlyReject ? false
                        : accept == null ? true : reject == null ? false : null;
                    if (decision != null)
                        ApplyChange(change, decision.Value);
                }
                root.Save();
            }
            var missing = (accept ?? []).Concat(reject ?? []).Where(id => !seen.Contains(id)).Distinct().ToArray();
            if (missing.Length != 0)
                throw new ArgumentException("Revision ID not found: " + string.Join(", ", missing));
        }
        return stream.ToArray();
    }

    private static IEnumerable<OpenXmlPart> EnumerateParts(OpenXmlPartContainer package)
    {
        var pending = new Queue<OpenXmlPartContainer>();
        var visited = new HashSet<OpenXmlPart>();
        pending.Enqueue(package);
        while (pending.Count != 0)
        {
            foreach (var relationship in pending.Dequeue().Parts)
            {
                var part = relationship.OpenXmlPart;
                if (!visited.Add(part)) continue;
                yield return part;
                pending.Enqueue(part);
            }
        }
    }

    private static bool IsRevision(OpenXmlElement element)
    {
        string name = element.LocalName;
        return name is "ins" or "del" or "moveFrom" or "moveTo" or
            "cellIns" or "cellDel" or "cellMerge" or
            "moveFromRangeStart" or "moveFromRangeEnd" or "moveToRangeStart" or "moveToRangeEnd"
            || name.EndsWith("PrChange", StringComparison.Ordinal);
    }

    private static void ApplyChange(OpenXmlElement change, bool accept)
    {
        string name = change.LocalName;
        if (name.EndsWith("PrChange", StringComparison.Ordinal))
        {
            if (!accept && change.Parent is OpenXmlCompositeElement properties)
            {
                var previous = change.ChildElements.FirstOrDefault();
                if (previous != null)
                {
                    foreach (var child in properties.ChildElements.Where(child => child != change).ToArray())
                        child.Remove();
                    foreach (var child in previous.ChildElements.ToArray())
                        properties.InsertBefore(child.CloneNode(true), change);
                }
            }
            change.Remove();
            return;
        }

        bool keepContent = name is "ins" or "moveTo" ? accept : !accept;
        if (change.Parent is TableCellProperties cellProperties &&
            cellProperties.Parent is TableCell cell)
        {
            if (name is "cellIns" or "cellDel")
            {
                if ((name == "cellIns") != accept) cell.Remove();
                else change.Remove();
                return;
            }
            if (change is CellMerge merge)
            {
                var state = accept ? merge.VerticalMerge?.Value : merge.VerticalMergeOriginal?.Value;
                cellProperties.RemoveAllChildren<VerticalMerge>();
                if (state == VerticalMergeRevisionValues.Continue)
                    cellProperties.InsertBefore(new VerticalMerge { Val = MergedCellValues.Continue }, change);
                else if (state == VerticalMergeRevisionValues.Restart)
                    cellProperties.InsertBefore(new VerticalMerge { Val = MergedCellValues.Restart }, change);
                change.Remove();
                return;
            }
        }
        if (change.Parent is TableRowProperties && change.Parent.Parent is TableRow row)
        {
            if (!keepContent) row.Remove();
            else change.Remove();
            return;
        }
        if (change.Parent is NumberingProperties numbering)
        {
            if (!keepContent) numbering.Remove();
            else change.Remove();
            return;
        }
        var paragraphProperties = change.Parent as ParagraphProperties
            ?? (change.Parent as ParagraphMarkRunProperties)?.Parent as ParagraphProperties;
        if (paragraphProperties?.Parent is Paragraph paragraph)
        {
            if (!keepContent && paragraph.NextSibling<Paragraph>() is Paragraph next)
            {
                // The remaining paragraph break belongs to the second paragraph.
                paragraph.ParagraphProperties?.Remove();
                if (next.ParagraphProperties is ParagraphProperties nextProperties)
                    paragraph.PrependChild((ParagraphProperties)nextProperties.CloneNode(true));
                foreach (var child in next.ChildElements.Where(child => child is not ParagraphProperties).ToArray())
                    paragraph.Append(child.CloneNode(true));
                next.Remove();
            }
            change.Remove();
            return;
        }
        if (name.EndsWith("RangeStart", StringComparison.Ordinal) || name.EndsWith("RangeEnd", StringComparison.Ordinal))
        {
            change.Remove();
            return;
        }
        if (!keepContent)
        {
            change.Remove();
            return;
        }
        if (change.Parent is not OpenXmlCompositeElement parent)
            return;
        foreach (var child in change.ChildElements.ToArray())
        {
            var copy = child.CloneNode(true);
            foreach (var deletedText in copy.Descendants<DeletedText>().ToArray())
                deletedText.InsertAfterSelf(new Text(deletedText.Text) { Space = deletedText.Space });
            foreach (var deletedText in copy.Descendants<DeletedText>().ToArray())
                deletedText.Remove();
            foreach (var deletedField in copy.Descendants<DeletedFieldCode>().ToArray())
            {
                deletedField.InsertAfterSelf(new FieldCode(deletedField.Text) { Space = deletedField.Space });
                deletedField.Remove();
            }
            parent.InsertBefore(copy, change);
        }
        change.Remove();
    }
}
