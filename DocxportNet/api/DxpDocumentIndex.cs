using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace DocxportNet.API;

public sealed class DxpDocumentIndex
{
    public IReadOnlyDictionary<string, string> Bookmarks { get; }
    public IReadOnlyList<string> SequenceIdentifiers { get; }
    public DxpRefIndex RefIndex { get; }
    public IReadOnlyList<DxpCaptionEntry> Captions { get; }

    public DxpDocumentIndex(
        IReadOnlyDictionary<string, string> bookmarks,
        IReadOnlyList<string> sequenceIdentifiers,
        DxpRefIndex? refIndex = null,
        IReadOnlyList<DxpCaptionEntry>? captions = null)
    {
        Bookmarks = bookmarks;
        SequenceIdentifiers = sequenceIdentifiers;
        RefIndex = refIndex ?? DxpRefIndex.Empty;
        Captions = captions ?? [];
    }

    public static DxpDocumentIndex Build(WordprocessingDocument doc)
    {
        var bookmarks = new Dictionary<string, StringBuilder>(StringComparer.OrdinalIgnoreCase);
        var openBookmarks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var seqIdentifiers = new List<string>();

        var body = doc.MainDocumentPart?.Document?.Body;
        if (body != null)
        {
            foreach (var el in body.Descendants())
            {
                switch (el)
                {
                    case BookmarkStart bs:
                    {
                        var name = bs.Name?.Value;
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            openBookmarks.Add(name!);
                            if (!bookmarks.ContainsKey(name!))
                                bookmarks[name!] = new StringBuilder();
                        }
                        break;
                    }
                    case BookmarkEnd be:
                    {
                        var name = be.Id?.Value;
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            openBookmarks.Remove(name!);
                        }
                        break;
                    }
                    case Text text:
                        if (openBookmarks.Count > 0)
                        {
                            foreach (var name in openBookmarks)
                                bookmarks[name].Append(text.Text);
                        }
                        break;
                    case FieldCode fieldCode:
                        TryCollectSeqIdentifier(fieldCode.Text, seqIdentifiers);
                        break;
                    case SimpleField simple:
                        if (simple.Instruction != null)
                            TryCollectSeqIdentifier(simple.Instruction.Value, seqIdentifiers);
                        break;
                }
            }
        }

        var resolvedBookmarks = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var kvp in bookmarks)
            resolvedBookmarks[kvp.Key] = kvp.Value.ToString();

        return new DxpDocumentIndex(resolvedBookmarks, seqIdentifiers, DxpRefIndex.Empty);
    }

    private static void TryCollectSeqIdentifier(string? instruction, List<string> seqIdentifiers)
    {
        if (string.IsNullOrWhiteSpace(instruction))
            return;

        var text = instruction!.Trim();
        if (!text.StartsWith("SEQ", StringComparison.OrdinalIgnoreCase))
            return;

        var parts = text.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2)
            return;

        var identifier = parts[1];
        if (!string.IsNullOrWhiteSpace(identifier))
            seqIdentifiers.Add(identifier);
    }
}

public sealed record DxpCaptionEntry(
    string Text,
    string? SequenceIdentifier,
    int DocumentOrder
);
