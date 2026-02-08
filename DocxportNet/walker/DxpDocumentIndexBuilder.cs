using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Word;
using System.Text;

namespace DocxportNet.Walker;

internal static class DxpDocumentIndexBuilder
{
    public static DxpDocumentIndex Build(WordprocessingDocument doc, DxpStyleResolver styles)
    {
        var bookmarks = new Dictionary<string, StringBuilder>(StringComparer.OrdinalIgnoreCase);
        var bookmarkOrder = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var paragraphNumbers = new Dictionary<string, DxpRefParagraphNumber>(StringComparer.OrdinalIgnoreCase);
        var refBookmarks = new Dictionary<string, DxpRefBookmark>(StringComparer.OrdinalIgnoreCase);
        var refFootnotes = new Dictionary<string, DxpRefFootnote>(StringComparer.OrdinalIgnoreCase);
        var refEndnotes = new Dictionary<string, DxpRefEndnote>(StringComparer.OrdinalIgnoreCase);
        var refHyperlinks = new Dictionary<string, DxpRefHyperlink>(StringComparer.OrdinalIgnoreCase);
        var captions = new List<DxpCaptionEntry>();
        var footnotesById = BuildFootnoteIndex(doc);
        var endnotesById = BuildEndnoteIndex(doc);
        var seqIdentifiers = new List<string>();
        var openById = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var openNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int order = 0;
        int paragraphOrder = 0;

        var lists = new DxpLists();
        lists.Init(doc);

        var body = doc.MainDocumentPart?.Document?.Body;
        if (body != null)
        {
            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                paragraphOrder++;
                var marker = lists.MaterializeMarker(paragraph, styles).marker;
                var fullNumber = NormalizeMarker(marker);
                var currentLevel = fullNumber != null ? ExtractCurrentLevel(fullNumber) : null;
                var numericOnly = fullNumber != null ? StripNonNumeric(fullNumber) : null;

                if (IsCaptionParagraph(paragraph, styles))
                {
                    var text = ExtractParagraphText(paragraph);
                    var seqId = TryExtractSeqIdentifier(paragraph);
                    if (!string.IsNullOrWhiteSpace(text) || !string.IsNullOrWhiteSpace(seqId))
                        captions.Add(new DxpCaptionEntry(text, seqId, paragraphOrder));
                }

                foreach (var el in paragraph.Descendants())
                {
                    switch (el)
                    {
                        case BookmarkStart bs:
                        {
                            var name = bs.Name?.Value;
                            if (!string.IsNullOrWhiteSpace(name))
                            {
                                openNames.Add(name!);
                                if (!bookmarks.ContainsKey(name!))
                                    bookmarks[name!] = new StringBuilder();
                                if (!bookmarkOrder.ContainsKey(name!))
                                    bookmarkOrder[name!] = order++;
                                if (fullNumber != null && !paragraphNumbers.ContainsKey(name!))
                                {
                                    paragraphNumbers[name!] = new DxpRefParagraphNumber(
                                        name!,
                                        fullNumber,
                                        currentLevel ?? fullNumber,
                                        numericOnly ?? string.Empty);
                                }
                            }
                            var id = bs.Id?.Value;
                            if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(name))
                                openById[id!] = name!;
                            break;
                        }
                        case BookmarkEnd be:
                        {
                            var id = be.Id?.Value;
                            if (!string.IsNullOrWhiteSpace(id) && openById.TryGetValue(id!, out var name))
                            {
                                openNames.Remove(name);
                                openById.Remove(id!);
                            }
                            break;
                        }
                        case Text text:
                            if (openNames.Count > 0)
                            {
                                foreach (var name in openNames)
                                    bookmarks[name].Append(text.Text);
                            }
                            break;
                        case FootnoteReference fnRef:
                            if (openNames.Count > 0 && fnRef.Id?.Value is long fnId && footnotesById.TryGetValue(fnId, out var footnote))
                            {
                                foreach (var name in openNames)
                                {
                                    if (!refFootnotes.ContainsKey(name))
                                        refFootnotes[name] = new DxpRefFootnote(name, footnote.Mark, footnote.Text);
                                }
                            }
                            break;
                        case EndnoteReference enRef:
                            if (openNames.Count > 0 && enRef.Id?.Value is long enId && endnotesById.TryGetValue(enId, out var endnote))
                            {
                                foreach (var name in openNames)
                                {
                                    if (!refEndnotes.ContainsKey(name))
                                        refEndnotes[name] = new DxpRefEndnote(name, endnote.Mark, endnote.Text);
                                }
                            }
                            break;
                        case Hyperlink link:
                            if (openNames.Count > 0)
                            {
                                var target = ResolveHyperlinkTarget(link);
                                if (!string.IsNullOrWhiteSpace(target))
                                {
                                    foreach (var name in openNames)
                                    {
                                        if (!refHyperlinks.ContainsKey(name))
                                            refHyperlinks[name] = new DxpRefHyperlink(name, target!);
                                    }
                                }
                            }
                            break;
                        case FieldCode code:
                            TryCollectSeqIdentifier(code.Text, seqIdentifiers);
                            break;
                        case SimpleField simple:
                            if (simple.Instruction != null)
                                TryCollectSeqIdentifier(simple.Instruction.Value, seqIdentifiers);
                            break;
                    }
                }
            }
        }

        var resolvedBookmarks = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var kvp in bookmarks)
        {
            string text = kvp.Value.ToString();
            resolvedBookmarks[kvp.Key] = text;
            refBookmarks[kvp.Key] = new DxpRefBookmark(
                kvp.Key,
                text,
                bookmarkOrder.TryGetValue(kvp.Key, out var ord) ? ord : 0);
        }

        var refIndex = new DxpRefIndex(
            refBookmarks,
            refFootnotes,
            refEndnotes,
            refHyperlinks,
            paragraphNumbers);

        return new DxpDocumentIndex(resolvedBookmarks, seqIdentifiers, refIndex, captions);
    }

    private static Dictionary<long, (string Mark, string Text)> BuildFootnoteIndex(WordprocessingDocument doc)
    {
        var result = new Dictionary<long, (string, string)>();
        var footnotes = doc.MainDocumentPart?.FootnotesPart?.Footnotes;
        if (footnotes == null)
            return result;

        int index = 0;
        foreach (var fn in footnotes.Elements<Footnote>())
        {
            var id = fn.Id?.Value;
            if (id == null)
                continue;

            var type = fn.Type?.Value;
            if (type == FootnoteEndnoteValues.Separator ||
                type == FootnoteEndnoteValues.ContinuationSeparator ||
                type == FootnoteEndnoteValues.ContinuationNotice)
                continue;

            index++;
            var text = ExtractFootnoteText(fn);
            result[id.Value] = (index.ToString(System.Globalization.CultureInfo.InvariantCulture), text);
        }

        return result;
    }

    private static Dictionary<long, (string Mark, string Text)> BuildEndnoteIndex(WordprocessingDocument doc)
    {
        var result = new Dictionary<long, (string, string)>();
        var endnotes = doc.MainDocumentPart?.EndnotesPart?.Endnotes;
        if (endnotes == null)
            return result;

        int index = 0;
        foreach (var en in endnotes.Elements<Endnote>())
        {
            var id = en.Id?.Value;
            if (id == null)
                continue;

            var type = en.Type?.Value;
            if (type == FootnoteEndnoteValues.Separator ||
                type == FootnoteEndnoteValues.ContinuationSeparator ||
                type == FootnoteEndnoteValues.ContinuationNotice)
                continue;

            index++;
            var text = ExtractEndnoteText(en);
            result[id.Value] = (index.ToString(System.Globalization.CultureInfo.InvariantCulture), text);
        }

        return result;
    }

    private static string ExtractEndnoteText(Endnote endnote)
    {
        var sb = new StringBuilder();
        foreach (var text in endnote.Descendants<Text>())
            sb.Append(text.Text);
        return sb.ToString().Trim();
    }

    private static string? ResolveHyperlinkTarget(Hyperlink link)
    {
        if (!string.IsNullOrWhiteSpace(link.Anchor?.Value))
            return link.Anchor!.Value;
        if (!string.IsNullOrWhiteSpace(link.Id?.Value))
            return link.Id!.Value;
        return null;
    }

    private static string ExtractFootnoteText(Footnote footnote)
    {
        var sb = new StringBuilder();
        foreach (var text in footnote.Descendants<Text>())
            sb.Append(text.Text);
        return sb.ToString().Trim();
    }

    private static void TryCollectSeqIdentifier(string? instruction, List<string> seqIdentifiers)
    {
        if (TryParseSeqIdentifier(instruction, out var identifier) && identifier != null)
            seqIdentifiers.Add(identifier);
    }

    private static bool TryParseSeqIdentifier(string? instruction, out string? identifier)
    {
        identifier = null;
        if (string.IsNullOrWhiteSpace(instruction))
            return false;

        var text = instruction!.Trim();
        if (!text.StartsWith("SEQ", StringComparison.OrdinalIgnoreCase))
            return false;

        var parts = text.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2)
            return false;

        var value = parts[1];
        if (string.IsNullOrWhiteSpace(value))
            return false;
        identifier = value;
        return true;
    }

    private static string? TryExtractSeqIdentifier(Paragraph paragraph)
    {
        foreach (var code in paragraph.Descendants<FieldCode>())
        {
            if (TryParseSeqIdentifier(code.Text, out var identifier))
                return identifier;
        }
        foreach (var simple in paragraph.Descendants<SimpleField>())
        {
            if (TryParseSeqIdentifier(simple.Instruction?.Value, out var identifier))
                return identifier;
        }
        return null;
    }

    private static bool IsCaptionParagraph(Paragraph paragraph, DxpStyleResolver styles)
    {
        var chain = styles.GetParagraphStyleChain(paragraph);
        foreach (var info in chain)
        {
            if (string.Equals(info.StyleId, DxpWordBuiltInStyleId.wdStyleCaption, StringComparison.OrdinalIgnoreCase))
                return true;
            if (string.Equals(info.Name, "Caption", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static string ExtractParagraphText(Paragraph paragraph)
    {
        var sb = new StringBuilder();
        foreach (var text in paragraph.Descendants<Text>())
            sb.Append(text.Text);
        return sb.ToString().Trim();
    }

    private static string? NormalizeMarker(string? marker)
    {
        if (string.IsNullOrWhiteSpace(marker))
            return null;
        return marker!.Trim();
    }

    private static string ExtractCurrentLevel(string fullNumber)
    {
        var parts = fullNumber.Split('.');
        for (int i = parts.Length - 1; i >= 0; i--)
        {
            var part = parts[i].Trim();
            if (part.Length > 0)
                return part;
        }
        return fullNumber;
    }

    private static string StripNonNumeric(string text)
    {
        var sb = new StringBuilder(text.Length);
        foreach (char ch in text)
        {
            if (char.IsDigit(ch) || ch == '.')
                sb.Append(ch);
        }
        return sb.ToString();
    }
}
