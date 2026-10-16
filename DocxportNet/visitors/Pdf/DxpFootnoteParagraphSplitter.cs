using System.Text;
using System.Text.RegularExpressions;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Fields;

namespace DocxportNet.Visitors.Pdf;

internal static class DxpFootnoteParagraphSplitter
{
    internal static bool MatchesRenderedPrefix(string content, IReadOnlyList<string> renderedWords)
    {
        var sourceWords = Regex.Matches(content, @"\S+").Select(match => match.Value).ToArray();
        return renderedWords.Count > 0 && renderedWords.Count < sourceWords.Length &&
            sourceWords.Take(renderedWords.Count).SequenceEqual(renderedWords);
    }

    internal static bool TryMapRenderedSplit(string content, IReadOnlyList<string> renderedWords,
        int fittingRenderedWords, out int fittingSourceWords,
        IReadOnlyCollection<string>? footnoteMarkers = null)
    {
        if (TryMapExactRenderedSplit(content, renderedWords, fittingRenderedWords,
                out fittingSourceWords)) return true;
        fittingSourceWords = 0;
        if (footnoteMarkers == null || footnoteMarkers.Count == 0 ||
            fittingRenderedWords <= 0 || fittingRenderedWords > renderedWords.Count) return false;
        var markers = footnoteMarkers.Where(marker => marker.Length > 0).Distinct()
            .OrderByDescending(marker => marker.Length).ToArray();
        if (markers.Length == 0) return false;
        var sourceWords = Regex.Matches(content, @"\S+").Select(match => match.Value).ToArray();
        var normalizedRendered = renderedWords.Where(word => !markers.Contains(word)).ToArray();
        var normalizedSource = new string[sourceWords.Length];
        var renderedIndex = 0;
        for (var index = 0; index < sourceWords.Length; index++)
        {
            var word = sourceWords[index];
            if (renderedIndex < normalizedRendered.Length && word == normalizedRendered[renderedIndex])
            {
                normalizedSource[index] = word;
                renderedIndex++;
                continue;
            }
            var marker = markers.FirstOrDefault(value => word.EndsWith(value, StringComparison.Ordinal));
            if (marker == null) return false;
            var contentWord = word[..^marker.Length];
            if (contentWord.Length > 0 &&
                (renderedIndex >= normalizedRendered.Length || contentWord != normalizedRendered[renderedIndex]))
                return false;
            normalizedSource[index] = contentWord;
            if (contentWord.Length > 0) renderedIndex++;
        }
        if (!normalizedSource.Where(word => word.Length > 0).SequenceEqual(normalizedRendered)) return false;
        var fittingContentWords = renderedWords.Take(fittingRenderedWords)
            .Count(word => !markers.Contains(word));
        if (fittingContentWords <= 0 || fittingContentWords >= normalizedRendered.Length) return false;
        var contentWords = 0;
        for (var index = 0; index < normalizedSource.Length; index++)
        {
            if (normalizedSource[index].Length > 0) contentWords++;
            if (contentWords != fittingContentWords) continue;
            fittingSourceWords = index + 1;
            return fittingSourceWords < sourceWords.Length;
        }
        return false;
    }

    private static bool TryMapExactRenderedSplit(string content, IReadOnlyList<string> renderedWords,
        int fittingRenderedWords, out int fittingSourceWords)
    {
        fittingSourceWords = 0;
        var sourceWords = Regex.Matches(content, @"\S+").Select(match => match.Value).ToArray();
        if (fittingRenderedWords <= 0 || fittingRenderedWords > renderedWords.Count ||
            renderedWords.Count == 0 || sourceWords.Length <= 1) return false;
        if (renderedWords.Count <= sourceWords.Length)
        {
            if (!sourceWords.Take(fittingRenderedWords)
                    .SequenceEqual(renderedWords.Take(fittingRenderedWords))) return false;
            fittingSourceWords = fittingRenderedWords;
            return fittingSourceWords < sourceWords.Length;
        }

        var renderedText = string.Concat(renderedWords);
        var sourceText = string.Concat(sourceWords);
        if (renderedText.Length != sourceText.Length ||
            !renderedText.Where(character => !char.IsDigit(character))
                .SequenceEqual(sourceText.Where(character => !char.IsDigit(character))) ||
            !renderedText.Where(char.IsDigit).SequenceEqual(sourceText.Where(char.IsDigit)))
            return false;
        var offset = renderedWords.Take(fittingRenderedWords).Sum(word => word.Length);
        var renderedPrefix = renderedText[..offset];
        var sourcePrefix = sourceText[..offset];
        if (!renderedPrefix.Where(character => !char.IsDigit(character))
                .SequenceEqual(sourcePrefix.Where(character => !char.IsDigit(character))) ||
            !renderedPrefix.Where(char.IsDigit).SequenceEqual(sourcePrefix.Where(char.IsDigit)))
            return false;
        var sourceOffset = 0;
        for (var index = 0; index < sourceWords.Length; index++)
        {
            sourceOffset += sourceWords[index].Length;
            if (sourceOffset != offset) continue;
            fittingSourceWords = index + 1;
            return fittingSourceWords < sourceWords.Length;
        }
        return false;
    }

    internal static bool TrySplit(Paragraph paragraph, int firstFragmentWords,
        out Paragraph? continuation)
    {
        continuation = null;
        if (!TryGetTextElements(paragraph, out var textElements, out var bookmarks,
                out var characters, out var content)) return false;
        var words = Regex.Matches(content, @"\S+");
        if (firstFragmentWords <= 0 || firstFragmentWords >= words.Count) return false;

        var split = words[firstFragmentWords].Index;
        var clone = (Paragraph)paragraph.Clone();
        if (!TryGetTextElements(clone, out var clonedElements, out var clonedBookmarks,
                out var clonedCharacters, out _) ||
            clonedElements.Count != textElements.Count || clonedBookmarks.Count != bookmarks.Count ||
            clonedCharacters.Count != characters.Count)
            return false;

        for (var index = 0; index < textElements.Count; index++)
        {
            var text = textElements[index].Text;
            var value = text.Content;
            var localSplit = Math.Clamp(split - textElements[index].Offset, 0, value.Length);
            text.Content = value[..localSplit];
            clonedElements[index].Text.Content = value[localSplit..];
        }
        for (var index = 0; index < bookmarks.Count; index++)
        {
            if (bookmarks[index].Offset < split)
                RemoveBookmark(clonedBookmarks[index].Owner, clonedBookmarks[index].Field);
            else
                RemoveBookmark(bookmarks[index].Owner, bookmarks[index].Field);
        }
        for (var index = 0; index < characters.Count; index++)
        {
            if (characters[index].Offset < split)
                RemoveInline(clonedCharacters[index].Owner, clonedCharacters[index].Character);
            else
                RemoveInline(characters[index].Owner, characters[index].Character);
        }
        paragraph.Format.SpaceAfter = Unit.Zero;
        clone.Format.SpaceBefore = Unit.Zero;
        clone.Format.FirstLineIndent = Unit.Zero;
        clone.Format.PageBreakBefore = true;
        paragraph.Format.Borders.Bottom.Visible = false;
        clone.Format.Borders.Top.Visible = false;
        continuation = clone;
        return true;
    }

    internal static bool TryGetText(Paragraph paragraph, out string content)
    {
        return TryGetTextElements(paragraph, out _, out _, out _, out content);
    }

    private static bool TryGetTextElements(Paragraph paragraph,
        out List<(Text Text, int Offset)> elements,
        out List<(BookmarkField Field, int Offset, DocumentObject Owner)> bookmarks,
        out List<(Character Character, int Offset, DocumentObject Owner)> characters,
        out string content)
    {
        elements = new List<(Text, int)>();
        bookmarks = new List<(BookmarkField, int, DocumentObject)>();
        characters = new List<(Character, int, DocumentObject)>();
        var builder = new StringBuilder();
        var offset = 0;
        foreach (var element in paragraph.Elements)
            if (!AppendTextElements(element, paragraph, elements, bookmarks, characters,
                    builder, ref offset))
            {
                content = string.Empty;
                return false;
            }
        content = builder.ToString();
        return elements.Count > 0;
    }

    private static bool AppendTextElements(DocumentObject element, DocumentObject owner,
        List<(Text Text, int Offset)> elements,
        List<(BookmarkField Field, int Offset, DocumentObject Owner)> bookmarks,
        List<(Character Character, int Offset, DocumentObject Owner)> characters,
        StringBuilder builder, ref int offset)
    {
        if (element is Text text)
        {
            elements.Add((text, offset));
            builder.Append(text.Content);
            offset += text.Content.Length;
            return true;
        }
        if (element is Character { SymbolName: SymbolName.Tab } character)
        {
            characters.Add((character, offset, owner));
            builder.Append(' ');
            offset++;
            return true;
        }
        if (element is BookmarkField bookmark)
        {
            bookmarks.Add((bookmark, offset, owner));
            return true;
        }
        var children = element switch
        {
            FormattedText formatted => formatted.Elements,
            Hyperlink hyperlink => hyperlink.Elements,
            _ => null
        };
        if (children == null) return false;
        foreach (var child in children)
            if (!AppendTextElements(child, element, elements, bookmarks, characters,
                    builder, ref offset)) return false;
        return true;
    }

    private static void RemoveBookmark(DocumentObject owner, BookmarkField bookmark)
    {
        RemoveInline(owner, bookmark);
    }

    private static void RemoveInline(DocumentObject owner, DocumentObject inline)
    {
        if (owner is Paragraph paragraph)
        {
            for (var index = 0; index < paragraph.Elements.Count; index++)
                if (ReferenceEquals(paragraph.Elements[index], inline))
                {
                    paragraph.Elements.RemoveObjectAt(index);
                    return;
                }
        }
        else if (owner is FormattedText formatted)
            for (var index = 0; index < formatted.Elements.Count; index++)
                if (ReferenceEquals(formatted.Elements[index], inline))
                {
                    formatted.Elements.RemoveObjectAt(index);
                    return;
                }
        else if (owner is Hyperlink hyperlink)
            for (var childIndex = 0; childIndex < hyperlink.Elements.Count; childIndex++)
                if (ReferenceEquals(hyperlink.Elements[childIndex], inline))
                {
                    hyperlink.Elements.RemoveObjectAt(childIndex);
                    return;
                }
    }
}
