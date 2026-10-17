using DocumentFormat.OpenXml.Wordprocessing;
using MigraDocParagraph = MigraDoc.DocumentObjectModel.Paragraph;

namespace DocxportNet.Visitors.Pdf;

internal sealed class DxpPdfParagraphFragments
{
    internal sealed record Fragment(Paragraph Source, MigraDocParagraph Root, int RenderedTextStart,
        int Index, bool StyleRefCandidate);

    private readonly Dictionary<MigraDocParagraph, Fragment> _fragments = new();

    internal void Clear() => _fragments.Clear();

    internal void RegisterSource(MigraDocParagraph paragraph, Paragraph source, bool styleRefCandidate)
        => _fragments.Add(paragraph, new Fragment(source, paragraph, 0, 0, styleRefCandidate));

    internal void RegisterContinuation(MigraDocParagraph current, MigraDocParagraph next, int textLength)
    {
        var previous = _fragments[current];
        _fragments.Add(next, previous with
        {
            RenderedTextStart = previous.RenderedTextStart + textLength,
            Index = previous.Index + 1
        });
    }

    internal bool TryGet(MigraDocParagraph paragraph, out Fragment fragment)
        => _fragments.TryGetValue(paragraph, out fragment!);

    internal MigraDocParagraph RootOf(MigraDocParagraph paragraph)
        => _fragments.TryGetValue(paragraph, out var fragment) ? fragment.Root : paragraph;

    internal bool HasInheritedFlag(HashSet<MigraDocParagraph> flags, MigraDocParagraph paragraph)
        => flags.Contains(RootOf(paragraph));

    internal bool TryGetInheritedValue<T>(Dictionary<MigraDocParagraph, T> values,
        MigraDocParagraph paragraph, out T value)
        => values.TryGetValue(paragraph, out value!) || values.TryGetValue(RootOf(paragraph), out value!);
}
