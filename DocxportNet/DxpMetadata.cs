using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Doc;
using DocxportNet.Walker.Context;

namespace DocxportNet;

/// <summary>Core document properties shared by DOC and DOCX.</summary>
public sealed record DxpCoreMetadata
{
    public string? Title { get; init; }
    public string? Subject { get; init; }
    public string? Creator { get; init; }
    public string? LastModifiedBy { get; init; }
    public string? Revision { get; init; }
    public string? Created { get; init; }
    public string? Modified { get; init; }
    public string? Description { get; init; }
    public string? Category { get; init; }
    public string? Keywords { get; init; }
}

/// <summary>Application and document statistics, when available in the source.</summary>
public sealed record DxpExtendedMetadata
{
    public string? Application { get; init; }
    public string? ApplicationVersion { get; init; }
    public string? Template { get; init; }
    public string? Pages { get; init; }
    public string? Words { get; init; }
    public string? Characters { get; init; }
    public string? Lines { get; init; }
    public string? Paragraphs { get; init; }
    public string? TotalTime { get; init; }
}

public sealed record DxpLanguageRatio(string Code, double Ratio);

/// <summary>Document metadata and the presence of revision markup and comments.</summary>
public sealed record DxpDocumentMetadata(
    DxpCoreMetadata CoreProperties,
    DxpExtendedMetadata? ExtendedProperties,
    IReadOnlyList<DxpLanguageRatio>? Language,
    bool HasTrackedChanges,
    bool HasComments);

/// <summary>Reads metadata from DOCX and binary DOC without altering the source.</summary>
public static class DxpMetadata
{
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly HashSet<string> RevisionNames = new(StringComparer.Ordinal)
    {
        "ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd",
        "moveToRangeStart", "moveToRangeEnd", "conflictIns", "conflictDel"
    };

    public static DxpDocumentMetadata Inspect(string path)
    {
        if (path == null) throw new ArgumentNullException(nameof(path));
        return Inspect(File.ReadAllBytes(path));
    }

    public static DxpDocumentMetadata Inspect(byte[] bytes)
    {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (DocInputProjection.IsBinaryDoc(bytes))
        {
            using var index = new DocTextIndexWalker().Index(new MemoryStream(bytes, writable: false));
            return InspectDoc(index);
        }
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        return Inspect(document);
    }

    public static DxpDocumentMetadata Inspect(WordprocessingDocument document)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var core = document.PackageProperties;
        var extended = document.ExtendedFilePropertiesPart?.Properties;
        var main = document.MainDocumentPart;
        var roots = EnumerateTextRoots(document).ToArray();
        bool hasComments = main?.WordprocessingCommentsPart?.Comments?.Elements<Comment>().Any() == true ||
            roots.Any(root => root.Descendants<CommentReference>().Any());
        var metadata = new DxpCoreMetadata
        {
            Title = core.Title, Subject = core.Subject, Creator = core.Creator,
            LastModifiedBy = core.LastModifiedBy, Revision = core.Revision,
            Created = FormatDate(core.Created), Modified = FormatDate(core.Modified),
            Description = core.Description, Category = core.Category, Keywords = core.Keywords
        };
        var statistics = extended == null ? null : new DxpExtendedMetadata
        {
            Application = extended.Application?.Text,
            ApplicationVersion = extended.ApplicationVersion?.Text,
            Template = extended.Template?.Text,
            Pages = extended.Pages?.Text, Words = extended.Words?.Text,
            Characters = extended.Characters?.Text, Lines = extended.Lines?.Text,
            Paragraphs = extended.Paragraphs?.Text, TotalTime = extended.TotalTime?.Text
        };
        return new DxpDocumentMetadata(metadata, statistics, ReadDocxLanguages(document),
            roots.Any(root => root.Descendants().Any(element =>
                element.NamespaceUri == WordNamespace && RevisionNames.Contains(element.LocalName))),
            hasComments);
    }

    private static DxpDocumentMetadata InspectDoc(DocTextIndex index)
    {
        var structure = index.Structure;
        var core = new DxpCoreMetadata
        {
            Title = DocSummaryInformation.ReadTitle(structure),
            Subject = DocSummaryInformation.ReadSubject(structure),
            Creator = DocSummaryInformation.ReadAuthor(structure),
            LastModifiedBy = DocSummaryInformation.ReadLastAuthor(structure),
            Revision = DocSummaryInformation.ReadRevisionNumber(structure),
            Created = FormatDate(DocSummaryInformation.ReadCreated(structure)),
            Modified = FormatDate(DocSummaryInformation.ReadModified(structure)),
            Description = DocSummaryInformation.ReadComments(structure),
            Keywords = DocSummaryInformation.ReadKeywords(structure)
        };
        var pages = DocSummaryInformation.ReadPageCount(structure);
        var words = DocSummaryInformation.ReadWordCount(structure);
        var characters = DocSummaryInformation.ReadCharacterCount(structure);
        DxpExtendedMetadata? extended = pages == null && words == null && characters == null
            ? null : new DxpExtendedMetadata
            {
                Pages = pages?.ToString(CultureInfo.InvariantCulture),
                Words = words?.ToString(CultureInfo.InvariantCulture),
                Characters = characters?.ToString(CultureInfo.InvariantCulture)
            };
        var comments = structure.FindLocation("CommentReferences");
        return new DxpDocumentMetadata(core, extended, ReadDocLanguages(index),
            index.CharacterFormatting.Any(range => range.Formatting.InsertedRevision == true ||
                range.Formatting.DeletedRevision == true),
            comments?.IsPresent == true && comments.Length > 4);
    }

    private static IReadOnlyList<DxpLanguageRatio>? ReadDocxLanguages(WordprocessingDocument document)
    {
        var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var styles = new DxpStyleResolver(document);
        foreach (var root in EnumerateTextRoots(document))
            foreach (var paragraph in root.Descendants<Paragraph>())
                foreach (var run in paragraph.Descendants<Run>())
                {
                    var letters = CountLetters(run.InnerText);
                    if (letters == 0) continue;
                    var language = run.RunProperties?.Languages;
                    var code = NormalizeLanguage(language?.Val?.Value) ??
                        NormalizeLanguage(language?.EastAsia?.Value) ??
                        NormalizeLanguage(language?.Bidi?.Value) ??
                        NormalizeLanguage(styles.ResolveRunLanguage(paragraph, run)) ??
                        NormalizeLanguage(styles.ResolveParagraphLanguage(paragraph));
                    Add(counts, code, letters);
                }
        return ToRatios(counts);
    }

    private static IReadOnlyList<DxpLanguageRatio>? ReadDocLanguages(DocTextIndex index)
    {
        var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var ranges = index.CharacterFormatting.OrderBy(range => range.CpStart).ToArray();
        var defaultLanguage = index.DefaultCharacterFormatting.LanguageId;
        foreach (var part in index.Structure.Parts.Where(part => part.Name is
            "Main" or "Headers" or "Footnotes" or "Endnotes" or "Comments"))
        {
            var rangeIndex = 0;
            foreach (var span in index.GetPartSpans(part.Name))
            {
                var text = span.Text;
                for (var i = 0; i < text.Length; i++)
                {
                    if (!char.IsLetter(text[i])) continue;
                    var cp = span.CpStart + (uint)i;
                    while (rangeIndex < ranges.Length && ranges[rangeIndex].CpEnd <= cp)
                        rangeIndex++;
                    var languageId = rangeIndex < ranges.Length && ranges[rangeIndex].CpStart <= cp
                        ? ranges[rangeIndex].Formatting.LanguageId ?? defaultLanguage
                        : defaultLanguage;
                    if (languageId is not ushort id) continue;
                    string? code;
                    try { code = NormalizeLanguage(CultureInfo.GetCultureInfo(id).Name); }
                    catch (CultureNotFoundException) { continue; }
                    Add(counts, code, 1);
                }
            }
        }
        return ToRatios(counts);
    }

    private static IEnumerable<OpenXmlElement> EnumerateTextRoots(WordprocessingDocument document)
    {
        var main = document.MainDocumentPart;
        if (main?.Document != null) yield return main.Document;
        if (main == null) yield break;
        foreach (var part in main.HeaderParts)
            if (part.Header != null) yield return part.Header;
        foreach (var part in main.FooterParts)
            if (part.Footer != null) yield return part.Footer;
        if (main.FootnotesPart?.Footnotes != null) yield return main.FootnotesPart.Footnotes;
        if (main.EndnotesPart?.Endnotes != null) yield return main.EndnotesPart.Endnotes;
    }

    private static string? FormatDate(DateTime? date) =>
        date?.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture);

    private static int CountLetters(string? text) =>
        text?.Count(char.IsLetter) ?? 0;

    private static string? NormalizeLanguage(string? language)
    {
        if (string.IsNullOrWhiteSpace(language) ||
            string.Equals(language, "none", StringComparison.OrdinalIgnoreCase)) return null;
        var value = language!;
        var separator = value.IndexOfAny(new[] { '-', '_' });
        return (separator < 0 ? value : value.Substring(0, separator)).Trim().ToLowerInvariant();
    }

    private static void Add(IDictionary<string, int> counts, string? code, int letters)
    {
        if (string.IsNullOrEmpty(code)) return;
        var key = code!;
        counts[key] = counts.TryGetValue(key, out var current) ? current + letters : letters;
    }

    private static IReadOnlyList<DxpLanguageRatio>? ToRatios(Dictionary<string, int> counts)
    {
        var total = counts.Values.Sum();
        return total == 0 ? null : counts.OrderByDescending(item => item.Value)
            .ThenBy(item => item.Key, StringComparer.Ordinal)
            .Select(item => new DxpLanguageRatio(item.Key, Math.Round((double)item.Value / total, 4)))
            .ToArray();
    }
}

