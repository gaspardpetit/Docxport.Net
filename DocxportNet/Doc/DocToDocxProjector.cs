using System.Text;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Doc;

/// <summary>Reports which indexed content was represented by a basic DOCX projection.</summary>
public sealed record DocxProjectionCoverage(
    DocPartRange? ProjectedMainPart,
    IReadOnlyList<DocPartRange> DeferredParts,
    IReadOnlyList<DocLocation> DeferredLocations,
    IReadOnlyList<string> DeferredStreams,
    IReadOnlyDictionary<ushort, int> OmittedCharacters,
    IReadOnlyDictionary<ushort, int> ApproximateCharacters);

public sealed record DocxProjectionResult(byte[] DocxBytes, DocxProjectionCoverage Coverage);

/// <summary>Projects indexed main-story text into a minimal WordprocessingML package.</summary>
public sealed class DocToDocxProjector
{
    public DocxProjectionResult Project(DocTextIndex index)
    {
        if (index == null) throw new ArgumentNullException(nameof(index));
        var main = index.Parts.FirstOrDefault(x => x.Name == "Main");
        var spans = index.GetPartSpans("Main");
        ValidateCoverage(main, spans);

        var omitted = new SortedDictionary<ushort, int>();
        var approximate = new SortedDictionary<ushort, int>();
        HashSet<uint>? sectionEnds = null;
        using var output = new MemoryStream();
        using (var document = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document, true))
        {
            var part = document.AddMainDocumentPart();
            var body = new Body();
            var paragraph = new Paragraph();
            var pendingText = new StringBuilder();
            var endedWithParagraph = false;
            char? pendingHighSurrogate = null;

            void FlushText()
            {
                if (pendingText.Length == 0) return;
                paragraph.AppendChild(new Run(new Text(pendingText.ToString())
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
                pendingText.Clear();
            }

            foreach (var span in spans)
            {
                var text = span.Text;
                for (var offset = 0; offset < text.Length; offset++)
                {
                    var character = text[offset];
                    if (pendingHighSurrogate is char high)
                    {
                        pendingHighSurrogate = null;
                        if (char.IsLowSurrogate(character))
                        {
                            pendingText.Append(high).Append(character);
                            endedWithParagraph = false;
                            continue;
                        }
                        Count(omitted, high);
                    }
                    if (char.IsHighSurrogate(character))
                    {
                        pendingHighSurrogate = character;
                        continue;
                    }
                    switch (character)
                    {
                        case '\r':
                            FlushText();
                            body.AppendChild(paragraph);
                            paragraph = new Paragraph();
                            endedWithParagraph = true;
                            break;
                        case '\t':
                            FlushText();
                            paragraph.AppendChild(new Run(new TabChar()));
                            endedWithParagraph = false;
                            break;
                        case '\u000B':
                            FlushText();
                            paragraph.AppendChild(new Run(new Break()));
                            endedWithParagraph = false;
                            break;
                        case '\u0007': // Cell or row mark; table structure is not indexed yet.
                            FlushText();
                            paragraph.AppendChild(new Run(new TabChar()));
                            Count(approximate, character);
                            endedWithParagraph = false;
                            break;
                        case '\u000C': // A section mark occurs at a section boundary; otherwise this is a page break.
                            FlushText();
                            sectionEnds ??= new HashSet<uint>(index.Sections
                                .Take(Math.Max(0, index.Sections.Count - 1))
                                .Select(x => uint.Parse(x.Attributes["cpEnd"], CultureInfo.InvariantCulture)));
                            if (sectionEnds.Contains(span.CpStart + (uint)offset + 1))
                            {
                                body.AppendChild(paragraph);
                                paragraph = new Paragraph();
                                Count(approximate, character);
                                endedWithParagraph = true;
                            }
                            else
                            {
                                paragraph.AppendChild(new Run(new Break { Type = BreakValues.Page }));
                                endedWithParagraph = false;
                            }
                            break;
                        default:
                            if (IsXmlCharacter(character))
                            {
                                pendingText.Append(character);
                                endedWithParagraph = false;
                            }
                            else
                                Count(omitted, character);
                            break;
                    }
                }
            }
            if (pendingHighSurrogate is char unmatchedHigh)
                Count(omitted, unmatchedHigh);
            FlushText();
            if (!endedWithParagraph || !body.Elements<Paragraph>().Any())
                body.AppendChild(paragraph);
            part.Document = new Document(body);
            part.Document.Save();
        }

        var coverage = new DocxProjectionCoverage(main,
            index.Parts.Where(x => x.Name != "Main").ToArray(),
            index.Structure.Locations.Where(x => x.IsPresent && x.FibIndex != 33).ToArray(),
            EnumerateStreams(index.Structure.Root)
                .Where(x => x != "WordDocument" && x != index.Structure.FibBase.TableStreamName)
                .ToArray(),
            omitted, approximate);
        return new DocxProjectionResult(output.ToArray(), coverage);
    }

    private static void ValidateCoverage(DocPartRange? main, IReadOnlyList<DocIndexedTextSpan> spans)
    {
        if (main == null) return;
        var expected = main.CpStart;
        foreach (var span in spans)
        {
            if (span.CpStart != expected)
                throw new InvalidDataException("The main document text index has a gap or overlap.");
            expected = span.CpEnd;
        }
        if (expected != main.CpEnd)
            throw new InvalidDataException("The main document text index does not cover its full CP range.");
    }

    private static bool IsXmlCharacter(char value) =>
        value is '\t' or '\n' or '\r' || value >= ' ' &&
        (value < '\uD800' || value > '\uDFFF') &&
        value is not '\uFFFE' and not '\uFFFF';

    private static void Count(IDictionary<ushort, int> counts, char value)
    {
        var key = (ushort)value;
        counts[key] = counts.TryGetValue(key, out var existing) ? existing + 1 : 1;
    }

    private static IEnumerable<string> EnumerateStreams(DocStructureNode node)
    {
        if (node.Kind == "Stream" && node.Attributes.TryGetValue("path", out var path))
            yield return path;
        foreach (var child in node.Children)
            foreach (var stream in EnumerateStreams(child))
                yield return stream;
    }
}
