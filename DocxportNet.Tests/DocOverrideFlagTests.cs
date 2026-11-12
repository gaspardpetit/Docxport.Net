using System.Buffers.Binary;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Doc;
using DocxportNet.API;
using DocxportNet.Visitors.Markdown;

namespace DocxportNet.Tests;

public sealed class DocOverrideFlagTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 1, 8 })
            foreach (var mode in new[] { "none", "start", "format", "both" })
                yield return new object[] { depth, mode };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void IndependentStartAndFormattingFlagsSurviveBinaryRoundTrip(int depth, string mode)
    {
        var binary = DxpDocExport.Export(CreateSource(depth, mode));
        using var input = new MemoryStream(binary);
        using var index = new DocTextIndexWalker().Index(input);
        var instance = Assert.Single(index.Lists.Instances);
        var hasStart = mode is "start" or "both";
        var hasFormat = mode is "format" or "both";
        Assert.Equal(hasStart, instance.StartOverrides.ContainsKey(depth));
        Assert.Equal(hasFormat, instance.FormattingOverrides.ContainsKey(depth));
        if (hasStart) Assert.Equal(7, instance.StartOverrides[depth]);
        if (hasFormat) Assert.Equal(7, instance.FormattingOverrides[depth].StartAt);
        // Word's formatting-only controls retain the abstract start. Flattening
        // replacements in DOCX projection is a separate existing issue.
        if (mode != "format")
        {
            using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
            using var projected = WordprocessingDocument.Open(projectedStream, false);
            var output = DxpExport.ExportToString(projected,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Equal(Enumerable.Range(hasStart ? 7 : 2, 6).Select(n => "BEGIN" + n + "END"),
                Regex.Matches(output, "BEGIN.*?END").Select(m => m.Value));
        }
        var location = index.Structure.FindLocation("ListOverrides")!;
        var node = new DocStructureNode("FibLocation", "ListOverrides", location.StreamName,
            location.Offset, location.Length);
        DocListNavigator.Expand(index.Structure, node);
        var table = Assert.Single(node.Children);
        var records = table.Children.Where(x => x.Kind == "LFOData").SelectMany(x => x.Children)
            .Where(x => x.Kind == "LFOLVL").ToArray();
        if (mode == "none") { Assert.Empty(records); return; }
        var record = Assert.Single(records);
        var bytes = index.Structure.ReadRange(record.StreamName!, record.Offset!.Value, 8);
        Assert.Equal(depth | (hasStart ? 0x10 : 0) | (hasFormat ? 0x20 : 0), bytes[4]);
        // When both flags are set, LFOLVL.iStartAt is ignored; the nested LVL owns the start.
        Assert.Equal(hasStart && !hasFormat ? 7 : 0, BinaryPrimitives.ReadInt32LittleEndian(bytes));
    }

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            Level Level(int index, int start) => new(new StartNumberingValue { Val = start },
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "BEGIN%" + (index + 1) + "END" }) { LevelIndex = index };
            var number = new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 };
            if (mode != "none")
            {
                var replacement = new LevelOverride { LevelIndex = depth };
                if (mode is "start" or "both") replacement.AppendChild(new StartOverrideNumberingValue { Val = 7 });
                if (mode is "format" or "both") replacement.AppendChild(Level(depth, 7));
                number.AppendChild(replacement);
            }
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
                new AbstractNum(Enumerable.Range(0, depth + 1).Select(i => Level(i, 2))) { AbstractNumberId = 1 }, number);
            main.Document = new Document(new Body(Enumerable.Range(0, 6).Select(i =>
                new Paragraph(new ParagraphProperties(new NumberingProperties(
                    new NumberingLevelReference { Val = depth }, new NumberingId { Val = 1 })),
                    new Run(new Text("Item" + i))))));
        }
        return stream.ToArray();
    }
}
