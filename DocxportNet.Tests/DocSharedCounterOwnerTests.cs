using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Doc;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class DocSharedCounterOwnerTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 1, 8 })
            foreach (var mode in new[] { "start-1", "start-7", "none", "distinct", "format", "both" })
                yield return new object[] { depth, mode };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void SharedCounterOwnershipSurvivesDocxProjection(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            Assert.Equal(mode == "distinct" ? 2 : 1, index.Lists.Definitions.Count);
            Assert.Equal(2, index.Lists.Instances.Count);
        }
        using var sourceStream = new MemoryStream(bytes);
        using var source = WordprocessingDocument.Open(sourceStream, false);
        var numbering = source.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var original = numbering.OuterXml;
        _ = DocOpenXmlListReader.Read(source.MainDocumentPart, _ => DocCharacterFormatting.Empty, _ => DocParagraphFormatting.Empty);
        Assert.Equal(original, numbering.OuterXml);
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(projected));
        var projectedNumbering = projected.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        Assert.Equal(mode == "distinct" ? 2 : 1, projectedNumbering.Elements<AbstractNum>().Count());
        Assert.Equal(2, projectedNumbering.Elements<NumberingInstance>().Count());
        Assert.Equal(mode is "format" or "both" ? 1 : 0,
            projectedNumbering.Descendants<LevelOverride>().Count(x => x.Level != null));
        var expected = Labels(depth, mode);
        foreach (var document in new[] { source, projected })
        {
            DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig()),
                new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
                new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()),
                new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
            foreach (var visitor in visitors)
            {
                var output = DxpExport.ExportToString(document, visitor);
                if (visitor is DxpHtmlVisitor or DxpMarkdownVisitor)
                    output = new HtmlParser().ParseDocument(output).Body!.TextContent;
                Assert.Equal(string.Concat(expected.Select((marker, i) => marker + "Item" + i)), Regex.Replace(output, @"\s+", ""));
            }
        }
    }

    public static string[] Labels(int depth, string mode) => (mode switch
    {
        "start-1" => new[] { "1", "2", "1", "2", "3", "4" },
        "start-7" => ["1", "2", "7", "8", "9", "10"],
        "distinct" => ["1", "2", "1", "2", "3", "3"],
        "format" => ["1", "2", "III", "IV", "5", "VI"],
        "both" => ["1", "2", "VII", "VIII", "9", "X"],
        _ => ["1", "2", "3", "4", "5", "6"]
    }).Select(value => "BEGIN" + value + "END").ToArray();

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(DocOverrideFlagTests.CreateSource(depth, "none"));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var numbering = main.NumberingDefinitionsPart!.Numbering!;
            foreach (var level in numbering.Descendants<Level>()) level.StartNumberingValue!.Val = 1;
            var number = new NumberingInstance(new AbstractNumId { Val = mode == "distinct" ? 2 : 1 }) { NumberID = 2 };
            if (mode == "distinct")
            {
                var definition = (AbstractNum)numbering.Elements<AbstractNum>().Single().CloneNode(true);
                definition.AbstractNumberId = 2;
                numbering.InsertBefore(definition, numbering.Elements<NumberingInstance>().First());
            }
            if (mode is "start-1" or "start-7" or "format" or "both")
            {
                var replacement = new LevelOverride { LevelIndex = depth };
                if (mode is "start-1" or "start-7" or "both")
                    replacement.AppendChild(new StartOverrideNumberingValue { Val = mode == "start-7" ? 7 : mode == "both" ? 7 : 1 });
                if (mode is "format" or "both")
                    replacement.AppendChild(new Level(new StartNumberingValue { Val = 7 },
                        new NumberingFormat { Val = NumberFormatValues.UpperRoman },
                        new LevelText { Val = "BEGIN%" + (depth + 1) + "END" }) { LevelIndex = depth });
                number.AppendChild(replacement);
            }
            numbering.AppendChild(number);
            var ids = new[] { 1, 1, 2, 2, 1, 2 };
            foreach (var (paragraph, i) in main.Document.Body!.Elements<Paragraph>().Select((p, i) => (p, i)))
                paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val = ids[i];
        }
        return stream.ToArray();
    }
}
