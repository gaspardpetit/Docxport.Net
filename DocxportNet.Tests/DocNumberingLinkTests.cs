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

public sealed class DocNumberingLinkTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 1, 8 })
            foreach (var mode in new[] { "valid", "inherited", "override", "missing-reverse" })
                yield return new object[] { depth, mode };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void LinkedLevelsSurviveDoc(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            Assert.Equal(2, index.Lists.Definitions.Count);
            Assert.Equal(2, index.Lists.Instances.Count);
        }
        using var sourceStream = new MemoryStream(bytes);
        using var source = WordprocessingDocument.Open(sourceStream, false);
        var numbering = source.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var original = numbering.OuterXml;
        var originalStyles = source.MainDocumentPart.StyleDefinitionsPart!.Styles!.OuterXml;
        _ = DocOpenXmlListReader.Read(source.MainDocumentPart, _ => DocCharacterFormatting.Empty, _ => DocParagraphFormatting.Empty);
        Assert.Equal(original, numbering.OuterXml);
        Assert.Equal(originalStyles, source.MainDocumentPart.StyleDefinitionsPart!.Styles!.OuterXml);
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(projected));
        var projectedNumbering = projected.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        Assert.Equal(2, projectedNumbering.Elements<AbstractNum>().Count());
        Assert.Equal(2, projectedNumbering.Elements<NumberingInstance>().Count());
        Assert.Empty(projectedNumbering.Descendants<LevelOverride>().Where(x => x.Level != null));
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

    public static string[] Labels(int depth, string mode) =>
        (mode switch
        {
            "override" => new[] { "1", "2", "7", "8", "3", "9" },
            "missing-reverse" => ["", "", "1", "2", "", "3"],
            _ => ["1", "2", "1", "2", "3", "3"]
        }).Select(value => value == "" ? "" : "BEGIN" + value + "END").ToArray();

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(DocSharedCounterOwnerTests.CreateSource(depth, "none"));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var numbering = main.NumberingDefinitionsPart!.Numbering!;
            var definition = numbering.Elements<AbstractNum>().Single();
            definition.AbstractNumberId = 2;
            if (mode != "missing-reverse") definition.InsertAt(new StyleLink { Val = "Linked" }, 0);
            numbering.InsertBefore(new AbstractNum(new NumberingStyleLink { Val = "Linked" })
                { AbstractNumberId = 1 }, definition);
            numbering.Elements<NumberingInstance>().Single(x => x.NumberID!.Value == 2).AbstractNumId!.Val = 2;
            if (mode == "override")
                numbering.Elements<NumberingInstance>().Single(x => x.NumberID!.Value == 2)
                    .AppendChild(new LevelOverride(new StartOverrideNumberingValue { Val = 7 }) { LevelIndex = depth });
            var linked = new Style(new StyleName { Val = "Linked" })
                { Type = StyleValues.Numbering, StyleId = "Linked" };
            var properties = new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 2 }));
            var styles = new Styles();
            if (mode == "inherited")
            {
                linked.AppendChild(new BasedOn { Val = "Base" });
                styles.AppendChild(new Style(new StyleName { Val = "Base" }, properties)
                    { Type = StyleValues.Numbering, StyleId = "Base" });
            }
            else linked.AppendChild(properties);
            styles.AppendChild(linked);
            main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
        }
        return stream.ToArray();
    }
}
