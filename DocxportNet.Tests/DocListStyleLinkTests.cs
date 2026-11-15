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

public sealed class DocListStyleLinkTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 1, 8 })
            foreach (var mode in new[] { "linked", "unlinked" })
                yield return new object[] { depth, mode };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void ExplicitListLevelStyleLinksSurviveDoc(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var location = index.Structure.FindLocation("ListDefinitions")!;
            var node = new DocStructureNode("FibLocation", "ListDefinitions", location.StreamName, location.Offset, location.Length);
            DocListNavigator.Expand(index.Structure, node);
            var record = node.Children.Single().Children.Single(x => x.Kind == "LSTF");
            var bytesAtLevel = index.Structure.ReadRange(record.StreamName!, record.Offset!.Value + 8 + depth * 2, 2);
            var styleIndex = System.Buffers.Binary.BinaryPrimitives.ReadUInt16LittleEndian(bytesAtLevel);
            Assert.Equal(mode == "linked" ? 15 : 0x0fff, styleIndex);
        }
        using var sourceStream = new MemoryStream(bytes);
        using var source = WordprocessingDocument.Open(sourceStream, false);
        var numbering = source.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var original = numbering.OuterXml;
        var styles = source.MainDocumentPart.StyleDefinitionsPart!.Styles!;
        var originalStyles = styles.OuterXml;
        _ = DocOpenXmlListReader.Read(source.MainDocumentPart, _ => DocCharacterFormatting.Empty, _ => DocParagraphFormatting.Empty);
        Assert.Equal(original, numbering.OuterXml);
        Assert.Equal(originalStyles, styles.OuterXml);
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(projected));
        var projectedLevel = projected.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.Elements<AbstractNum>()
            .Single().Elements<Level>().Single(x => x.LevelIndex!.Value == depth);
        var linkedId = projectedLevel.ParagraphStyleIdInLevel?.Val?.Value;
        if (mode == "linked")
            Assert.Equal("Nested", projected.MainDocumentPart.StyleDefinitionsPart!.Styles!.Elements<Style>()
                .Single(x => x.StyleId?.Value == linkedId).StyleName!.Val!.Value);
        else Assert.Null(linkedId);
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

    // Editable markers and ListTemplate linked style verified in Word on 2026-11-15.
    public static string[] Labels(int depth, string mode) =>
        ["BEGIN2END", "BEGIN5END", "BEGIN6END", "BEGIN3END", "BEGIN5END", "BEGIN6END"];

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(DocStyleLevelContextTests.CreateSource(depth, "full-style"));
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var level = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.Elements<AbstractNum>()
                .Single().Elements<Level>().Single(x => x.LevelIndex!.Value == depth);
            if (mode == "linked") level.InsertAfter(new ParagraphStyleIdInLevel { Val = "Nested" }, level.NumberingFormat!);
        }
        return stream.ToArray();
    }
}
