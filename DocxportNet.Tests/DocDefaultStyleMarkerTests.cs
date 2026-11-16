using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Doc;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class DocDefaultStyleMarkerTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 8 })
            foreach (var mode in new[] { "normal-implicit", "normal-explicit", "custom-implicit", "custom-explicit", "decoy-implicit", "decoy-explicit", "custom-cancelled" })
                yield return new object[] { depth, mode };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void DefaultParagraphStyleRetainsItsMarkersThroughDoc(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        using var sourceStream = new MemoryStream(bytes);
        using var source = WordprocessingDocument.Open(sourceStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(source));
        var main = source.MainDocumentPart!;
        var before = new[] { main.Document.OuterXml, main.StyleDefinitionsPart!.Styles!.OuterXml, main.NumberingDefinitionsPart!.Numbering!.OuterXml };
        var binary = DxpExport.ExportToBytes(source, new DxpDocVisitor(), options: null);
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(projected));
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
                Assert.Equal(string.Concat(Labels(depth, mode).Select((marker, i) => marker + "Item" + i)), Regex.Replace(output, @"\s+", ""));
            }
        }
        Assert.Equal(before, new[] { main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml, main.NumberingDefinitionsPart.Numbering.OuterXml });
    }

    public static string[] Labels(int depth, string mode) => mode.EndsWith("cancelled", StringComparison.Ordinal)
        ? ["", "", ""] : ["[7]", "[8]", "[9]"];

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var id = mode.StartsWith("normal", StringComparison.Ordinal) ? "Normal" : "DefaultList";
            var styles = new Styles(new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new FontSize { Val = "24" }))));
            if (mode.StartsWith("decoy", StringComparison.Ordinal))
                styles.Append(new Style(new StyleName { Val = "Normal" }) { StyleId = "Normal", Type = StyleValues.Paragraph });
            styles.Append(new Style(new StyleName { Val = "Default List" }, new StyleParagraphProperties(
                new NumberingProperties(new NumberingLevelReference { Val = depth }, new NumberingId { Val = 1 })))
                { StyleId = id, Type = StyleValues.Paragraph, Default = true });
            main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
                new AbstractNum(new Level(new StartNumberingValue { Val = 7 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal }, new LevelText { Val = $"[%{depth + 1}]" })
                    { LevelIndex = depth }) { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
            var body = new Body();
            for (var i = 0; i < 3; i++)
            {
                var paragraph = new Paragraph();
                if (mode.EndsWith("explicit", StringComparison.Ordinal))
                    paragraph.Append(new ParagraphProperties(new ParagraphStyleId { Val = id }));
                else if (mode.EndsWith("cancelled", StringComparison.Ordinal))
                    paragraph.Append(new ParagraphProperties(new NumberingProperties(new NumberingId { Val = 0 })));
                paragraph.Append(new Run(new Text("Item" + i)));
                body.Append(paragraph);
            }
            main.Document = new Document(body);
        }
        return stream.ToArray();
    }
}
