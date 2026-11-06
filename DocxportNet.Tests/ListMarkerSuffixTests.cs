using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class ListMarkerSuffixTests
{
    public static IEnumerable<object[]> Cases() =>
        from format in new[] { "markdown", "html", "text" }
        from suffix in new[] { "nothing", "space", "tab", "omitted" }
        select new object[] { format, suffix };

    [Theory]
    [MemberData(nameof(Cases))]
    public void ExportersHonorNothingAndRetainWhitespaceNormalization(string format, string suffix)
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var level = new Level(new StartNumberingValue { Val = 1 },
            new NumberingFormat { Val = NumberFormatValues.Decimal }, new LevelText { Val = "[%1]" })
            { LevelIndex = 0 };
        if (suffix != "omitted") level.Append(new LevelSuffix { Val = suffix switch
            { "nothing" => LevelSuffixValues.Nothing, "space" => LevelSuffixValues.Space, _ => LevelSuffixValues.Tab } });
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
            new AbstractNum(level) { AbstractNumberId = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
        main.Document = new Document(new Body(new Paragraph(new ParagraphProperties(
            new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })),
            new Run(new Text("Text")))));
        DxpITextVisitor visitor = format switch
        {
            "html" => new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            "text" => new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()),
            _ => new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig())
        };
        var output = DxpExport.ExportToString(document, visitor);
        if (format == "html") output = new HtmlParser().ParseDocument(output).Body!.TextContent;
        Assert.Equal(suffix == "nothing" ? "[1]Text" : "[1] Text", output.Trim());
    }
}
