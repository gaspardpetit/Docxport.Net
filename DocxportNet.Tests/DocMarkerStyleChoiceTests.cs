using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Doc;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class DocMarkerStyleChoiceTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 8 })
            foreach (var scope in new[] { "style", "properties", "numbering", "id", "level", "default" })
                foreach (var branch in new[] { "choice", "fallback", "direct" })
                    yield return new object[] { depth, scope, branch };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void SelectedStyleMarkerPropertiesSurviveDoc(int depth, string scope, string branch)
    {
        using var input = new MemoryStream(CreateSource(depth, scope, branch));
        using var source = WordprocessingDocument.Open(input, false);
        var main = source.MainDocumentPart!;
        _ = main.Document.Descendants().ToArray();
        _ = main.StyleDefinitionsPart!.Styles!.Descendants().ToArray();
        _ = main.NumberingDefinitionsPart!.Numbering!.Descendants().ToArray();
        var before = new[] { main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml, main.NumberingDefinitionsPart.Numbering.OuterXml };
        var binary = DxpExport.ExportToBytes(source, new DxpDocVisitor(), options: null);
        Assert.Equal(before, new[] { main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml, main.NumberingDefinitionsPart.Numbering.OuterXml });
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(projected));
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
                Assert.Equal("[7]Item0[8]Item1[9]Item2", Regex.Replace(output, @"\s+", ""));
            }
        }
    }

    public static byte[] CreateSource(int depth, string scope, string branch)
    {
        string Select(string selected, string rejected) => branch switch
        {
            "direct" => selected,
            "choice" => $"<mc:AlternateContent><mc:Choice Requires=\"w\">{selected}</mc:Choice><mc:Fallback>{rejected}</mc:Fallback></mc:AlternateContent>",
            _ => $"<mc:AlternateContent><mc:Choice Requires=\"x\">{rejected}</mc:Choice><mc:Fallback>{selected}</mc:Fallback></mc:AlternateContent>"
        };
        var numPr = $"<w:numPr><w:ilvl w:val=\"{depth}\"/><w:numId w:val=\"1\"/></w:numPr>";
        var wrong = "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr>";
        string Style(string content) => $"<w:style w:type=\"paragraph\" w:styleId=\"List\"{(scope == "default" ? " w:default=\"1\"" : "")}><w:name w:val=\"List\"/>{content}</w:style>";
        var style = scope switch
        {
            "style" or "default" => Select(Style($"<w:pPr>{numPr}</w:pPr>"), Style($"<w:pPr>{wrong}</w:pPr>")),
            "properties" => Style(Select($"<w:pPr>{numPr}</w:pPr>", $"<w:pPr>{wrong}</w:pPr>")),
            "numbering" => Style("<w:pPr>" + Select(numPr, wrong) + "</w:pPr>"),
            "id" => Style($"<w:pPr><w:numPr><w:ilvl w:val=\"{depth}\"/>" + Select("<w:numId w:val=\"1\"/>", "<w:numId w:val=\"2\"/>") + "</w:numPr></w:pPr>"),
            _ => Style("<w:pPr><w:numPr>" + Select($"<w:ilvl w:val=\"{depth}\"/>", $"<w:ilvl w:val=\"{8 - depth}\"/>") + "<w:numId w:val=\"1\"/></w:numPr></w:pPr>")
        };
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body();
            for (var i = 0; i < 3; i++)
            {
                var paragraph = new Paragraph();
                if (scope != "default") paragraph.Append(new ParagraphProperties(new ParagraphStyleId { Val = "List" }));
                paragraph.Append(new Run(new Text("Item" + i)));
                body.Append(paragraph);
            }
            main.Document = new Document(body);
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles($"<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x=\"urn:marker-unsupported\" mc:Ignorable=\"x\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/><w:sz w:val=\"24\"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type=\"paragraph\" w:styleId=\"Normal\"{(scope == "default" ? "" : " w:default=\"1\"")}><w:name w:val=\"Normal\"/></w:style>{style}</w:styles>");
            var levels = Enumerable.Range(0, 9).Select(i => new Level(new StartNumberingValue { Val = i == depth ? 7 : 70 },
                new NumberingFormat { Val = NumberFormatValues.Decimal }, new LevelText { Val = i == depth ? $"[%{i + 1}]" : $"WRONG%{i + 1}" }) { LevelIndex = i });
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(new AbstractNum(levels) { AbstractNumberId = 1 },
                new AbstractNum(new Level(new StartNumberingValue { Val = 70 }, new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "(%1)" }) { LevelIndex = 0 }) { AbstractNumberId = 2 },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 }, new NumberingInstance(new AbstractNumId { Val = 2 }) { NumberID = 2 });
        }
        return stream.ToArray();
    }
}
