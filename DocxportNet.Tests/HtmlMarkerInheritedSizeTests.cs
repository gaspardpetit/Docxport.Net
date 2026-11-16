using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Visitors.Html;

namespace DocxportNet.Tests;

public sealed class HtmlMarkerInheritedSizeTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var lane in new[] { "mark", "style", "defaults" })
            foreach (var levelProperties in new[] { "absent", "empty", "fonts" })
                foreach (var literal in new[] { false, true })
                    foreach (var rtl in new[] { false, true })
                        yield return new object[] { lane, levelProperties, literal, rtl };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void InheritsSizeWithoutRequiringLevelRunProperties(string lane, string levelProperties, bool literal, bool rtl)
    {
        using var stream = new MemoryStream(CreateSource(lane, levelProperties, literal, rtl));
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(document));
        var main = document.MainDocumentPart!;
        var styles = main.StyleDefinitionsPart!.Styles!;
        var numbering = main.NumberingDefinitionsPart!.Numbering!;
        _ = main.Document!.Descendants().ToArray();
        _ = styles.Descendants().ToArray();
        _ = numbering.Descendants().ToArray();
        string[] Snapshot() => [main.Document.OuterXml, styles.OuterXml, numbering.OuterXml];
        var before = Snapshot();
        foreach (var rich in new[] { true, false })
        {
            var config = rich ? DxpHtmlVisitorConfig.CreateRichConfig() : DxpHtmlVisitorConfig.CreatePlainConfig();
            var html = DxpExport.ExportToString(document, new DxpHtmlVisitor(config));
            var markers = new HtmlParser().ParseDocument(html).QuerySelectorAll(".dxp-marker");
            Assert.Equal(3, markers.Length);
            Assert.Equal(literal ? new[] { "12", "12", "12" } : new[] { "1", "2", "3" },
                markers.Select(marker => marker.TextContent.Trim()));
            foreach (var marker in markers)
                if (rich && rtl) Assert.Contains("font-size:15pt", marker.GetAttribute("style"));
                else Assert.DoesNotContain("font-size:15pt", marker.OuterHtml);
        }
        Assert.Equal(before, Snapshot());
    }

    public static byte[] CreateSource(string lane, string levelProperties, bool literal, bool rtl)
    {
        static RunFonts Fonts() => new() { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" };
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var defaults = new RunPropertiesBaseStyle(Fonts(), new FontSize { Val = "24" });
            var styleRun = new StyleRunProperties(Fonts(), new FontSize { Val = "24" });
            if (lane == "defaults") defaults.AppendChild(new FontSizeComplexScript { Val = "30" });
            if (lane == "style") styleRun.AppendChild(new FontSizeComplexScript { Val = "30" });
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new DocDefaults(new RunPropertiesDefault(defaults)),
                new Style(new StyleName { Val = "Normal" }, styleRun)
                { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });
            var level = new Level(new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = literal ? NumberFormatValues.Bullet : NumberFormatValues.Decimal },
                new LevelText { Val = literal ? "12" : "%1" }) { LevelIndex = 0 };
            if (levelProperties != "absent")
            {
                var properties = new NumberingSymbolRunProperties();
                if (levelProperties == "fonts") properties.AppendChild(Fonts());
                level.AppendChild(properties);
            }
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
                new AbstractNum(level) { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
            var body = new Body();
            for (var i = 0; i < 3; i++)
            {
                var mark = new ParagraphMarkRunProperties(Fonts(), new FontSize { Val = "24" });
                if (lane == "mark") mark.AppendChild(new FontSizeComplexScript { Val = "30" });
                body.AppendChild(new Paragraph(
                    new ParagraphProperties(new NumberingProperties(new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = 1 }), new BiDi { Val = rtl }, mark),
                    new Run(new RunProperties(Fonts(), new FontSize { Val = "24" }), new Text($"Item{i}"))));
            }
            main.Document = new Document(body);
        }
        return stream.ToArray();
    }
}
