using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Visitors.Html;
using System.Xml.Linq;

namespace DocxportNet.Tests;

public sealed class HtmlTableFidelityTests
{
    [Fact]
    public void RichHtmlUsesExplicitRunSizeAndFixedTableGridWithoutDefaultCellSpacing()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(new Style(
                new StyleName { Val = "Normal" },
                new StyleRunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }))
                { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });
            styles.Styles.Save();
            main.Document = new Document(new Body(new Table(
                new TableProperties(new TableWidth { Width = "9468", Type = TableWidthUnitValues.Dxa },
                    new TableLayout { Type = TableLayoutValues.Fixed },
                    new TableCellMarginDefault(new TableCellRightMargin
                        { Width = 108, Type = TableWidthValues.Dxa })),
                new TableGrid(new GridColumn { Width = "1242" }, new GridColumn { Width = "8226" }),
                new TableRow(Cell("Label"), Cell("Value", 200)),
                new TableRow(Cell("Next"), Cell("Another")))));
            main.Document.Save();
        }

        var html = DxpExport.ExportToString(stream.ToArray(),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()));
        Assert.Contains(".dxp-table-rich .dxp-paragraph { margin: 0; line-height: normal; }", html);
        Assert.Contains("padding: 0 5.4pt;", html);
        Assert.Contains("width:473.4pt;table-layout:fixed;", html);
        Assert.Contains("<col style=\"width:62.1pt\">", html);
        Assert.Contains("<col style=\"width:411.3pt\">", html);
        Assert.Contains("<p class=\"dxp-paragraph\" style=\"font-size:10.5pt;\">", html);
        Assert.Contains("padding-right:5.4pt;", html);
        Assert.Contains("padding-left:10pt;", html);
        Assert.Contains("font-size:10.5pt;", html);
    }


    [Fact]
    public void RichHtmlKeepsEmptyParagraphsInsideCells()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(new Table(
                new TableRow(new TableCell(
                    new Paragraph(new Run(new Text("First"))),
                    new Paragraph(new ParagraphProperties(new SpacingBetweenLines { Line = "200", LineRule = LineSpacingRuleValues.Exact })),
                    new Paragraph(new Run(new Text("Second"))))))));
            main.Document.Save();
        }

        var html = DxpExport.ExportToString(stream.ToArray(),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()));
        var paragraphs = XDocument.Parse(html).Descendants("td").Single().Elements("p").ToArray();
        Assert.Equal(3, paragraphs.Length);
        Assert.Equal("First", paragraphs[0].Value);
        Assert.Equal("\u00a0", paragraphs[1].Value);
        Assert.Contains("line-height:10pt;", (string?)paragraphs[1].Attribute("style"));
        Assert.Equal("Second", paragraphs[2].Value);
        Assert.Contains("Second</p>", html);
    }

    [Fact]
    public void InsideHorizontalBorderDoesNotBecomeVerticalOrOuterBorder()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(new Table(
                new TableProperties(new TableBorders(new InsideHorizontalBorder
                    { Val = BorderValues.Dashed, Size = 4 })),
                new TableRow(Cell("A"), Cell("B")),
                new TableRow(Cell("C"), Cell("D")))));
            main.Document.Save();
        }

        var html = DxpExport.ExportToString(stream.ToArray(),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()));
        var table = XDocument.Parse(html).Descendants("table").Single();
        Assert.DoesNotContain("border:", (string?)table.Attribute("style") ?? "");
        var rows = table.Elements("tr").ToArray();
        Assert.Equal(2, rows.Length);
        foreach (var cell in rows[0].Elements("td"))
        {
            var css = (string?)cell.Attribute("style") ?? "";
            Assert.Contains("border-bottom:0.5pt dashed", css);
            Assert.DoesNotContain("border-left:", css);
            Assert.DoesNotContain("border-right:", css);
            Assert.DoesNotContain("border-top:", css);
        }
        foreach (var cell in rows[1].Elements("td"))
            Assert.DoesNotContain("border:", (string?)cell.Attribute("style") ?? "");
    }

    private static TableCell Cell(string text, short? leftTwips = null)
    {
        var cell = new TableCell(new Paragraph(
            new ParagraphProperties(new ParagraphMarkRunProperties(new FontSize { Val = "21" })),
            new Run(new RunProperties(new RunFonts { Ascii = "Arial" },
                new FontSize { Val = "21" }), new Text(text))));
        if (leftTwips is short left)
            cell.PrependChild(new TableCellProperties(new TableCellMargin(
                new TableCellLeftMargin { Width = left, Type = TableWidthValues.Dxa })));
        return cell;
    }
}
