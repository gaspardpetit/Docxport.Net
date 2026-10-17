using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker.Context;

namespace DocxportNet.Tests;

public class TableLogicalBorderTests
{
    [Theory]
    [InlineData(false, false, "#FF0000", "#0000FF")]
    [InlineData(true, false, "#0000FF", "#FF0000")]
    [InlineData(true, true, "#0000FF", "#FF0000")]
    public void LogicalTableAndCellBordersFollowTableDirection(bool rightToLeft,
        bool styledDirection,
        string expectedLeft, string expectedRight)
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body());
        var tableProperties = new TableProperties(new TableBorders(
            new StartBorder { Val = BorderValues.Single, Size = 8, Color = "FF0000" },
            new EndBorder { Val = BorderValues.Single, Size = 8, Color = "0000FF" }));
        if (styledDirection)
        {
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new Style(new StyleTableProperties(new BiDiVisual())) {
                StyleId = "RtlTable", Type = StyleValues.Table
            });
            tableProperties.TableStyle = new TableStyle { Val = "RtlTable" };
        }
        else if (rightToLeft)
            tableProperties.Append(new BiDiVisual());

        var resolver = new DxpTableStyleResolver(document);
        var resolved = resolver.ResolveTableStyle(tableProperties, 1, 2);

        Assert.Equal(rightToLeft, resolved.RightToLeft);
        Assert.Equal(expectedLeft, resolved.ComputedTableStyle.TableBorders?.Left?.ColorCss);
        Assert.Equal(expectedRight, resolved.ComputedTableStyle.TableBorders?.Right?.ColorCss);

        var first = DxpTableStyleResolver.ComputeCellStyle(resolved, null, 0, 0, 1, 1);
        var last = DxpTableStyleResolver.ComputeCellStyle(resolved, null, 0, 1, 1, 1);
        Assert.Equal(expectedLeft, rightToLeft ? last.Borders?.Left?.ColorCss : first.Borders?.Left?.ColorCss);
        Assert.Equal(expectedRight, rightToLeft ? first.Borders?.Right?.ColorCss : last.Borders?.Right?.ColorCss);

        var direct = new TableCellProperties(new TableCellBorders(
            new StartBorder { Val = BorderValues.Single, Size = 8, Color = "00FF00" },
            new EndBorder { Val = BorderValues.Single, Size = 8, Color = "FFFF00" }));
        var directStyle = DxpTableStyleResolver.ComputeCellStyle(resolved, direct, 0, 0, 1, 1);
        Assert.Equal(rightToLeft ? "#FFFF00" : "#00FF00", directStyle.Borders?.Left?.ColorCss);
        Assert.Equal(rightToLeft ? "#00FF00" : "#FFFF00", directStyle.Borders?.Right?.ColorCss);
    }
}
