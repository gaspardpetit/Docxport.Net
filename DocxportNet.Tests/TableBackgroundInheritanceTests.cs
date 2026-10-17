using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker.Context;

namespace DocxportNet.Tests;

public class TableBackgroundInheritanceTests
{
    [Fact]
    public void TableShadingFlowsIntoCellsAndDirectClearRemovesIt()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body());
        var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(
            new Style(new StyleTableProperties(new Shading { Fill = "FF0000" })) {
                StyleId = "Base", Type = StyleValues.Table
            },
            new Style(new BasedOn { Val = "Base" },
                new StyleTableCellProperties(new Shading { Fill = "00FF00" })) {
                StyleId = "Derived", Type = StyleValues.Table
            });
        var resolver = new DxpTableStyleResolver(document);

        var baseStyle = resolver.ResolveTableStyle(new TableProperties(new TableStyle { Val = "Base" }), 1, 1);
        var derivedStyle = resolver.ResolveTableStyle(new TableProperties(new TableStyle { Val = "Derived" }), 1, 1);
        var directStyle = resolver.ResolveTableStyle(new TableProperties(
            new TableStyle { Val = "Derived" }, new Shading { Fill = "0000FF" }), 1, 1);

        Assert.Equal("#FF0000", DxpTableStyleResolver.ComputeCellStyle(baseStyle, null, 0, 0, 1, 1).BackgroundColorCss);
        Assert.Equal("#00FF00", DxpTableStyleResolver.ComputeCellStyle(derivedStyle, null, 0, 0, 1, 1).BackgroundColorCss);
        Assert.Equal("#0000FF", DxpTableStyleResolver.ComputeCellStyle(directStyle, null, 0, 0, 1, 1).BackgroundColorCss);

        var cleared = DxpTableStyleResolver.ComputeCellStyle(directStyle,
            new TableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Fill = "auto" }), 0, 0, 1, 1);
        Assert.Null(cleared.BackgroundColorCss);
    }
}
