using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Doc;

namespace DocxportNet.Tests;

public sealed class DocTableBorderProjectionTests
{
    [Fact]
    public void SharedTableBordersStayOnTableWithoutDuplicateCellEdges()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "fixtures", "doc", "legacy",
            "WordLegacyTableBordersAllStories.doc"));
        var projected = DxpDocToDocx.Project(path).DocxBytes;
        using var document = WordprocessingDocument.Open(new MemoryStream(projected), false);
        var table = Assert.Single(document.MainDocumentPart!.Document!.Body!.Elements<Table>());
        var borders = table.TableProperties!.GetFirstChild<TableBorders>();
        Assert.NotNull(borders?.InsideHorizontalBorder);
        Assert.NotNull(borders.InsideVerticalBorder);
        var margins = table.TableProperties.GetFirstChild<TableCellMarginDefault>();
        Assert.Null(margins?.GetFirstChild<TableCellLeftMargin>());
        Assert.Equal((short)108, margins?.GetFirstChild<TableCellRightMargin>()?.Width?.Value);
        Assert.All(table.Descendants<TableCell>(), cell =>
        {
            var direct = cell.TableCellProperties?.TableCellBorders;
            Assert.Null(direct?.TopBorder);
            Assert.Null(direct?.LeftBorder);
            Assert.Null(direct?.BottomBorder);
            Assert.Null(direct?.RightBorder);
        });
    }
}
