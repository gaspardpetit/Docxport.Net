using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker.Context;

namespace DocxportNet.Tests;

public class TableLookMaskTests
{
    [Fact]
    public void BitmaskEnablesWordAuthoredFirstRowFormattingWithoutBooleanAttributes()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalFirstRowRunStyleAllStories.docx"));
        using var document = WordprocessingDocument.Open(path, false);
        var table = document.MainDocumentPart!.Document!.Body!.Descendants<Table>().First();
        table.TableProperties!.TableLook = new TableLook { Val = "0020" };
        var resolver = new DxpTableStyleResolver(document);

        var enabled = resolver.ResolveTableStyle(table.TableProperties, 2, 2);
        var firstCell = DxpTableStyleResolver.ComputeCellStyle(enabled, null, 0, 0, 1, 1);
        Assert.True(enabled.Look.FirstRow);
        Assert.True(firstCell.ConditionalBold);
        Assert.Equal("#FF0000", firstCell.ConditionalTextColorCss);

        table.TableProperties.TableLook.FirstRow = false;
        var disabled = resolver.ResolveTableStyle(table.TableProperties, 2, 2);
        Assert.False(disabled.Look.FirstRow);
        Assert.Null(DxpTableStyleResolver.ComputeCellStyle(disabled, null, 0, 0, 1, 1).ConditionalBold);
    }

    [Fact]
    public void BitmaskSuppliesEachOmittedTableLookFlag()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordConditionalFirstRowRunStyleAllStories.docx"));
        using var document = WordprocessingDocument.Open(path, false);
        var resolver = new DxpTableStyleResolver(document);
        var properties = new TableProperties(new TableLook { Val = "07E0" });

        var look = resolver.ResolveTableStyle(properties, 2, 2).Look;

        Assert.True(look.FirstRow);
        Assert.True(look.LastRow);
        Assert.True(look.FirstColumn);
        Assert.True(look.LastColumn);
        Assert.True(look.NoHBand);
        Assert.True(look.NoVBand);
    }
}
