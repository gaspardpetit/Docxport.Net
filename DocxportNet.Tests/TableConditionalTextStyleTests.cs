using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Walker.Context;

namespace DocxportNet.Tests;

public class TableConditionalTextStyleTests
{
    [Theory]
    [InlineData("WordConditionalFirstRowRunStyleAllStories.docx", false)]
    [InlineData("WordConditionalFirstRowParagraphAllStories.docx", true)]
    public void FirstRowConditionalTextFormattingIsAvailableToExporters(string fixture, bool centered)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc", fixture));
        using var document = WordprocessingDocument.Open(path, false);
        if (!centered)
        {
            var tableStyle = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .Elements<Style>().First(style => style.StyleId == "LogicalStart");
            var firstRow = tableStyle.ChildElements.First(element =>
                element.LocalName == "tblStylePr" &&
                element.GetAttributes().Any(attribute => attribute.LocalName == "type" &&
                    attribute.Value == "firstRow"));
            var runProperties = firstRow.ChildElements.First(element => element.LocalName == "rPr");
            runProperties.Append(new Italic { Val = false },
                new Underline { Val = UnderlineValues.None });
        }
        var table = document.MainDocumentPart!.Document!.Body!.Descendants<Table>().First();
        var rows = table.Elements<TableRow>().ToArray();
        var resolver = new DxpTableStyleResolver(document);
        var resolved = resolver.ResolveTableStyle(table.TableProperties, rows.Length,
            rows[0].Elements<TableCell>().Count());
        var firstCell = rows[0].Elements<TableCell>().First();
        var secondCell = rows[1].Elements<TableCell>().First();

        var firstStyle = DxpTableStyleResolver.ComputeCellStyle(resolved,
            firstCell.TableCellProperties, 0, 0, 1, 1);
        var secondStyle = DxpTableStyleResolver.ComputeCellStyle(resolved,
            secondCell.TableCellProperties, 1, 0, 1, 1);

        Assert.Equal("#FF0000", firstStyle.ConditionalTextColorCss);
        Assert.True(firstStyle.ConditionalBold);
        Assert.Equal(centered ? null : false, firstStyle.ConditionalItalic);
        Assert.Equal(centered ? null : false, firstStyle.ConditionalUnderline);
        Assert.Equal(centered ? DxpComputedTextAlign.Center : null,
            firstStyle.ConditionalTextAlign);
        Assert.Null(secondStyle.ConditionalTextColorCss);
        Assert.Null(secondStyle.ConditionalBold);
        Assert.Null(secondStyle.ConditionalTextAlign);
    }
}
