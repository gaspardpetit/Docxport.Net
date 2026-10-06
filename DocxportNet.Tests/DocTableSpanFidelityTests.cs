using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Doc;

namespace DocxportNet.Tests;

public sealed class DocTableSpanFidelityTests
{
    [Fact]
    public void TableSpanExportsAndIsReadByDocProjector()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "samples", "TestTableSpan.docx"));
        var source = File.ReadAllBytes(path);
        var binary = DxpDocExport.Export(source);

        using var index = new DocTextIndexWalker().Index(new MemoryStream(binary));
        Assert.Contains(index.ParagraphStyles, range =>
            range.Formatting?.TableTerminator == true &&
            range.Formatting.TableCellEdges is { Count: 4 });

        var projected = DxpDocToDocx.Project(binary).DocxBytes;
        using var sourceStream = new MemoryStream(source);
        using var sourceDoc = WordprocessingDocument.Open(sourceStream, false);
        using var outputStream = new MemoryStream(projected);
        using var outputDoc = WordprocessingDocument.Open(outputStream, false);
        Assert.Equal(sourceDoc.MainDocumentPart!.Document.Body!.Descendants<Table>().Count(),
            outputDoc.MainDocumentPart!.Document.Body!.Descendants<Table>().Count());
        Assert.Equal(sourceDoc.MainDocumentPart.Document.Body.InnerText,
            outputDoc.MainDocumentPart.Document.Body.InnerText);
        static string CellShape(TableCell cell)
        {
            var merge = cell.TableCellProperties?.GetFirstChild<VerticalMerge>();
            var mergeKind = merge == null ? "none" :
                merge.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";
            return $"{cell.InnerText}|{cell.TableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1}|{mergeKind}";
        }
        static string[][][] TableShape(WordprocessingDocument document) =>
            document.MainDocumentPart!.Document.Body!.Elements<Table>()
                .Select(table => table.Elements<TableRow>()
                    .Select(row => row.Elements<TableCell>()
                        .Select(CellShape).ToArray()).ToArray()).ToArray();
        Assert.Equal(TableShape(sourceDoc), TableShape(outputDoc));
    }
}
