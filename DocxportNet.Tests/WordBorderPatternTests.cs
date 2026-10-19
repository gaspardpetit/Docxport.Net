using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Visitors.Pdf;
using PdfSharp.Fonts;
using PdfSharp.Pdf.IO;

namespace DocxportNet.Tests;

public class WordBorderPatternTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitCompoundCellBordersProduceClosedBoxesWithAlignedRailGaps(bool triple)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var table = new Table(new TableProperties(
                    new TableWidth { Type = TableWidthUnitValues.Dxa, Width = "6000" }),
                new TableGrid(Enumerable.Range(0, 3).Select(_ => new GridColumn { Width = "2000" })));
            for (var row = 0; row < 3; row++)
            {
                var cells = new TableRow();
                for (var column = 0; column < 3; column++)
                    cells.Append(new TableCell(
                        new TableCellProperties(
                            new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2000" },
                            new TableCellBorders(
                                new TopBorder { Val = triple ? BorderValues.Triple : BorderValues.Double, Size = 16 },
                                new RightBorder { Val = triple ? BorderValues.Triple : BorderValues.Double, Size = 16 },
                                new BottomBorder { Val = triple ? BorderValues.Triple : BorderValues.Double, Size = 16 },
                                new LeftBorder { Val = triple ? BorderValues.Triple : BorderValues.Double, Size = 16 })),
                        new Paragraph(new Run(new Text($"{row},{column}")))));
                table.Append(cells);
            }
            main.Document = new Document(new Body(table));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Single(pdf.Pages);
        var content = string.Concat(pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => Encoding.Latin1.GetString(
                ((PdfSharp.Pdf.Advanced.PdfContent)reference.Value!).Stream!.UnfilteredValue)));
        var boxes = Regex.Matches(content,
                @"(?m)(?<x>-?[\d.]+) (?<y>-?[\d.]+) (?<w>[\d.]+) (?<h>[\d.]+) re\s+S$")
            .Cast<Match>()
            .Select(match => (
                X: double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                Y: double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture),
                W: double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture),
                H: double.Parse(match.Groups["h"].Value, CultureInfo.InvariantCulture)))
            .ToArray();
        Assert.Equal(triple ? 19 : 10, boxes.Length);
        var outer = boxes.OrderByDescending(box => box.W * box.H).First();
        var cellBoxes = boxes.Where(box => box != outer).ToArray();
        var grid = triple
            ? cellBoxes.OrderBy(box => box.W * box.H).Take(9).OrderBy(box => box.Y).ThenBy(box => box.X).ToArray()
            : cellBoxes.OrderBy(box => box.Y).ThenBy(box => box.X).ToArray();
        Assert.Equal(9, grid.Length);
        Assert.Equal(3, grid.Select(box => Math.Round(box.X, 2)).Distinct().Count());
        Assert.Equal(3, grid.Select(box => Math.Round(box.Y, 2)).Distinct().Count());
        var rail = grid[0].X - outer.X;
        Assert.InRange(rail, 1, 10);
        Assert.Equal(rail, grid[0].Y - outer.Y, 0.02);
        Assert.Equal(rail, outer.X + outer.W - grid[2].X - grid[2].W, 0.02);
        Assert.Equal(rail, outer.Y + outer.H - grid[6].Y - grid[6].H, 0.02);
        if (triple)
        {
            var middle = cellBoxes.Except(grid).ToArray();
            Assert.Equal(9, middle.Length);
            foreach (var inner in grid)
            {
                var enclosing = middle.Where(box => box.X < inner.X && box.Y < inner.Y &&
                    box.X + box.W > inner.X + inner.W && box.Y + box.H > inner.Y + inner.H).ToArray();
                Assert.Single(enclosing);
                Assert.Equal((rail / 2), inner.X - enclosing[0].X, 0.02);
                Assert.Equal((rail / 2), inner.Y - enclosing[0].Y, 0.02);
            }
        }
        for (var row = 0; row < 3; row++)
        {
            var cells = grid.Skip(row * 3).Take(3).ToArray();
            Assert.All(cells, cell => Assert.Equal(cells[0].Y, cell.Y, 0.02));
            Assert.All(cells, cell => Assert.Equal(cells[0].Y + cells[0].H, cell.Y + cell.H, 0.02));
            for (var column = 0; column < 2; column++)
                Assert.Equal(rail, cells[column + 1].X - cells[column].X - cells[column].W, 0.02);
        }
        for (var row = 0; row < 2; row++)
            Assert.Equal(rail, grid[(row + 1) * 3].Y - grid[row * 3].Y - grid[row * 3].H, 0.02);
    }

    [Theory]
    [InlineData(false, "[6 2 2 2]0 d")]
    [InlineData(true, "[6 2 2 2 2 2]0 d")]
    public void ThreeByThreePatternedGridKeepsWordCadenceAndOneSharedEdge(
        bool twoDots, string expectedDash)
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        var style = twoDots ? BorderValues.DotDotDash : BorderValues.DotDash;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var table = new Table(new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Dxa, Width = "6000" }),
                new TableGrid(Enumerable.Range(0, 3).Select(_ => new GridColumn { Width = "2000" })));
            for (var row = 0; row < 3; row++)
            {
                var cells = new TableRow();
                for (var column = 0; column < 3; column++)
                {
                    var borders = new TableCellBorders(
                        new TopBorder { Val = style, Size = 16 },
                        new RightBorder { Val = style, Size = 16 },
                        new BottomBorder { Val = style, Size = 16 },
                        new LeftBorder { Val = style, Size = 16 });
                    cells.Append(new TableCell(
                        new TableCellProperties(
                            new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2000" }, borders),
                        new Paragraph(new Run(new Text($"{row + 1},{column + 1}")))));
                }
                table.Append(cells);
            }
            main.Document = new Document(new Body(table));
            main.Document.Save();
        }

        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Single(pdf.Pages);
        var content = string.Concat(pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => Encoding.Latin1.GetString(
                ((PdfSharp.Pdf.Advanced.PdfContent)reference.Value!).Stream!.UnfilteredValue)));
        Assert.Contains(expectedDash, content);

        var strokes = Regex.Matches(content,
                @"(?m)(?<x1>-?[\d.]+) (?<y1>-?[\d.]+) m\s+(?<x2>-?[\d.]+) (?<y2>-?[\d.]+) l\s+S")
            .Cast<Match>()
            .Select(match => (
                X1: double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1: double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2: double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2: double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)))
            .ToArray();
        var longVertical = strokes.Where(line => Math.Abs(line.X1 - line.X2) < .01 &&
            Math.Abs(line.Y1 - line.Y2) > 10).ToArray();
        var longHorizontal = strokes.Where(line => Math.Abs(line.Y1 - line.Y2) < .01 &&
            Math.Abs(line.X1 - line.X2) > 10).ToArray();
        Assert.Equal(12, longVertical.Length);
        Assert.Equal(12, longHorizontal.Length);
        Assert.Equal(4, longVertical.Select(line => Math.Round(line.X1, 2)).Distinct().Count());
        Assert.Equal(4, longHorizontal.Select(line => Math.Round(line.Y1, 2)).Distinct().Count());
        Assert.All(longVertical.GroupBy(line => Math.Round(line.X1, 2)), group => Assert.Equal(3, group.Count()));
        Assert.All(longHorizontal.GroupBy(line => Math.Round(line.Y1, 2)), group => Assert.Equal(3, group.Count()));
    }
}
