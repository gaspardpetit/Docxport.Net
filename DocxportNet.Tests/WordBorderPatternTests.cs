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
    [Fact]
    public void MixedWidthAndColorCellsKeepEachDoubleAndTripleBoxClosed()
    {
        var styles = new[] { "single", "double", "triple" };
        var sizes = new uint[] { 8, 16, 24 };
        var colors = new[] { "174D8A", "B33229", "26734D" };
        var content = ExportRawBorderStyles((row, column) =>
        {
            var index = (row + column) % 3;
            return (styles[index], sizes[index], colors[index]);
        });
        var widthCommands = Regex.Matches(content, @"(?m)(?<width>[\d.]+) w").Cast<Match>().ToArray();
        var boxes = Regex.Matches(content,
                @"(?m)(?<x>-?[\d.]+) (?<y>-?[\d.]+) (?<w>[\d.]+) (?<h>[\d.]+) re\s+S$")
            .Cast<Match>()
            .Select(match => (
                X: double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                Y: double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture),
                W: double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture),
                H: double.Parse(match.Groups["h"].Value, CultureInfo.InvariantCulture),
                Stroke: double.Parse(widthCommands.Last(command => command.Index < match.Index)
                    .Groups["width"].Value, CultureInfo.InvariantCulture)))
            .ToArray();
        Assert.Equal(15, boxes.Length);
        Assert.Equal(6, boxes.Count(box => Math.Abs(box.Stroke - 2) < .02));
        Assert.Equal(9, boxes.Count(box => Math.Abs(box.Stroke - 3) < .02));
        var cells = boxes.GroupBy(box => (
            X: Math.Round(box.X + box.W / 2, 2), Y: Math.Round(box.Y + box.H / 2, 2))).ToArray();
        Assert.Equal(6, cells.Length);
        Assert.Equal(3, cells.Count(group => group.Count() == 2));
        Assert.Equal(3, cells.Count(group => group.Count() == 3));
        foreach (var group in cells)
        {
            var rails = group.OrderBy(box => box.W).ToArray();
            for (var index = 1; index < rails.Length; index++)
            {
                Assert.True(rails[index].X < rails[index - 1].X);
                Assert.True(rails[index].Y < rails[index - 1].Y);
                Assert.True(rails[index].X + rails[index].W > rails[index - 1].X + rails[index - 1].W);
                Assert.True(rails[index].Y + rails[index].H > rails[index - 1].Y + rails[index - 1].H);
                Assert.Equal(2 * rails[index].Stroke,
                    rails[index - 1].X - rails[index].X, .02);
            }
        }
    }

    [Theory]
    [InlineData("threeDEmboss", true)]
    [InlineData("threeDEngrave", false)]
    public void ThreeDimensionalBordersReverseTheirThreeTouchingShades(string style, bool emboss)
    {
        var content = ExportRawBorderStyle(style);
        var widths = Regex.Matches(content, @"(?m)(?<w>[\d.]+) w").Cast<Match>().ToArray();
        var colors = Regex.Matches(content,
            @"(?m)(?<r>[\d.]+) (?<g>[\d.]+) (?<b>[\d.]+) RG").Cast<Match>().ToArray();
        var vertical = Regex.Matches(content,
                @"(?m)(?<x1>-?[\d.]+) (?<y1>-?[\d.]+) m\s+(?<x2>-?[\d.]+) (?<y2>-?[\d.]+) l\s+S$")
            .Cast<Match>()
            .Where(match => Math.Abs(double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture) -
                double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture)) < .01 &&
                Math.Abs(double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture) -
                    double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)) > 10)
            .Select(match => (
                X: double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Width: double.Parse(widths.Last(command => command.Index < match.Index)
                    .Groups["w"].Value, CultureInfo.InvariantCulture),
                Red: double.Parse(colors.Last(command => command.Index < match.Index)
                    .Groups["r"].Value, CultureInfo.InvariantCulture)))
            .ToArray();
        Assert.Equal(12, vertical.Length);
        Assert.Equal(12, vertical.Select(line => Math.Round(line.X, 2)).Distinct().Count());
        var firstBoundary = vertical.OrderBy(line => line.X).Take(3)
            .OrderBy(line => line.X).ToArray();
        Assert.Equal(3, firstBoundary.Length);
        Assert.Equal(.75, firstBoundary[0].Width, .02);
        Assert.Equal(2, firstBoundary[1].Width, .02);
        Assert.Equal(.75, firstBoundary[2].Width, .02);
        Assert.Equal(emboss, firstBoundary[0].Red > firstBoundary[2].Red);
        Assert.InRange(Math.Abs(firstBoundary[0].Red - firstBoundary[2].Red), .2, .3);
    }

    [Fact]
    public void DecorativeBordersUseRepeatedAngledStrokesOnTheWholeGrid()
    {
        static (double X1, double Y1, double X2, double Y2)[] Segments(string content) =>
            Regex.Matches(content,
                    @"(?m)(?<x1>-?[\d.]+) (?<y1>-?[\d.]+) m\s+(?<x2>-?[\d.]+) (?<y2>-?[\d.]+) l\s+S$")
                .Cast<Match>()
                .Select(match => (
                    X1: double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                    Y1: double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                    X2: double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                    Y2: double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)))
                .ToArray();

        var wave = Segments(ExportRawBorderStyle("wave"));
        var doubleWave = Segments(ExportRawBorderStyle("doubleWave"));
        var hatch = Segments(ExportRawBorderStyle("dashDotStroked"));
        Assert.True(wave.Length > 500);
        Assert.InRange((double)doubleWave.Length / wave.Length, 1.9, 2.1);
        Assert.True(hatch.Length > 500);
        foreach (var segments in new[] { wave, doubleWave, hatch })
        {
            Assert.All(segments, line =>
            {
                Assert.InRange(Math.Abs(line.X2 - line.X1), .01, 4);
                Assert.InRange(Math.Abs(line.Y2 - line.Y1), .01, 4);
            });
            Assert.True(segments.Count(line => Math.Abs(line.X2 - line.X1) > .3 &&
                Math.Abs(line.Y2 - line.Y1) > .3) > segments.Length * .9);
            Assert.True(segments.Max(line => Math.Max(line.X1, line.X2)) -
                segments.Min(line => Math.Min(line.X1, line.X2)) > 200);
            Assert.True(segments.Max(line => Math.Max(line.Y1, line.Y2)) -
                segments.Min(line => Math.Min(line.Y1, line.Y2)) > 30);
        }
    }

    [Theory]
    [InlineData("thinThickSmallGap", 2, .75, .75)]
    [InlineData("thickThinSmallGap", .75, 2, .75)]
    [InlineData("thinThickMediumGap", 2, 1, 1)]
    [InlineData("thickThinMediumGap", 1, 2, 1)]
    [InlineData("thinThickLargeGap", 1.5, .75, 2)]
    [InlineData("thickThinLargeGap", .75, 1.5, 2)]
    public void TwoRailWordStylesTurnAtEveryCellCorner(
        string style, double firstWidth, double secondWidth, double gap)
    {
        var content = ExportRawBorderStyle(style);
        var widthCommands = Regex.Matches(content, @"(?m)(?<width>[\d.]+) w").Cast<Match>().ToArray();
        var strokes = Regex.Matches(content,
                @"(?m)(?<x1>-?[\d.]+) (?<y1>-?[\d.]+) m\s+(?<x2>-?[\d.]+) (?<y2>-?[\d.]+) l\s+S$")
            .Cast<Match>()
            .Select(match => (
                Width: double.Parse(widthCommands.Last(command => command.Index < match.Index)
                    .Groups["width"].Value, CultureInfo.InvariantCulture),
                X1: double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                Y1: double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                X2: double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                Y2: double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)))
            .ToArray();
        var vertical = strokes.Where(line => Math.Abs(line.X1 - line.X2) < .01 &&
            Math.Abs(line.Y1 - line.Y2) > 10).OrderBy(line => line.X1).ToArray();
        var horizontal = strokes.Where(line => Math.Abs(line.Y1 - line.Y2) < .01 &&
            Math.Abs(line.X1 - line.X2) > 10).OrderByDescending(line => line.Y1).ToArray();
        Assert.Equal(36, vertical.Length);
        Assert.Equal(36, horizontal.Length);
        Assert.Equal(18, vertical.Count(line => Math.Abs(line.Width - firstWidth) < .02));
        Assert.Equal(18, vertical.Count(line => Math.Abs(line.Width - secondWidth) < .02));
        Assert.Equal(18, horizontal.Count(line => Math.Abs(line.Width - firstWidth) < .02));
        Assert.Equal(18, horizontal.Count(line => Math.Abs(line.Width - secondWidth) < .02));
        Assert.All(vertical, line => Assert.InRange(Math.Abs(line.Y2 - line.Y1), 10, 30));
        Assert.All(horizontal, line => Assert.InRange(Math.Abs(line.X2 - line.X1), 90, 110));
        // Every stroke terminates at a perpendicular rail. Long strokes across
        // the whole table would pass the width checks but fail these endpoints.
        Assert.All(vertical, line =>
        {
            foreach (var end in new[] { line.Y1, line.Y2 })
                Assert.Contains(horizontal, cross =>
                    Math.Abs(cross.Y1 - end) <= (line.Width + cross.Width) / 2 + .02 &&
                    Math.Min(cross.X1, cross.X2) - cross.Width / 2 <= line.X1 &&
                    Math.Max(cross.X1, cross.X2) + cross.Width / 2 >= line.X1);
        });
        Assert.All(horizontal, line =>
        {
            foreach (var end in new[] { line.X1, line.X2 })
                Assert.Contains(vertical, cross =>
                    Math.Abs(cross.X1 - end) <= (line.Width + cross.Width) / 2 + .02 &&
                    Math.Min(cross.Y1, cross.Y2) - cross.Width / 2 <= line.Y1 &&
                    Math.Max(cross.Y1, cross.Y2) + cross.Width / 2 >= line.Y1);
        });
    }

    [Theory]
    [InlineData("thinThickThinSmallGap", .75, 2, .75)]
    [InlineData("thinThickThinMediumGap", 1, 2, 1)]
    [InlineData("thinThickThinLargeGap", .75, 1.5, 2)]
    public void ThreeRailWordStylesKeepClosedBoxesAndMeasuredGaps(
        string style, double thinWidth, double middleWidth, double gap)
    {
        var content = ExportRawBorderStyle(style);
        var widthCommands = Regex.Matches(content, @"(?m)(?<width>[\d.]+) w").Cast<Match>().ToArray();
        var boxes = Regex.Matches(content,
                @"(?m)(?<x>-?[\d.]+) (?<y>-?[\d.]+) (?<w>[\d.]+) (?<h>[\d.]+) re\s+S$")
            .Cast<Match>()
            .Select(match => (
                Stroke: double.Parse(widthCommands.Last(command => command.Index < match.Index)
                    .Groups["width"].Value, CultureInfo.InvariantCulture),
                X: double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                Y: double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture),
                W: double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture),
                H: double.Parse(match.Groups["h"].Value, CultureInfo.InvariantCulture)))
            .ToArray();
        Assert.Equal(19, boxes.Length);
        var outer = boxes.OrderByDescending(box => box.W * box.H).First();
        var inner = boxes.Where(box => box != outer).OrderBy(box => box.W * box.H).Take(9).ToArray();
        var middle = boxes.Where(box => box != outer).Except(inner).ToArray();
        Assert.Equal(9, middle.Length);
        Assert.Equal(thinWidth, outer.Stroke, .02);
        Assert.All(inner, box => Assert.Equal(thinWidth, box.Stroke, .02));
        Assert.All(middle, box => Assert.Equal(middleWidth, box.Stroke, .02));
        var expectedOffset = (thinWidth + middleWidth) / 2 + gap;
        foreach (var box in inner)
        {
            var enclosing = middle.Where(mid => mid.X < box.X && mid.Y < box.Y &&
                mid.X + mid.W > box.X + box.W && mid.Y + mid.H > box.Y + box.H).ToArray();
            Assert.Single(enclosing);
            Assert.Equal(expectedOffset, box.X - enclosing[0].X, .02);
            Assert.Equal(expectedOffset, box.Y - enclosing[0].Y, .02);
            Assert.True(outer.X < enclosing[0].X && outer.Y < enclosing[0].Y);
            Assert.True(outer.X + outer.W > enclosing[0].X + enclosing[0].W &&
                outer.Y + outer.H > enclosing[0].Y + enclosing[0].H);
        }
        Assert.Equal(3, inner.Select(box => Math.Round(box.X, 2)).Distinct().Count());
        Assert.Equal(3, inner.Select(box => Math.Round(box.Y, 2)).Distinct().Count());
    }

    private static string ExportRawBorderStyle(string style) =>
        ExportRawBorderStyles((_, _) => (style, 16, "1B2530"));

    private static string ExportRawBorderStyles(Func<int, int, (string Style, uint Size, string Color)> borderAt)
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
                {
                    var (style, size, color) = borderAt(row, column);
                    T Edge<T>(T border) where T : BorderType
                    {
                        border.Size = size;
                        border.Color = color;
                        border.SetAttribute(new OpenXmlAttribute("w", "val",
                            "http://schemas.openxmlformats.org/wordprocessingml/2006/main", style));
                        return border;
                    }
                    var borders = new TableCellBorders(Edge(new TopBorder()), Edge(new RightBorder()),
                        Edge(new BottomBorder()), Edge(new LeftBorder()));
                    cells.Append(new TableCell(new TableCellProperties(
                            new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2000" }, borders),
                        new Paragraph(new Run(new Text($"{row},{column}")))));
                }
                table.Append(cells);
            }
            main.Document = new Document(new Body(table));
            main.Document.Save();
        }
        var bytes = DxpExport.ExportToBytes(source.ToArray(), new DxpFlowPdfVisitor());
        using var pdf = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        Assert.Single(pdf.Pages);
        return string.Concat(pdf.Pages[0].Contents.Elements
            .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
            .Select(reference => Encoding.Latin1.GetString(
                ((PdfSharp.Pdf.Advanced.PdfContent)reference.Value!).Stream!.UnfilteredValue)));
    }

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
