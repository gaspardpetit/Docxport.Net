using System.Globalization;
using System.Text;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;

namespace DocxportNet.Visitors.Pdf;

internal static class DxpPdfMarkupPrintLayout
{
    internal sealed record Balloon(double Top, double ConnectorX, string Description);

    internal static void Apply(PdfDocument document, IReadOnlyList<IReadOnlyList<Balloon>> pages)
    {
        var scales = new (double Scale, double Offset)[document.PageCount];
        for (var pageIndex = 0; pageIndex < document.PageCount; pageIndex++)
        {
            var page = document.Pages[pageIndex];
            var width = page.Width.Point;
            var height = page.Height.Point;
            var scale = width / (width + 194.4);
            var offset = (1 - scale) * height / 2;
            scales[pageIndex] = (scale, offset);
            var content = page.Contents.CreateSingleContent().Stream!.UnfilteredValue;
            var prefix = Encoding.ASCII.GetBytes(string.Create(CultureInfo.InvariantCulture,
                $"q\n{scale:R} 0 0 {scale:R} 0 {offset:R} cm\n"));
            var suffix = Encoding.ASCII.GetBytes("\nQ\n");
            page.Contents.ReplaceContent([.. prefix, .. content, .. suffix]);
            foreach (var annotation in page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>())
            {
                var rectangle = annotation.Rectangle;
                annotation.Rectangle = new PdfRectangle(
                    new XPoint(rectangle.X1 * scale, rectangle.Y1 * scale + offset),
                    new XPoint(rectangle.X2 * scale, rectangle.Y2 * scale + offset));
                if (annotation.Elements.GetArray("/QuadPoints") is not { } quads) continue;
                for (var index = 0; index < quads.Elements.Count; index++)
                    if (Number(quads.Elements[index]) is { } coordinate)
                        quads.Elements[index] = new PdfReal(index % 2 == 0
                            ? coordinate * scale : coordinate * scale + offset);
            }
            using (var background = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Prepend))
                background.DrawRectangle(new XSolidBrush(XColor.FromArgb(242, 242, 242)),
                    width - 194.5, offset, 194.5, height - 2 * offset);
            if (pageIndex >= pages.Count || pages[pageIndex].Count == 0) continue;
            using var graphics = XGraphics.FromPdfPage(page);
            var labelFont = new XFont("Arial", 6.8, XFontStyleEx.Bold);
            var bodyFont = new XFont("Arial", 6.8);
            var pen = new XPen(XColor.FromArgb(82, 163, 250), 0.5);
            foreach (var balloon in pages[pageIndex])
            {
                var left = width - 194.5 + 17.1;
                var top = offset + balloon.Top * scale - 0.15;
                graphics.DrawLine(pen, Math.Min(left, balloon.ConnectorX * scale + 2.5),
                    top + 6.1, left, top + 6.1);
                graphics.DrawRoundedRectangle(pen, XBrushes.White,
                    new XRect(left, top, width - 3.9 - left, 12.3), new XSize(2, 2));
                graphics.DrawString("Formatted:", labelFont, XBrushes.Black,
                    new XPoint(left + 3.5, top + 8.7));
                graphics.DrawString(" " + balloon.Description, bodyFont, XBrushes.Black,
                    new XPoint(left + 3.5 + graphics.MeasureString("Formatted:", labelFont).Width,
                        top + 8.7));
            }
        }

        var names = document.Internals.Catalog.Names?.NameTree;
        if (names == null) return;
        foreach (var name in names.GetNames(true))
        {
            if (names.GetValue(name, true) is not PdfArray destination ||
                destination.Elements.Count < 2 ||
                destination.Elements[0] is not PdfReference pageReference) continue;
            var pageIndex = Enumerable.Range(0, document.PageCount).FirstOrDefault(index =>
                document.Pages[index].Reference?.ObjectID == pageReference.ObjectID, -1);
            if (pageIndex < 0) continue;
            var (scale, offset) = scales[pageIndex];
            if (destination.Elements.Count >= 4 &&
                destination.Elements[1].ToString() == "/XYZ")
            {
                if (Number(destination.Elements[2]) is { } left)
                    destination.Elements[2] = new PdfReal(left * scale);
                if (Number(destination.Elements[3]) is { } top)
                    destination.Elements[3] = new PdfReal(top * scale + offset);
            }
            else if (destination.Elements[1] is PdfLiteral literal)
            {
                var parts = literal.ToString().Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 4 || parts[0] != "/XYZ" ||
                    !double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture,
                        out var left) ||
                    !double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture,
                        out var top)) continue;
                parts[1] = (left * scale).ToString("R", CultureInfo.InvariantCulture);
                parts[2] = (top * scale + offset).ToString("R", CultureInfo.InvariantCulture);
                destination.Elements[1] = new PdfLiteral(string.Join(' ', parts));
            }
        }
    }

    private static double? Number(PdfItem item) => item switch
    {
        PdfReal real => real.Value,
        PdfInteger integer => integer.Value,
        _ => null
    };
}
