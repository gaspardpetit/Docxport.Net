using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using System.Text.RegularExpressions;

namespace DocxportNet.Visitors.Pdf;

internal static class DxpHtmlChunkImporter
{
    private const string MissingImagePixel =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4//8/AwAI/AL+p5qgoAAAAABJRU5ErkJggg==";

    internal static void Append(Section section, Stream stream, string bodyFont,
        Action<Table, double, double, double>? onBorderedTable = null)
    {
        using var reader = new StreamReader(stream);
        var document = new HtmlParser().ParseDocument(reader.ReadToEnd());
        if (document.Body == null) return;
        var initialCount = section.Elements.Count;
        foreach (var child in document.Body.ChildNodes)
            AppendBlock(section, child, bodyFont, initialCount, onBorderedTable);
    }

    private static void AppendBlock(Section section, INode node, string bodyFont, int initialCount,
        Action<Table, double, double, double>? onBorderedTable)
    {
        if (node is not IElement element)
        {
            if (!string.IsNullOrWhiteSpace(node.TextContent))
                AppendParagraph(section, node, null, initialCount);
            return;
        }

        var tag = element.LocalName.ToLowerInvariant();
        if (tag == "table")
        {
            AppendTable(section, element, bodyFont, onBorderedTable);
            return;
        }
        if (tag is "ul" or "ol")
        {
            var items = element.Children.Where(child => child.LocalName == "li").ToArray();
            for (var index = 0; index < items.Length; index++)
            {
                var paragraph = section.AddParagraph();
                paragraph.Format.Font.Name = bodyFont;
                paragraph.Format.Font.Size = Unit.FromPoint(12);
                paragraph.Format.LeftIndent = Unit.FromPoint(18);
                paragraph.Format.AddTabStop(Unit.FromPoint(36));
                paragraph.Format.SpaceAfter = Unit.FromPoint(index == items.Length - 1 ? 14 : 0);
                var marker = paragraph.AddFormattedText(tag == "ul" ? "●" :
                    (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
                if (tag == "ul")
                {
                    marker.Font.Name = "Arial";
                    marker.Font.Size = Unit.FromPoint(9);
                }
                paragraph.AddTab();
                AppendInline(paragraph, items[index], false, false, false);
            }
            return;
        }
        if (tag is "div" or "section" or "article" &&
            element.Children.Any(child => IsBlock(child.LocalName)))
        {
            foreach (var child in element.ChildNodes)
                AppendBlock(section, child, bodyFont, initialCount, onBorderedTable);
            return;
        }
        if (IsBlock(tag))
            AppendParagraph(section, element, tag, initialCount);
        else if (!string.IsNullOrWhiteSpace(element.TextContent))
            AppendParagraph(section, element, null, initialCount);
    }

    private static bool IsBlock(string tag)
        => tag is "p" or "div" or "ul" or "ol" or "table" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6";

    private static void AppendTable(Section section, IElement element, string bodyFont,
        Action<Table, double, double, double>? onBorderedTable)
    {
        var sourceRows = element.QuerySelectorAll("tr")
            .Select(row => row.Children.Where(cell => cell.LocalName is "td" or "th").ToArray())
            .Where(cells => cells.Length > 0).ToArray();
        if (sourceRows.Length == 0) return;
        var border = element.HasAttribute("border")
            ? ParsePixels(element.GetAttribute("border"), 1) * 0.72 : 0;
        var spacing = ParsePixels(element.GetAttribute("cellspacing"), 2) * 0.75;
        var padding = ParsePixels(element.GetAttribute("cellpadding"), 1) * 0.75;
        var separateCells = onBorderedTable != null && border > 0;
        var paddingAdjustment = separateCells ? padding - 3 : 0;
        var spacingAdjustment = separateCells ? spacing - 1.5 : 0;
        var borderAdjustment = separateCells ? border - 0.72 : 0;
        var table = section.AddTable();
        table.Rows.LeftIndent = Unit.FromPoint(-6);
        table.LeftPadding = Unit.FromPoint(Math.Max(0, 5.3 - borderAdjustment -
            (separateCells && spacing == 0 ? border / 2 : 0)));
        table.RightPadding = Unit.FromPoint(separateCells && padding < 3 ? -5 : 0);
        table.TopPadding = Unit.FromPoint(5.36 + paddingAdjustment + spacingAdjustment);
        table.BottomPadding = Unit.FromPoint(6);
        if (element.HasAttribute("border"))
        {
            table.Borders.Width = Unit.FromPoint(onBorderedTable == null ? 0.72 : border);
            table.Borders.Color = onBorderedTable == null
                ? Color.FromRgb(192, 192, 192) : Colors.White;
            if (separateCells)
                onBorderedTable!(table, border, spacing, padding);
        }
        var columnCount = sourceRows.Max(cells => cells.Length);
        for (var column = 0; column < columnCount; column++)
            table.AddColumn(Unit.FromPoint(60 + 2 * paddingAdjustment + spacingAdjustment));
        foreach (var sourceCells in sourceRows)
        {
            var row = table.AddRow();
            row.HeightRule = RowHeightRule.Exactly;
            row.Height = Unit.FromPoint(33.24 + 2 * paddingAdjustment +
                spacingAdjustment - borderAdjustment);
            for (var column = 0; column < sourceCells.Length; column++)
            {
                var cell = row.Cells[column];
                var paragraph = cell.AddParagraph();
                if (column > 0) paragraph.Format.LeftIndent = Unit.FromPoint(-0.84);
                paragraph.Format.Font.Name = bodyFont;
                paragraph.Format.Font.Size = Unit.FromPoint(12);
                paragraph.Format.Font.Bold = sourceCells[column].LocalName == "th";
                AppendInline(paragraph, sourceCells[column], false, false, false);
            }
        }
        var spacer = section.AddParagraph();
        spacer.Format.SpaceAfter = Unit.FromPoint((element.NextElementSibling == null ? 12.67 : 1.7) +
            spacingAdjustment + borderAdjustment);
    }

    private static double ParsePixels(string? value, double fallback)
        => double.TryParse(value, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out var pixels) &&
            double.IsFinite(pixels) && pixels >= 0
            ? pixels : fallback;

    private static void AppendParagraph(Section section, INode node, string? tag,
        int initialCount)
    {
        var firstImportedParagraph = section.Elements.Count == initialCount;
        var paragraph = section.AddParagraph();
        paragraph.Format.Font.Name = "Times New Roman";
        paragraph.Format.Font.Size = Unit.FromPoint(tag switch
        {
            "h1" => 24,
            "h2" => 18,
            "h3" => 14,
            _ => 12
        });
        paragraph.Format.Font.Bold = tag is "h1" or "h2" or "h3" or "h4" or "h5" or "h6";
        if (tag is "h1" or "h2" or "h3")
        {
            paragraph.Format.SpaceBefore = Unit.FromPoint(14.1);
            paragraph.Format.SpaceAfter = Unit.FromPoint(14);
        }
        else
        {
            if (firstImportedParagraph && initialCount > 0)
                paragraph.Format.SpaceBefore = Unit.FromPoint(14);
            paragraph.Format.SpaceAfter = Unit.FromPoint(14);
        }
        AppendInline(paragraph, node, false, false, false);
        if (paragraph.Elements.OfType<MigraDoc.DocumentObjectModel.Shapes.Image>().Any())
            paragraph.Format.SpaceBefore = Unit.FromPoint(paragraph.Format.SpaceBefore.Point - 0.49);
    }

    private static void AppendInline(MigraDoc.DocumentObjectModel.Paragraph paragraph, INode node,
        bool bold, bool italic, bool underline)
    {
        if (node is IElement element)
        {
            var tag = element.LocalName.ToLowerInvariant();
            if (tag == "img")
            {
                var source = element.GetAttribute("src") ?? "";
                var match = Regex.Match(source,
                    @"^data:image/(?:png|jpeg);base64,(?<data>[A-Za-z0-9+/=]+)$",
                    RegexOptions.IgnoreCase);
                if (!double.TryParse(element.GetAttribute("width"),
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var widthPixels) ||
                    !double.TryParse(element.GetAttribute("height"),
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var heightPixels) ||
                    widthPixels <= 0 || heightPixels <= 0)
                    return;
                var image = paragraph.AddImage("base64:" +
                    (match.Success ? match.Groups["data"].Value : MissingImagePixel));
                image.Width = Unit.FromPoint(widthPixels * 0.75);
                image.Height = Unit.FromPoint(heightPixels * 0.75);
                if (!match.Success)
                {
                    image.LineFormat.Visible = true;
                    image.LineFormat.Color = Color.FromRgb(192, 192, 192);
                    image.LineFormat.Width = Unit.FromPoint(0.5);
                }
                return;
            }
            if (tag == "br")
            {
                paragraph.AddLineBreak();
                return;
            }
            bold |= tag is "b" or "strong";
            italic |= tag is "i" or "em";
            underline |= tag == "u";
            foreach (var child in element.ChildNodes)
                AppendInline(paragraph, child, bold, italic, underline);
            return;
        }
        if (node.NodeType != NodeType.Text) return;
        var value = Regex.Replace(node.TextContent, @"\s+", " ");
        if (node.PreviousSibling is IElement previous &&
            previous.LocalName.Equals("img", StringComparison.OrdinalIgnoreCase) &&
            paragraph.Elements.Count > 0 &&
            paragraph.Elements[paragraph.Elements.Count - 1] is MigraDoc.DocumentObjectModel.Shapes.Image)
            value = value.TrimStart(' ');
        if (value.Length == 0) return;
        var span = paragraph.AddFormattedText(value);
        if (bold) span.Bold = true;
        if (italic) span.Italic = true;
        if (underline) span.Underline = Underline.Single;
    }
}
