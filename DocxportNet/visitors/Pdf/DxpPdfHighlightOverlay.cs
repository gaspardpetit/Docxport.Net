using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Symbols;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace DocxportNet.Visitors.Pdf;

internal sealed record DxpPdfHighlight(string Text, HighlightColorValues? Color, string? Shading = null,
    double LineSize = 0, bool Footer = false, double RightPadding = 0,
    StringBuilder? ParagraphText = null, int SourceOffset = 0);

internal static class DxpPdfHighlightOverlay
{
    private static readonly Regex UnicodeMapTarget = new(
        @"(<[0-9A-Fa-f]{4}>\s*(?:<[0-9A-Fa-f]{4}>\s*)?<)([0-9A-Fa-f]{4,6})(>)");

    public static void Save(PdfDocument document, Stream output, IReadOnlyList<DxpPdfHighlight> highlights)
    {
        var repairAfterSave = RepairFontUnicodeMaps(document);
        if (highlights.Count == 0)
        {
            SaveWithoutHighlights(document, output, repairAfterSave);
            return;
        }
        var colors = new Dictionary<string, (XColor Color, bool Shading, double LineSize, bool Footer,
            double RightPadding)>(
            StringComparer.Ordinal);
        var ambiguous = new HashSet<string>(StringComparer.Ordinal);
        var anchored = new List<(DxpPdfHighlight Highlight, XColor Color)>();
        foreach (var highlight in highlights)
        {
            var text = Normalize(highlight.Text);
            if (text.Length < 3 || !(highlight.Color is { } value && TryColor(value, out var color) ||
                TryShadingColor(highlight.Shading, out color))) continue;
            if (highlight.ParagraphText != null)
            {
                anchored.Add((highlight, color));
                continue;
            }
            if (colors.ContainsKey(text)) ambiguous.Add(text);
            colors[text] = (color, highlight.Shading != null, highlight.LineSize, highlight.Footer,
                highlight.RightPadding);
        }
        foreach (var text in ambiguous) colors.Remove(text);
        if (colors.Count == 0 && anchored.Count == 0)
        {
            SaveWithoutHighlights(document, output, repairAfterSave);
            return;
        }
        using var stream = new MemoryStream();
        document.Save(stream, false);
        stream.Position = 0;
        using var parsed = PdfPigDocument.Open(stream);
        using var modified = PdfReader.Open(new MemoryStream(stream.ToArray()), PdfDocumentOpenMode.Modify);
        var pageTexts = parsed.GetPages()
            .Select(page => Normalize(string.Concat(page.Letters.Select(letter => letter.Value))))
            .ToArray();
        foreach (var text in colors.Keys.ToArray())
        {
            var occurrences = 0;
            foreach (var pageText in pageTexts)
                for (var offset = 0; (offset = pageText.IndexOf(text, offset,
                         StringComparison.Ordinal)) >= 0 && occurrences < 2; offset += text.Length)
                    occurrences++;
            if (occurrences != 1) colors.Remove(text);
        }
        var anchoredMatches = new Dictionary<int, List<(int Offset, int Length, XColor Color,
            bool Shading, double LineSize, bool Footer, double RightPadding)>>();
        foreach (var (highlight, color) in anchored)
        {
            var paragraphText = Normalize(highlight.ParagraphText!.ToString());
            var text = Normalize(highlight.Text);
            if (highlight.SourceOffset < 0 || highlight.SourceOffset + text.Length > paragraphText.Length ||
                string.CompareOrdinal(paragraphText, highlight.SourceOffset, text, 0, text.Length) != 0)
                continue;
            var matchPage = -1;
            var matchOffset = -1;
            var occurrences = 0;
            for (var pageIndex = 0; pageIndex < pageTexts.Length && occurrences < 2; pageIndex++)
                for (var offset = 0; (offset = pageTexts[pageIndex].IndexOf(paragraphText, offset,
                         StringComparison.Ordinal)) >= 0 && occurrences < 2; offset += paragraphText.Length)
                {
                    occurrences++;
                    matchPage = pageIndex;
                    matchOffset = offset + highlight.SourceOffset;
                }
            if (occurrences != 1) continue;
            if (!anchoredMatches.TryGetValue(matchPage, out var pageMatches))
                anchoredMatches[matchPage] = pageMatches = new();
            pageMatches.Add((matchOffset, text.Length, color, highlight.Shading != null,
                highlight.LineSize, highlight.Footer, highlight.RightPadding));
        }
        if (colors.Count == 0 && anchoredMatches.Count == 0)
        {
            if (repairAfterSave) RepairFontUnicodeMaps(modified);
            modified.Save(output, false);
            return;
        }
        for (var pageIndex = 0; pageIndex < modified.PageCount; pageIndex++)
        {
            var page = parsed.GetPage(pageIndex + 1);
            var letters = new List<Letter>();
            var whitespaceBefore = new List<bool>();
            var content = new StringBuilder();
            var pendingWhitespace = false;
            foreach (var letter in page.Letters)
            {
                foreach (var character in letter.Value)
                {
                    if (char.IsWhiteSpace(character))
                    {
                        pendingWhitespace = true;
                        continue;
                    }
                    content.Append(character);
                    letters.Add(letter);
                    whitespaceBefore.Add(pendingWhitespace);
                    pendingWhitespace = false;
                }
            }
            using var graphics = XGraphics.FromPdfPage(modified.Pages[pageIndex], XGraphicsPdfPageOptions.Prepend);
            var pageText = content.ToString();
            var matches = new List<(int Offset, int Length, XColor Color, bool Shading,
                double LineSize, bool Footer,
                double RightPadding)>();
            foreach (var (text, style) in colors)
            {
                for (var offset = 0; (offset = pageText.IndexOf(text, offset, StringComparison.Ordinal)) >= 0;
                    offset += text.Length)
                    matches.Add((offset, text.Length, style.Color, style.Shading, style.LineSize, style.Footer,
                        style.RightPadding));
            }
            if (anchoredMatches.TryGetValue(pageIndex, out var sourceMatches))
                matches.AddRange(sourceMatches);
            var highlighted = new Dictionary<int, XColor>();
            foreach (var (offset, length, color, _, _, _, _) in matches)
                for (var index = offset; index < offset + length; index++) highlighted[index] = color;
            var rectangles = new List<(int Offset, int End, XColor Color, double Left,
                double Right, double Top, double Bottom)>();
            foreach (var (offset, length, color, shading, lineSize, footer, rightPadding) in matches)
            {
                var first = letters[offset];
                var last = letters[offset + length - 1];
                if (Math.Abs(first.StartBaseLine.Y - last.StartBaseLine.Y) > 2) continue;
                var size = first.PointSize;
                var left = first.StartBaseLine.X;
                if (offset > 0 && !whitespaceBefore[offset] &&
                    highlighted.TryGetValue(offset - 1, out var previousColor) &&
                    (!shading || previousColor == color) &&
                    Math.Abs(letters[offset - 1].StartBaseLine.Y - first.StartBaseLine.Y) < 2)
                {
                    var previousRight = letters[offset - 1].StartBaseLine.X + letters[offset - 1].Width;
                    if (left - previousRight is >= 0 and <= 4) left = previousRight;
                }
                var right = last.StartBaseLine.X + last.Width + rightPadding;
                var top = page.Height - first.StartBaseLine.Y - size * 0.93;
                var height = first.FontName?.Contains("Arial", StringComparison.OrdinalIgnoreCase) == true
                    ? size * 1.15 : size * 1.22;
                if (size < 10 && lineSize > size && lineSize <= 14)
                {
                    height = lineSize * 1.22;
                    if (footer) top = page.Height - first.StartBaseLine.Y - lineSize * 1.02;
                }
                rectangles.Add((offset, offset + length, color, left, right, top, top + height));
            }
            rectangles.Sort((first, second) => first.Offset.CompareTo(second.Offset));
            for (var index = 0; index < rectangles.Count; index++)
            {
                var current = rectangles[index];
                while (index + 1 < rectangles.Count &&
                    rectangles[index + 1] is var next && next.Offset == current.End &&
                    !whitespaceBefore[next.Offset] &&
                    next.Color == current.Color && next.Left - current.Right is >= -2 and <= 4 &&
                    Math.Min(next.Bottom, current.Bottom) > Math.Max(next.Top, current.Top))
                {
                    current = (current.Offset, next.End, current.Color, current.Left,
                        Math.Max(current.Right, next.Right), Math.Min(current.Top, next.Top),
                        Math.Max(current.Bottom, next.Bottom));
                    index++;
                }
                graphics.DrawRectangle(new XSolidBrush(current.Color), current.Left, current.Top,
                    Math.Max(0, current.Right - current.Left), current.Bottom - current.Top);
            }
        }
        if (repairAfterSave) RepairFontUnicodeMaps(modified);
        modified.Save(output, false);
    }

    private static void SaveWithoutHighlights(PdfDocument document, Stream output, bool repairAfterSave)
    {
        if (!repairAfterSave)
        {
            document.Save(output, false);
            return;
        }
        using var stream = new MemoryStream();
        document.Save(stream, false);
        using var modified = PdfReader.Open(new MemoryStream(stream.ToArray()), PdfDocumentOpenMode.Modify);
        RepairFontUnicodeMaps(modified);
        modified.Save(output, false);
    }

    private static bool RepairFontUnicodeMaps(PdfDocument document)
    {
        var repaired = new HashSet<PdfSharp.Pdf.PdfDictionary>();
        var repairAfterSave = false;
        foreach (PdfPage page in document.Pages)
        {
            var fonts = page.Resources.Elements.GetDictionary("/Font");
            if (fonts == null) continue;
            foreach (var value in fonts.Elements.Values)
            {
                var font = (value as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfSharp.Pdf.PdfDictionary ??
                    value as PdfSharp.Pdf.PdfDictionary;
                var mapValue = font?.Elements["/ToUnicode"];
                var map = (mapValue as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfSharp.Pdf.PdfDictionary ??
                    mapValue as PdfSharp.Pdf.PdfDictionary;
                var symbolFont = font?.Elements.GetName("/BaseFont") is string baseFont &&
                    baseFont.EndsWith("+Symbol", StringComparison.Ordinal);
                var wingdingsFont = font?.Elements.GetName("/BaseFont") is string wingdingsBaseFont &&
                    wingdingsBaseFont.EndsWith("+Wingdings", StringComparison.Ordinal);
                if (map?.Stream == null)
                {
                    if (map != null) repairAfterSave = true;
                    continue;
                }
                if (!repaired.Add(map)) continue;
                var source = Encoding.Latin1.GetString(map.Stream.UnfilteredValue);
                var corrected = RepairUnicodeMapText(source, symbolFont, wingdingsFont);
                if (corrected != source)
                {
                    map.Stream.Value = Encoding.Latin1.GetBytes(corrected);
                    map.Elements.Remove("/Filter");
                    map.Elements.Remove("/DecodeParms");
                }
            }
        }
        return repairAfterSave;
    }

    internal static string RepairUnicodeMapText(string source, bool symbolFont = false,
        bool wingdingsFont = false)
        => UnicodeMapTarget.Replace(source, match =>
        {
            var codepoint = int.Parse(match.Groups[2].Value,
                System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture);
            if (symbolFont && codepoint == 0xF0B7)
                return match.Groups[1].Value + "2022" + match.Groups[3].Value;
            if (wingdingsFont && codepoint is >= 0xF000 and <= 0xF0FF)
            {
                var translated = DxpWingdingsEncoding.Table[codepoint - 0xF000];
                if (translated.Length > 0 && !char.IsControl(translated[0]))
                    return match.Groups[1].Value + string.Concat(translated.Select(character =>
                        ((int)character).ToString("X4"))) + match.Groups[3].Value;
            }
            if (codepoint is < 0x10000 or > 0x10FFFF) return match.Value;
            var surrogate = char.ConvertFromUtf32(codepoint);
            return match.Groups[1].Value + ((int)surrogate[0]).ToString("X4") +
                ((int)surrogate[1]).ToString("X4") + match.Groups[3].Value;
        });

    private static string Normalize(string text) =>
        new(text.Where(character => !char.IsWhiteSpace(character)).ToArray());

    private static bool TryShadingColor(string? fill, out XColor color)
    {
        color = default;
        if (fill?.Length != 6 || !int.TryParse(fill, System.Globalization.NumberStyles.HexNumber,
            System.Globalization.CultureInfo.InvariantCulture, out var rgb)) return false;
        color = XColor.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255);
        return true;
    }

    private static bool TryColor(HighlightColorValues value, out XColor color)
    {
        color = value == HighlightColorValues.Yellow ? XColor.FromArgb(255, 255, 0) :
            value == HighlightColorValues.Green ? XColor.FromArgb(0, 255, 0) :
            value == HighlightColorValues.Cyan ? XColor.FromArgb(0, 255, 255) :
            value == HighlightColorValues.Magenta ? XColor.FromArgb(255, 0, 255) :
            value == HighlightColorValues.Blue ? XColor.FromArgb(0, 0, 255) :
            value == HighlightColorValues.Red ? XColor.FromArgb(255, 0, 0) :
            value == HighlightColorValues.DarkBlue ? XColor.FromArgb(0, 0, 139) :
            value == HighlightColorValues.DarkCyan ? XColor.FromArgb(0, 139, 139) :
            value == HighlightColorValues.DarkGreen ? XColor.FromArgb(0, 100, 0) :
            value == HighlightColorValues.DarkMagenta ? XColor.FromArgb(128, 0, 128) :
            value == HighlightColorValues.DarkRed ? XColor.FromArgb(139, 0, 0) :
            value == HighlightColorValues.DarkYellow ? XColor.FromArgb(128, 128, 0) :
            value == HighlightColorValues.DarkGray ? XColor.FromArgb(169, 169, 169) :
            value == HighlightColorValues.LightGray ? XColor.FromArgb(211, 211, 211) :
            value == HighlightColorValues.Black ? XColor.FromArgb(0, 0, 0) :
            value == HighlightColorValues.White ? XColor.FromArgb(255, 255, 255) : default;
        return value != HighlightColorValues.None;
    }
}
