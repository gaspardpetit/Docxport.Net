using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker.Context;

internal static class DxpParagraphStyleComputer
{
    public static DxpComputedParagraphStyle ComputeParagraphStyle(Paragraph p, DxpStyleEffectiveIndentTwips indent, DxpDocumentContext d)
    {
        double? marginLeftPt = null;
        if (indent.Left.HasValue)
        {
            double pt = DxpTwipValue.ToPoints(indent.Left.Value);
            if (pt > 0.0001)
                marginLeftPt = pt;
        }

        double? textIndentPt = null;
        if (indent.FirstLine.HasValue)
        {
            double pt = DxpTwipValue.ToPoints(indent.FirstLine.Value);
            if (Math.Abs(pt) > 0.0001)
                textIndentPt = pt;
        }
        else if (indent.Hanging.HasValue)
        {
            double pt = -DxpTwipValue.ToPoints(indent.Hanging.Value);
            if (Math.Abs(pt) > 0.0001)
                textIndentPt = pt;
        }

        DxpComputedTextAlign? align = null;
        var justification = p.ParagraphProperties?.Justification?.Val?.Value;
        if (justification == null && d.Styles is DxpStyleResolver resolver)
            justification = resolver.GetJustification(p);
        if (justification == JustificationValues.Center)
            align = DxpComputedTextAlign.Center;
        else if (justification == JustificationValues.Right)
            align = DxpComputedTextAlign.Right;
        else if (justification == JustificationValues.Both || justification == JustificationValues.Distribute)
            align = DxpComputedTextAlign.Justify;

        var borders = ComputeBorders(p.ParagraphProperties?.ParagraphBorders);
        var shd = p.ParagraphProperties?.Shading;
        var spacing = p.ParagraphProperties?.SpacingBetweenLines;
        if (d.Styles is DxpStyleResolver r)
        {
            borders = ComputeBorders(r.GetParagraphBorders(p));
            shd = r.GetParagraphShading(p);
            spacing = r.GetParagraphSpacing(p);
        }
        var background = ComputeBackground(shd);
        var (marginTopPt, marginBottomPt, lineHeightCss) = ComputeSpacing(spacing);

        return new DxpComputedParagraphStyle(
            MarginLeftPt: marginLeftPt,
            TextIndentPt: textIndentPt,
            MarginTopPt: marginTopPt,
            MarginBottomPt: marginBottomPt,
            TextAlign: align,
            LineHeightCss: lineHeightCss,
            Borders: borders,
            BackgroundColorCss: background);
    }

    private static (double? marginTopPt, double? marginBottomPt, string? lineHeightCss) ComputeSpacing(SpacingBetweenLines? spacing)
    {
        if (spacing == null)
            return (null, null, null);

        double? mt = null;
        double? mb = null;

        // before/after are in twips.
        if (DxpMeasurement.TryParse(spacing.Before?.Value, out var before))
            mt = before.Points;
        if (DxpMeasurement.TryParse(spacing.After?.Value, out var after))
            mb = after.Points;

        string? lineHeightCss = null;
        if (spacing.Line?.Value != null)
        {
            var rule = spacing.LineRule?.Value;
            // Auto line spacing uses 240ths of a line, not a physical length.
            if (rule == LineSpacingRuleValues.Exact || rule == LineSpacingRuleValues.AtLeast)
            {
                if (DxpMeasurement.TryParse(spacing.Line.Value, out var line) && line.Twips > 0)
                {
                    lineHeightCss = line.Points.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + "pt";
                }
            }
            else if (int.TryParse(spacing.Line.Value, out int lineUnits) && lineUnits > 0)
            {
                double mult = lineUnits / 240.0;
                lineHeightCss = mult.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        return (mt, mb, lineHeightCss);
    }

    internal static DxpComputedBoxBorders? ComputeBorders(ParagraphBorders? bdr)
    {
        if (bdr == null)
            return null;

        var top = ToComputedBorder(bdr.TopBorder);
        var right = ToComputedBorder(bdr.RightBorder);
        var bottom = ToComputedBorder(bdr.BottomBorder);
        var left = ToComputedBorder(bdr.LeftBorder);

        if (top == null && right == null && bottom == null && left == null)
            return null;

        return new DxpComputedBoxBorders(top, right, bottom, left);
    }

    private static string? ComputeBackground(Shading? shd)
    {
        var fill = shd?.Fill?.Value;
        if (string.IsNullOrWhiteSpace(fill) || string.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase))
            return null;
        return ToCssColor(fill!);
    }

    internal static DxpComputedBorder? ToComputedBorder(BorderType? b)
    {
        if (b == null)
            return null;

        var val = b.Val?.Value;
        // For paragraphs, treat "none/nil" as absent to avoid emitting redundant `border-*:none`.
        if (val == BorderValues.None || val == BorderValues.Nil)
            return null;

        int sizeEighthPoints = b.Size != null ? (int)b.Size.Value : 0;
        if (sizeEighthPoints <= 0)
            return null;

        double pt = sizeEighthPoints / 8.0;
        var line = MapLineStyle(val);
        string? color = b.Color?.Value;
        if (string.IsNullOrEmpty(color) || string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase))
            color = "#000000";
        else
            color = ToCssColor(color!);

        return new DxpComputedBorder(WidthPt: pt, LineStyle: line, ColorCss: color);
    }

    private static DxpComputedBorderLineStyle MapLineStyle(BorderValues? val)
    {
        if (val == BorderValues.Dotted)
            return DxpComputedBorderLineStyle.Dotted;
        if (val == BorderValues.DashSmallGap || val == BorderValues.Dashed)
            return DxpComputedBorderLineStyle.Dashed;
        if (val == BorderValues.DotDash)
            return DxpComputedBorderLineStyle.DashDot;
        if (val == BorderValues.DotDotDash)
            return DxpComputedBorderLineStyle.DashDotDot;
        if (val == BorderValues.Double)
            return DxpComputedBorderLineStyle.Double;
        return DxpComputedBorderLineStyle.Solid;
    }

    private static string ToCssColor(string color)
    {
        if (color.StartsWith("#", StringComparison.Ordinal))
            return color;
        if (color.Length is 6 or 3)
            return "#" + color;
        return color;
    }
}
