using System.Globalization;
using System.Text;

namespace DocxportNet.API;

public static class DxpComputedParagraphStyleCssExtensions
{
    private static void AppendCssProperty(StringBuilder sb, string name, string value)
    {
        if (sb.Length > 0 && sb[sb.Length - 1] != ';')
            sb.Append(';');
        sb.Append(name).Append(':').Append(value).Append(';');
    }

    public static string? ToCss(this DxpComputedParagraphStyle style, bool includeTextAlign = true)
    {
        if (style.MarginLeftPt == null && style.MarginTopPt == null && style.MarginBottomPt == null && style.TextAlign == null && style.LineHeightCss == null && style.Borders == null && style.BackgroundColorCss == null)
            return null;

        var sb = new StringBuilder();

        // Keep existing ordering used by visitors (margin-left then text-align).
        if (style.MarginLeftPt is double ml && ml > 0.0001)
            AppendCssProperty(sb, "margin-left", ml.ToString("0.###", CultureInfo.InvariantCulture) + "pt");

        if (style.MarginTopPt is double mt)
            AppendCssProperty(sb, "margin-top", mt.ToString("0.###", CultureInfo.InvariantCulture) + "pt");

        if (style.MarginBottomPt is double mb)
            AppendCssProperty(sb, "margin-bottom", mb.ToString("0.###", CultureInfo.InvariantCulture) + "pt");

        if (includeTextAlign && style.TextAlign != null)
        {
            var v = style.TextAlign.Value switch {
                DxpComputedTextAlign.Left => "left",
                DxpComputedTextAlign.Center => "center",
                DxpComputedTextAlign.Right => "right",
                DxpComputedTextAlign.Justify => "justify",
                _ => "left"
            };
            AppendCssProperty(sb, "text-align", v);
        }

        if (!string.IsNullOrWhiteSpace(style.LineHeightCss))
            AppendCssProperty(sb, "line-height", style.LineHeightCss!);

        if (style.Borders != null)
        {
            // Prefer per-side border properties (paragraph borders can vary).
            if (style.Borders.Top != null)
                AppendCssProperty(sb, "border-top", style.Borders.Top.ToCssValue());
            if (style.Borders.Right != null)
                AppendCssProperty(sb, "border-right", style.Borders.Right.ToCssValue());
            if (style.Borders.Bottom != null)
                AppendCssProperty(sb, "border-bottom", style.Borders.Bottom.ToCssValue());
            if (style.Borders.Left != null)
                AppendCssProperty(sb, "border-left", style.Borders.Left.ToCssValue());
        }

        if (!string.IsNullOrWhiteSpace(style.BackgroundColorCss))
            AppendCssProperty(sb, "background-color", style.BackgroundColorCss!);

        return sb.Length == 0 ? null : sb.ToString();
    }
}
