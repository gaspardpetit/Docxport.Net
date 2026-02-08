using System.Globalization;
using System.Text;

namespace DocxportNet.API;

public static class DxpComputedTableStyleCssExtensions
{
    private static void AppendCssProperty(StringBuilder sb, string name, string value)
    {
        if (sb.Length > 0 && sb[sb.Length - 1] != ';')
            sb.Append(';');
        sb.Append(name).Append(':').Append(value).Append(';');
    }

    private static bool TryCollapseBoxToBorder(DxpComputedBoxBorders borders, out DxpComputedBorder? border)
    {
        border = null;
        var top = borders.Top;
        if (top == null)
            return false;
        if (!Equals(top, borders.Right) || !Equals(top, borders.Bottom) || !Equals(top, borders.Left))
            return false;
        border = top;
        return true;
    }

    private static void AppendBoxBorderCss(StringBuilder sb, DxpComputedBoxBorders borders)
    {
        if (TryCollapseBoxToBorder(borders, out var b) && b != null)
        {
            AppendCssProperty(sb, "border", b.ToCssValue());
            return;
        }

        if (borders.Top != null)
            AppendCssProperty(sb, "border-top", borders.Top.ToCssValue());
        if (borders.Right != null)
            AppendCssProperty(sb, "border-right", borders.Right.ToCssValue());
        if (borders.Bottom != null)
            AppendCssProperty(sb, "border-bottom", borders.Bottom.ToCssValue());
        if (borders.Left != null)
            AppendCssProperty(sb, "border-left", borders.Left.ToCssValue());
    }

    public static string? ToCss(this DxpComputedTableStyle style)
    {
        if (style.TableBorder == null && style.TableBorders == null && style.BorderCollapse == false)
            return null;

        var sb = new StringBuilder();
        if (style.TableBorders != null)
            AppendBoxBorderCss(sb, style.TableBorders);
        else if (style.TableBorder != null)
            AppendCssProperty(sb, "border", style.TableBorder.ToCssValue());
        if (style.BorderCollapse)
            AppendCssProperty(sb, "border-collapse", "collapse");
        return sb.Length == 0 ? null : sb.ToString();
    }

    public static string? ToCss(this DxpComputedTableCellStyle style)
    {
        if (style.Border == null && style.Borders == null && style.BackgroundColorCss == null && style.VerticalAlign == null)
            return null;
        var sb = new StringBuilder();
        if (style.Borders != null)
            AppendBoxBorderCss(sb, style.Borders);
        else if (style.Border != null)
            AppendCssProperty(sb, "border", style.Border.ToCssValue());
        if (!string.IsNullOrWhiteSpace(style.BackgroundColorCss))
            AppendCssProperty(sb, "background-color", style.BackgroundColorCss!);
        if (style.VerticalAlign != null)
        {
            var v = style.VerticalAlign.Value switch {
                DxpComputedVerticalAlign.Top => "top",
                DxpComputedVerticalAlign.Middle => "middle",
                DxpComputedVerticalAlign.Bottom => "bottom",
                _ => "baseline"
            };
            AppendCssProperty(sb, "vertical-align", v);
        }
        return sb.ToString();
    }

    public static string ToCssValue(this DxpComputedBorder border)
    {
        var line = border.LineStyle switch {
            DxpComputedBorderLineStyle.None => "none",
            DxpComputedBorderLineStyle.Solid => "solid",
            DxpComputedBorderLineStyle.Dotted => "dotted",
            DxpComputedBorderLineStyle.Dashed => "dashed",
            DxpComputedBorderLineStyle.Double => "double",
            _ => "solid"
        };
        if (border.LineStyle == DxpComputedBorderLineStyle.None)
            return "none";

        // CSS `double` borders often render as a single line at very small widths; apply a minimal
        // width so the double-strike effect is visible and closer to Word’s rendering.
        var widthPt = border.WidthPt;
        if (border.LineStyle == DxpComputedBorderLineStyle.Double && widthPt < 2.25)
            widthPt = 2.25;

        var pt = widthPt.ToString("0.###", CultureInfo.InvariantCulture) + "pt";
        return pt + " " + line + " " + border.ColorCss;
    }
}
