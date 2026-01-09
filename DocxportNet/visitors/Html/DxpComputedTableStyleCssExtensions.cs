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

	public static string? ToCss(this DxpComputedTableStyle style)
	{
		if (style.TableBorder == null && style.BorderCollapse == false)
			return null;

		var sb = new StringBuilder();
		if (style.TableBorder != null)
			AppendCssProperty(sb, "border", style.TableBorder.ToCssValue());
		if (style.BorderCollapse)
			AppendCssProperty(sb, "border-collapse", "collapse");
		return sb.Length == 0 ? null : sb.ToString();
	}

	public static string? ToCss(this DxpComputedTableCellStyle style)
	{
		if (style.Border == null && style.BackgroundColorCss == null && style.VerticalAlign == null)
			return null;
		var sb = new StringBuilder();
		if (style.Border != null)
			AppendCssProperty(sb, "border", style.Border.ToCssValue());
		if (!string.IsNullOrWhiteSpace(style.BackgroundColorCss))
			AppendCssProperty(sb, "background-color", style.BackgroundColorCss!);
		if (style.VerticalAlign != null)
		{
			var v = style.VerticalAlign.Value switch
			{
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
		var pt = border.WidthPt.ToString("0.###", CultureInfo.InvariantCulture) + "pt";
		var line = border.LineStyle switch
		{
			DxpComputedBorderLineStyle.Solid => "solid",
			_ => "solid"
		};
		return pt + " " + line + " " + border.ColorCss;
	}
}
