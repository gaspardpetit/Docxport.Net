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
		if (style.Border == null)
			return null;
		var sb = new StringBuilder();
		AppendCssProperty(sb, "border", style.Border.ToCssValue());
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
