using System.Globalization;
using System.Text;

namespace DocxportNet.API;

public static class DxpComputedTableStyleCssExtensions
{
	public static string? ToCss(this DxpComputedTableStyle style)
	{
		if (style.TableBorder == null && style.BorderCollapse == false)
			return null;

		var sb = new StringBuilder();
		if (style.TableBorder != null)
			sb.Append("border:").Append(style.TableBorder.ToCssValue()).Append(';');
		if (style.BorderCollapse)
			sb.Append("border-collapse:collapse;");
		return sb.Length == 0 ? null : sb.ToString();
	}

	public static string? ToCss(this DxpComputedTableCellStyle style)
	{
		if (style.Border == null)
			return null;
		return "border:" + style.Border.ToCssValue() + ";";
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

