using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using System.Globalization;

namespace DocxportNet.Walker;

internal static class DxpTableStyleComputer
{
	public static DxpComputedTableStyle ComputeTableStyle(TableProperties? tableProperties)
	{
		var b = PickBorder(tableProperties?.TableBorders);
		var border = ToComputedBorder(b);
		if (border == null)
			return new DxpComputedTableStyle(TableBorder: null, BorderCollapse: false, DefaultCellBorder: null);

		return new DxpComputedTableStyle(
			TableBorder: border,
			BorderCollapse: true,
			DefaultCellBorder: border);
	}

	public static DxpComputedTableCellStyle ComputeCellStyle(TableCellProperties? cellProperties, DxpComputedTableStyle tableStyle)
	{
		var b = PickBorder(cellProperties?.TableCellBorders);
		var border = ToComputedBorder(b) ?? tableStyle.DefaultCellBorder;
		return new DxpComputedTableCellStyle(Border: border);
	}

	private static DxpComputedBorder? ToComputedBorder(BorderType? b)
	{
		if (b == null)
			return null;

		int sizeEighthPoints = b.Size != null ? (int)b.Size.Value : 0;
		if (sizeEighthPoints <= 0)
			return null;

		double pt = sizeEighthPoints / 8.0;
		string? color = b.Color?.Value;
		if (string.IsNullOrEmpty(color) || string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase))
			color = "#000000";
		else
			color = ToCssColor(color!);

		return new DxpComputedBorder(
			WidthPt: pt,
			LineStyle: DxpComputedBorderLineStyle.Solid,
			ColorCss: color);
	}

	private static BorderType? PickBorder(TableBorders? borders)
	{
		if (borders == null)
			return null;

		foreach (var b in new BorderType?[]
			{
				borders.TopBorder,
				borders.LeftBorder,
				borders.BottomBorder,
				borders.RightBorder,
				borders.InsideHorizontalBorder,
				borders.InsideVerticalBorder
			})
		{
			if (b != null)
				return b;
		}

		return null;
	}

	private static BorderType? PickBorder(TableCellBorders? borders)
	{
		if (borders == null)
			return null;

		foreach (var b in new BorderType?[]
			{
				borders.TopBorder,
				borders.LeftBorder,
				borders.BottomBorder,
				borders.RightBorder,
				borders.InsideHorizontalBorder,
				borders.InsideVerticalBorder
			})
		{
			if (b != null)
				return b;
		}

		return null;
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
