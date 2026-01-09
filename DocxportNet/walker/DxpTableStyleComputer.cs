using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

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
		var direct = ComputeDirectCellStyle(cellProperties);
		var border = direct.Border ?? tableStyle.DefaultCellBorder;
		return new DxpComputedTableCellStyle(Border: border, BackgroundColorCss: direct.BackgroundColorCss, VerticalAlign: direct.VerticalAlign);
	}

	public static DxpComputedTableCellStyle ComputeDirectCellStyle(TableCellProperties? cellProperties)
	{
		var b = PickBorder(cellProperties?.TableCellBorders);
		var border = ToComputedBorder(b);

		string? background = null;
		var fill = cellProperties?.Shading?.Fill?.Value;
		if (!string.IsNullOrWhiteSpace(fill) && !string.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase))
			background = ToCssColor(fill!);

		DxpComputedVerticalAlign? verticalAlign = null;
		var v = cellProperties?.TableCellVerticalAlignment?.Val?.Value;
		if (v != null)
		{
			if (v == TableVerticalAlignmentValues.Top)
				verticalAlign = DxpComputedVerticalAlign.Top;
			else if (v == TableVerticalAlignmentValues.Center)
				verticalAlign = DxpComputedVerticalAlign.Middle;
			else if (v == TableVerticalAlignmentValues.Bottom)
				verticalAlign = DxpComputedVerticalAlign.Bottom;
		}

		return new DxpComputedTableCellStyle(Border: border, BackgroundColorCss: background, VerticalAlign: verticalAlign);
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
