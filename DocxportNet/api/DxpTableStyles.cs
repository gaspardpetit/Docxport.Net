namespace DocxportNet.API;

public enum DxpComputedBorderLineStyle
{
	Solid
}

public sealed record DxpComputedBorder(
	double WidthPt,
	DxpComputedBorderLineStyle LineStyle,
	string ColorCss
);

public sealed record DxpComputedTableStyle(
	DxpComputedBorder? TableBorder,
	bool BorderCollapse,
	DxpComputedBorder? DefaultCellBorder
);

public sealed record DxpComputedTableCellStyle(
	DxpComputedBorder? Border
);
