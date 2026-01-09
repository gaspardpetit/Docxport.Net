using DocxportNet.API;

namespace DocxportNet.Tests;

public class ComputedTableStyleCssTests
{
	[Fact]
	public void Table_ToCss_IsNullWhenEmpty()
	{
		var style = new DxpComputedTableStyle(TableBorder: null, BorderCollapse: false, DefaultCellBorder: null);
		Assert.Null(style.ToCss());
	}

	[Fact]
	public void Table_ToCss_HasStablePropertyOrder()
	{
		var border = new DxpComputedBorder(1.25, DxpComputedBorderLineStyle.Solid, "#000000");
		var style = new DxpComputedTableStyle(TableBorder: border, BorderCollapse: true, DefaultCellBorder: border);
		Assert.Equal("border:1.25pt solid #000000;border-collapse:collapse;", style.ToCss());
	}

	[Fact]
	public void Cell_ToCss_UsesBorderProperty()
	{
		var border = new DxpComputedBorder(0.5, DxpComputedBorderLineStyle.Solid, "#ff00ff");
		var style = new DxpComputedTableCellStyle(Border: border, BackgroundColorCss: null, VerticalAlign: null);
		Assert.Equal("border:0.5pt solid #ff00ff;", style.ToCss());
	}

	[Fact]
	public void Cell_ToCss_IncludesBackgroundAndVerticalAlign()
	{
		var border = new DxpComputedBorder(1, DxpComputedBorderLineStyle.Solid, "#000000");
		var style = new DxpComputedTableCellStyle(Border: border, BackgroundColorCss: "#ffffff", VerticalAlign: DxpComputedVerticalAlign.Middle);
		Assert.Equal("border:1pt solid #000000;background-color:#ffffff;vertical-align:middle;", style.ToCss());
	}
}
