using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker;

// Intentionally minimal for now: mirrors existing border-only behavior, but gives us a single place
// to add full table-style resolution later (tblStyle chain, tblLook conditional formatting, etc.).
internal static class DxpTableStyleResolver
{
	public static DxpComputedTableStyle ComputeTableStyle(TableProperties? tableProperties)
		=> DxpTableStyleComputer.ComputeTableStyle(tableProperties);

	public static DxpComputedTableCellStyle ComputeCellStyle(TableCellProperties? cellProperties, DxpComputedTableStyle tableStyle)
		=> DxpTableStyleComputer.ComputeCellStyle(cellProperties, tableStyle);
}

