using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.api;

public sealed class DxpCellModel
{
	public required TableCell Cell { get; init; }
	public required int Row { get; init; }
	public required int Col { get; init; }

	public int ColSpan { get; set; } = 1;
	public int RowSpan { get; set; } = 1;

	// true if this position is covered by a merge (so should not emit <td>)
	public bool IsCovered { get; set; }

	// points to the master cell (top-left) when covered
	public DxpCellModel? CoveredBy { get; set; }

	public TableCellProperties? TcPr => Cell.TableCellProperties;
}
