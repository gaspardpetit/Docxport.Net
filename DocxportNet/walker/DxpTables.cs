using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using System.Globalization;

namespace DocxportNet.Walker;

public class DxpTables
{
	private static IReadOnlyList<int?> ReadTblGridTwips(Table t)
	{
		var tg = t.Elements<TableGrid>().FirstOrDefault();
		if (tg == null)
			return Array.Empty<int?>();

		var cols = new List<int?>();
		foreach (var gc in tg.Elements<GridColumn>())
		{
			// GridColumn.Width is a StringValue in twips
			if (gc.Width?.Value != null && int.TryParse(gc.Width.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var w))
				cols.Add(w);
			else
				cols.Add(null);
		}
		return cols;
	}

	private static int GetGridSpan(TableCell tc)
	{
		var span = tc.TableCellProperties?.GridSpan?.Val?.Value;
		return span is > 1 ? span.Value : 1;
	}

	private static string? GetVMerge(TableCell tc)
	{
		// null => no merge
		// "restart" => starts a vertical merge region (sometimes val absent = restart depending on producers)
		// "continue" => continues
		var vm = tc.TableCellProperties?.VerticalMerge;
		if (vm == null)
			return null;
		return vm.Val?.Value.ToString(); // "restart"/"continue" or null
	}

	private static bool IsVMergeContinue(TableCell tc)
	{
		var vm = tc.TableCellProperties?.VerticalMerge;
		if (vm == null)
			return false;

		// Word often encodes <w:vMerge/> meaning "continue" in some documents
		if (vm.Val == null)
			return true;

		return vm.Val.Value == MergedCellValues.Continue;
	}

	private static bool IsVMergeRestartOrStart(TableCell tc)
	{
		var vm = tc.TableCellProperties?.VerticalMerge;
		if (vm == null)
			return false;

		// <w:vMerge w:val="restart"/> clearly starts
		if (vm.Val?.Value == MergedCellValues.Restart)
			return true;

		// Some producers use <w:vMerge/> on the first cell too; ambiguous.
		// We'll treat val==null as "continue" ONLY if we can find a master above.
		return false;
	}

	public DxpTableModel BuildTableModel(Table t)
	{
		var grid = ReadTblGridTwips(t);
		int colCount = grid.Count;

		// If no tblGrid, fallback: compute max columns by summing gridSpan in each row
		if (colCount == 0)
		{
			colCount = t.Elements<TableRow>()
				.Select(r => r.Elements<TableCell>().Sum(tc => GetGridSpan(tc)))
				.DefaultIfEmpty(0)
				.Max();
		}

		var rows = t.Elements<TableRow>().ToList();
		int rowCount = rows.Count;

		var cells = new DxpCellModel?[rowCount, colCount];

		// Track vertical merge masters by column (or by grid position)
		// We need the last "master" cell encountered at each column position.
		var vMergeMaster = new DxpCellModel?[colCount];

		for (int r = 0; r < rowCount; r++)
		{
			int c = 0;
			foreach (var tc in rows[r].Elements<TableCell>())
			{
				// find next free column slot (skip covered slots)
				while (c < colCount && cells[r, c] != null)
					c++;

				if (c >= colCount)
					break;

				int span = GetGridSpan(tc);
				var maxSpan = colCount - c;
				span = span < 1 ? 1 : span > maxSpan ? maxSpan : span;

				var cm = new DxpCellModel {
					Cell = tc,
					Row = r,
					Col = c,
					ColSpan = span,
					RowSpan = 1
				};

				// Vertical merge handling:
				// If this cell is "continue", it is covered by the master above at same column.
				if (IsVMergeContinue(tc))
				{
					var master = vMergeMaster[c];
					if (master != null)
					{
						cm.IsCovered = true;
						cm.CoveredBy = master;
						master.RowSpan += 1;
					}
					else
					{
						// No master above; treat as normal cell (best-effort)
						vMergeMaster[c] = cm;
					}
				}
				else
				{
					// Start/normal cell becomes potential master for subsequent continues
					vMergeMaster[c] = cm;
				}

				// Place in grid: master occupies [r,c], its colspan covers additional slots as covered
				cells[r, c] = cm;

				// Mark extra columns covered by colspan
				for (int k = 1; k < span; k++)
				{
					if (c + k >= colCount)
						break;
					cells[r, c + k] = new DxpCellModel {
						Cell = tc,
						Row = r,
						Col = c + k,
						ColSpan = 1,
						RowSpan = 1,
						IsCovered = true,
						CoveredBy = cm
					};
					// For vMerge, keep master pointer consistent across spanned columns too
					vMergeMaster[c + k] = vMergeMaster[c];
				}

				c += span;
			}

			// If row ends with fewer cells, remaining grid is empty (null)
			// Also: if a column has vMerge continue but Word may omit tc entirely in rare cases;
			// handling that requires deeper heuristics and is document-dependent.
		}

		return new DxpTableModel(colCount, rowCount, grid, cells);
	}
}
