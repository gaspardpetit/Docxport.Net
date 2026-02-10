using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using System.Globalization;

namespace DocxportNet.Walker.Parts;

public class DxpTables
{
    private static IEnumerable<TableCell> EnumerateRowCells(TableRow row)
    {
        foreach (var child in row.ChildElements)
        {
            switch (child)
            {
                case TableCell tc:
                    yield return tc;
                    break;

                case SdtCell sdtCell:
                {
                    var content = sdtCell.SdtContentCell;
                    if (content == null)
                        break;
                    foreach (var inner in content.Elements<TableCell>())
                        yield return inner;
                    break;
                }

                case CustomXmlCell cxCell:
                    foreach (var inner in cxCell.Elements<TableCell>())
                        yield return inner;
                    break;
            }
        }
    }

    private static IReadOnlyList<int?> ReadTblGridTwips(Table t)
    {
        var tg = t.Elements<TableGrid>().FirstOrDefault();
        if (tg == null)
            return [];

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

    public DxpTableModel BuildTableModel(Table t)
    {
        var grid = ReadTblGridTwips(t);
        int colCount = grid.Count;

        // If no tblGrid, fallback: compute max columns by summing gridSpan in each row
        if (colCount == 0)
        {
            colCount = t.Elements<TableRow>()
                .Select(r => EnumerateRowCells(r).Sum(tc => GetGridSpan(tc)))
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
            foreach (var tc in EnumerateRowCells(rows[r]))
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
