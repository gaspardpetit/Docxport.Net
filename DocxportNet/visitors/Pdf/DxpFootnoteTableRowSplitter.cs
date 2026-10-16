using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using System.Reflection;

namespace DocxportNet.Visitors.Pdf;

internal static class DxpFootnoteTableRowSplitter
{
    internal static bool TrySplit(Row row, IReadOnlyList<int> firstFragmentWords,
        out Row? firstFragment, out Row? continuation, int paragraphIndex = 0)
    {
        firstFragment = null;
        continuation = null;
        if (row.Cells.Count == 0 || row.Cells.Count != firstFragmentWords.Count ||
            row.Cells.Cast<Cell>().Any(cell => cell.MergeRight > 0 || cell.MergeDown > 0))
            return false;

        var first = (Row)row.Clone();
        var second = (Row)row.Clone();
        var splitAny = false;
        for (var index = 0; index < row.Cells.Count; index++)
        {
            var firstCell = first.Cells[index];
            var secondCell = second.Cells[index];
            if (paragraphIndex < 0 || firstCell.Elements.Count <= paragraphIndex ||
                firstCell.Elements[paragraphIndex] is not Paragraph paragraph ||
                !DxpFootnoteParagraphSplitter.TryGetText(paragraph, out var content))
                return false;
            var wordCount = System.Text.RegularExpressions.Regex.Matches(content, @"\S+").Count;
            if (firstFragmentWords[index] == wordCount)
            {
                for (var elementIndex = secondCell.Elements.Count - 1; elementIndex >= 0; elementIndex--)
                    secondCell.Elements.RemoveObjectAt(elementIndex);
                firstCell.Borders.Bottom.Visible = false;
                secondCell.Borders.Top.Visible = false;
                continue;
            }
            if (!DxpFootnoteParagraphSplitter.TrySplit(paragraph, firstFragmentWords[index],
                    out var remainingParagraph) || remainingParagraph == null)
                return false;

            splitAny = true;
            secondCell.Elements.RemoveObjectAt(paragraphIndex);
            secondCell.Elements.InsertObject(paragraphIndex, remainingParagraph);
            remainingParagraph.Format.PageBreakBefore = false;
            if (paragraphIndex == 0 && secondCell.Elements.Count == 2 &&
                secondCell.Elements[1] is Paragraph nextParagraph &&
                Math.Abs(remainingParagraph.Format.SpaceAfter.Point - 6) < 0.1 &&
                Math.Abs(nextParagraph.Format.SpaceBefore.Point - 5) < 0.1 &&
                Math.Abs(nextParagraph.Format.SpaceAfter.Point - 5) < 0.1)
            {
                nextParagraph.Format.SpaceBefore = Unit.FromPoint(14);
                nextParagraph.Format.SpaceAfter = Unit.FromPoint(-4);
            }
            for (var elementIndex = firstCell.Elements.Count - 1; elementIndex > paragraphIndex; elementIndex--)
                firstCell.Elements.RemoveObjectAt(elementIndex);
            for (var elementIndex = paragraphIndex - 1; elementIndex >= 0; elementIndex--)
                secondCell.Elements.RemoveObjectAt(elementIndex);
            firstCell.Borders.Bottom.Visible = false;
            secondCell.Borders.Top.Visible = false;
        }
        if (!splitAny) return false;
        first.Borders.Bottom.Visible = false;
        second.Borders.Top.Visible = false;
        second.HeadingFormat = false;
        firstFragment = first;
        continuation = second;
        return true;
    }

    internal static bool TryReplace(Row source, Row firstFragment, Row continuation)
    {
        var table = source.Table;
        if (table == null || source.Index < 0 || source.Index >= table.Rows.Count ||
            !ReferenceEquals(table.Rows[source.Index], source)) return false;
        const BindingFlags flags = BindingFlags.Instance | BindingFlags.NonPublic;
        var rowIndex = typeof(Row).GetField("_idx", flags);
        var rowTable = typeof(Row).GetField("_table", flags);
        var cellRow = typeof(Cell).GetField("_row", flags);
        var cellTable = typeof(Cell).GetField("_table", flags);
        var cellColumn = typeof(Cell).GetField("_clm", flags);
        if (rowIndex == null || rowTable == null || cellRow == null ||
            cellTable == null || cellColumn == null) return false;

        var position = source.Index;
        table.Rows.RemoveObjectAt(position);
        table.Rows.InsertObject(position, firstFragment);
        table.Rows.InsertObject(position + 1, continuation);
        for (var index = position; index < table.Rows.Count; index++)
        {
            var row = table.Rows[index];
            rowIndex.SetValue(row, index);
            rowTable.SetValue(row, table);
            for (var column = 0; column < row.Cells.Count; column++)
            {
                var cell = row.Cells[column];
                cellRow.SetValue(cell, row);
                cellTable.SetValue(cell, table);
                cellColumn.SetValue(cell, table.Columns[column]);
            }
        }
        return true;
    }
}
