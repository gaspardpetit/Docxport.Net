using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;

namespace DocxportNet.Visitors.Pdf;

internal static class DxpFootnoteTableRowSplitter
{
    internal static bool TrySplitAfterParagraphs(Row row, int firstFragmentParagraphs,
        out Row? firstFragment, out Row? continuation)
    {
        firstFragment = null;
        continuation = null;
        if (firstFragmentParagraphs < 1 || row.Cells.Count == 0 ||
            row.Cells.Cast<Cell>().Any(cell => cell.MergeRight > 0 || cell.MergeDown > 0 ||
                cell.Elements.Count <= firstFragmentParagraphs ||
                cell.Elements.Cast<DocumentObject>().Any(element => element is not Paragraph)))
            return false;

        var first = (Row)row.Clone();
        var second = (Row)row.Clone();
        for (var index = 0; index < row.Cells.Count; index++)
        {
            var firstCell = first.Cells[index];
            var secondCell = second.Cells[index];
            for (var elementIndex = firstCell.Elements.Count - 1;
                 elementIndex >= firstFragmentParagraphs; elementIndex--)
                firstCell.Elements.RemoveObjectAt(elementIndex);
            for (var elementIndex = firstFragmentParagraphs - 1; elementIndex >= 0; elementIndex--)
                secondCell.Elements.RemoveObjectAt(elementIndex);
            if (secondCell.Elements[0] is Paragraph firstContinuation)
                firstContinuation.Format.SpaceBefore = Unit.Zero;
            firstCell.Borders.Bottom.Visible = false;
            secondCell.Borders.Top.Visible = false;
        }
        first.Borders.Bottom.Visible = false;
        second.Borders.Top.Visible = false;
        second.HeadingFormat = false;
        firstFragment = first;
        continuation = second;
        return true;
    }

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
        var position = source.Index;
        table.Rows.RemoveObjectAt(position);
        table.Rows.InsertObject(position, firstFragment);
        table.Rows.InsertObject(position + 1, continuation);
        return true;
    }
}
