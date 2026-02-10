using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;
using System.Globalization;
using System.Text;

namespace DocxportNet.Walker;

public sealed partial class DxpFieldEvalMiddleware
{
    private sealed class DxpWalkerTableResolver : IDxpTableValueResolver
    {
        private readonly DxpIDocumentContext _document;

        public DxpWalkerTableResolver(DxpIDocumentContext document)
        {
            _document = document;
        }

        public Task<IReadOnlyList<double>> ResolveRangeAsync(string range, DxpFieldEvalContext context)
        {
            var model = _document.CurrentTableModel;
            if (model == null)
                return Task.FromResult<IReadOnlyList<double>>([]);

            if (!TryParseRange(range, out var startRow, out var startCol, out var endRow, out var endCol))
                return Task.FromResult<IReadOnlyList<double>>([]);

            var values = CollectRangeValues(model, startRow, startCol, endRow, endCol, context);
            return Task.FromResult<IReadOnlyList<double>>(values);
        }

        public Task<IReadOnlyList<double>> ResolveDirectionalRangeAsync(DxpTableRangeDirection direction, DxpFieldEvalContext context)
        {
            var model = _document.CurrentTableModel;
            var cell = _document.CurrentTableCell;
            if (model == null || cell == null)
                return Task.FromResult<IReadOnlyList<double>>([]);

            int row = cell.RowIndex;
            int col = cell.ColumnIndex;
            int startRow, endRow, startCol, endCol;

            switch (direction)
            {
                case DxpTableRangeDirection.Above:
                    startRow = 0;
                    endRow = row - 1;
                    startCol = col;
                    endCol = col;
                    break;
                case DxpTableRangeDirection.Below:
                    startRow = row + 1;
                    endRow = model.RowCount - 1;
                    startCol = col;
                    endCol = col;
                    break;
                case DxpTableRangeDirection.Left:
                    startRow = row;
                    endRow = row;
                    startCol = 0;
                    endCol = col - 1;
                    break;
                case DxpTableRangeDirection.Right:
                    startRow = row;
                    endRow = row;
                    startCol = col + 1;
                    endCol = model.ColumnCount - 1;
                    break;
                default:
                    return Task.FromResult<IReadOnlyList<double>>([]);
            }

            var values = CollectRangeValues(model, startRow, startCol, endRow, endCol, context);
            return Task.FromResult<IReadOnlyList<double>>(values);
        }

        private static List<double> CollectRangeValues(DxpTableModel model, int startRow, int startCol, int endRow, int endCol, DxpFieldEvalContext context)
        {
            var values = new List<double>();
            if (startRow > endRow || startCol > endCol)
                return values;

            startRow = Math.Max(0, startRow);
            startCol = Math.Max(0, startCol);
            endRow = Math.Min(model.RowCount - 1, endRow);
            endCol = Math.Min(model.ColumnCount - 1, endCol);

            for (int r = startRow; r <= endRow; r++)
            {
                for (int c = startCol; c <= endCol; c++)
                {
                    var cell = model.Cells[r, c];
                    if (cell == null || cell.IsCovered)
                        continue;

                    string text = ExtractCellText(cell.Cell);
                    if (TryParseNumber(text, context, out var number))
                        values.Add(number);
                }
            }

            return values;
        }

        private static bool TryParseRange(string range, out int startRow, out int startCol, out int endRow, out int endCol)
        {
            startRow = startCol = endRow = endCol = 0;
            if (string.IsNullOrWhiteSpace(range))
                return false;

            var parts = range.Split(':');
            if (parts.Length == 0 || parts.Length > 2)
                return false;

            if (!TryParseCell(parts[0], out startRow, out startCol))
                return false;

            if (parts.Length == 2)
            {
                if (!TryParseCell(parts[1], out endRow, out endCol))
                    return false;
            }
            else
            {
                endRow = startRow;
                endCol = startCol;
            }

            if (startRow > endRow)
                (startRow, endRow) = (endRow, startRow);
            if (startCol > endCol)
                (startCol, endCol) = (endCol, startCol);

            return true;
        }

        private static bool TryParseCell(string text, out int rowIndex, out int colIndex)
        {
            rowIndex = colIndex = 0;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            int i = 0;
            while (i < text.Length && char.IsLetter(text[i]))
                i++;
            if (i == 0 || i == text.Length)
                return false;

            var colPart = text.Substring(0, i).ToUpperInvariant();
            var rowPart = text.Substring(i);
            if (!int.TryParse(rowPart, NumberStyles.Integer, CultureInfo.InvariantCulture, out var row))
                return false;

            colIndex = ColumnLettersToIndex(colPart) - 1;
            rowIndex = row - 1;
            return rowIndex >= 0 && colIndex >= 0;
        }

        private static int ColumnLettersToIndex(string letters)
        {
            int value = 0;
            foreach (char ch in letters)
            {
                if (ch < 'A' || ch > 'Z')
                    return 0;
                value = value * 26 + (ch - 'A' + 1);
            }
            return value;
        }

        private static string ExtractCellText(TableCell cell)
        {
            var sb = new StringBuilder();
            foreach (var text in cell.Descendants<Text>())
                sb.Append(text.Text);
            return sb.ToString().Trim();
        }

        private static bool TryParseNumber(string text, DxpFieldEvalContext context, out double number)
        {
            var culture = context.Culture ?? CultureInfo.CurrentCulture;
            if (double.TryParse(text, NumberStyles.Any, culture, out number))
                return true;
            if (context.AllowInvariantNumericFallback &&
                double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out number))
                return true;
            number = 0;
            return false;
        }
    }
}
