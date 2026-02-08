using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;
using Microsoft.Extensions.Logging;
using System.Globalization;
using System.Text;

namespace DocxportNet.Walker;

public sealed class DxpFieldEvalVisitorMiddleware : DxpVisitorMiddlewareBase
{
    private readonly DxpFieldEvalContext _context;
    private readonly bool _includeDocumentProperties;
    private readonly bool _includeCustomProperties;
    private readonly Func<DateTimeOffset>? _nowProvider;
    private readonly ILogger? _logger;
    private bool _initialized;
    private int _paragraphOrder;

    public DxpFieldEvalVisitorMiddleware(
        DxpIVisitor next,
        DxpFieldEvalContext context,
        bool includeDocumentProperties = true,
        bool includeCustomProperties = false,
        Func<DateTimeOffset>? nowProvider = null,
        ILogger? logger = null)
        : base(next)
    {
        _context = context ?? throw new ArgumentNullException(nameof(context));
        _includeDocumentProperties = includeDocumentProperties;
        _includeCustomProperties = includeCustomProperties;
        _nowProvider = nowProvider;
        _logger = logger;
    }

    public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext documentContext)
    {
        if (!_initialized)
        {
            _context.InitFromDocumentContext(documentContext, _includeDocumentProperties, _includeCustomProperties);
            if (_nowProvider != null)
                _context.SetNow(_nowProvider);
            _context.TableResolver ??= new DxpWalkerTableResolver(documentContext);
            _context.RefResolver ??= new DocxportNet.Fields.Resolution.DxpRefIndexResolver(
                documentContext.DocumentIndex.RefIndex,
                () => _context.CurrentDocumentOrder);
            _initialized = true;
        }

        _paragraphOrder = 0;
        return _next.VisitDocumentBegin(doc, documentContext);
    }

    protected override bool ShouldForwardContent(DxpIDocumentContext d)
    {
        var fields = d.CurrentFields;
        if (!fields.IsInFieldResult)
            return true;
        return !fields.IsNestedField;
    }


    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        var previous = _context.Culture;
        var previousOutlineProvider = _context.CurrentOutlineLevelProvider;
        var previousOrder = _context.CurrentDocumentOrder;
        if (TryResolveParagraphCulture(p, d, _logger, out var culture))
            _context.Culture = culture;
        _context.CurrentOutlineLevelProvider = CreateOutlineLevelProvider(p, d);
        _context.CurrentDocumentOrder = ++_paragraphOrder;

        var inner = _next.VisitParagraphBegin(p, d, paragraph);
        return new DxpCompositeScope(inner, () => {
            _context.Culture = previous;
            _context.CurrentOutlineLevelProvider = previousOutlineProvider;
            _context.CurrentDocumentOrder = previousOrder;
        });
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        var previous = _context.Culture;
        if (TryResolveRunCulture(r, d, _logger, out var culture))
            _context.Culture = culture;

        var inner = _next.VisitRunBegin(r, d);
        return new DxpCompositeScope(inner, () => _context.Culture = previous);
    }

    private static bool TryResolveParagraphCulture(Paragraph p, DxpIDocumentContext d, ILogger? logger, out CultureInfo culture)
    {
        culture = CultureInfo.CurrentCulture;
        string? lang = null;

        if (d.Styles is DxpStyleResolver resolver)
            lang = resolver.ResolveParagraphLanguage(p) ?? resolver.GetDefaultLanguage();
        else
            lang = p.ParagraphProperties?.GetFirstChild<ParagraphMarkRunProperties>()
                ?.GetFirstChild<Languages>()?.Val?.Value;

        return TryCreateCulture(lang, logger, out culture);
    }

    private static bool TryResolveRunCulture(Run r, DxpIDocumentContext d, ILogger? logger, out CultureInfo culture)
    {
        culture = CultureInfo.CurrentCulture;
        string? lang = null;

        if (d.Styles is DxpStyleResolver resolver)
        {
            var paragraph = r.Ancestors<Paragraph>().FirstOrDefault();
            if (paragraph != null)
                lang = resolver.ResolveRunLanguage(paragraph, r);
        }

        lang ??= d.CurrentRun?.Language ?? r.RunProperties?.Languages?.Val?.Value;
        return TryCreateCulture(lang, logger, out culture);
    }

    private static bool TryCreateCulture(string? lang, ILogger? logger, out CultureInfo culture)
    {
        culture = CultureInfo.CurrentCulture;
        if (string.IsNullOrWhiteSpace(lang))
            return false;

        try
        {
            culture = new CultureInfo(lang);
            return true;
        }
        catch (CultureNotFoundException)
        {
            logger?.LogWarning("Invalid language tag '{Lang}' in document; using current culture.", lang);
            return false;
        }
    }

    private static Func<int> CreateOutlineLevelProvider(Paragraph p, DxpIDocumentContext d)
    {
        int? level = null;
        if (d.Styles is DxpStyleResolver resolver)
            level = resolver.GetOutlineLevel(p);
        else
            level = p.ParagraphProperties?.OutlineLevel?.Val?.Value;

        // Word stores outline levels as 0-based; SEQ \s expects 1-based.
        int resolved = level.HasValue ? level.Value + 1 : 0;
        return () => resolved;
    }

    private sealed class DxpCompositeScope : IDisposable
    {
        private readonly IDisposable _inner;
        private readonly Action _onDispose;
        private bool _disposed;

        public DxpCompositeScope(IDisposable inner, Action onDispose)
        {
            _inner = inner;
            _onDispose = onDispose;
        }

        public void Dispose()
        {
            if (_disposed)
                return;
            _disposed = true;
            _onDispose();
            _inner.Dispose();
        }
    }

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
