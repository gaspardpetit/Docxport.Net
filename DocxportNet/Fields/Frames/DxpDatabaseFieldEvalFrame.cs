using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields.Resolution;
using DocxportNet.Middleware;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpDatabaseFieldEvalFrame : DxpMiddleware, DxpIFieldEvalFrame
{
    private readonly DxpFieldEval _eval;
    private readonly string _instructionText;
    private bool _evaluated;

    public override DxpIVisitor Next { get; }

    public DxpDatabaseFieldEvalFrame(
        DxpIVisitor next,
        DxpFieldEval eval,
        string instructionText)
    {
        Next = next;
        _eval = eval;
        _instructionText = instructionText;
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d) => Evaluate(d);

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
        => DxpDisposable.Create(() => Evaluate(d));

    protected override bool ShouldForwardContent(DxpIDocumentContext d) => false;

    private void Evaluate(DxpIDocumentContext context)
    {
        if (_evaluated)
            return;
        _evaluated = true;

        var execution = _eval.ExecuteDatabaseAsync(
            new DxpFieldInstruction(_instructionText), context).GetAwaiter().GetResult();
        if (execution?.Result == null ||
            (execution.Value.Result.Rows.Count == 0 &&
             (!execution.Value.Request.IncludeColumnHeadings || execution.Value.Result.Columns.Count == 0)))
            return;

        OpenXmlElement block = BuildResult(execution.Value.Request, execution.Value.Result);
        var buffer = DxpFieldNodeBuffer.FromBlock(block);
        if (_eval.Context.StructuredFieldSpliceCollector?.Record(buffer) == true)
            return;
        if (Next is IDxpStructuredFieldResultSink sink && sink.TryRecordStructuredFieldResult(buffer))
            return;
        _eval.Context.DeferStructuredFieldResult(context, buffer);
    }

    private OpenXmlElement BuildResult(DxpDatabaseRequest request, DxpDatabaseResult result)
    {
        int columnCount = result.Columns.Count > 0
            ? result.Columns.Count
            : result.Rows.Select(row => row.Count).DefaultIfEmpty().Max();
        if (columnCount >= 62)
            return BuildTabbedResult(request, result, columnCount);

        var properties = new TableProperties(
            new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto });
        if ((request.TableFormatAttributes.GetValueOrDefault() & 16) != 0)
            properties.AppendChild(new TableLayout { Type = TableLayoutValues.Autofit });
        if ((request.TableFormatAttributes.GetValueOrDefault() & 1) != 0)
        {
            properties.AppendChild(new TableBorders(
                NewBorder<TopBorder>(), NewBorder<LeftBorder>(),
                NewBorder<BottomBorder>(), NewBorder<RightBorder>(),
                NewBorder<InsideHorizontalBorder>(), NewBorder<InsideVerticalBorder>()));
        }

        var table = new Table(properties, new TableGrid(
            Enumerable.Range(0, columnCount).Select(_ => new GridColumn())));
        if (request.IncludeColumnHeadings && result.Columns.Count > 0)
        {
            var header = BuildRow(result.Columns.Select(column => column.Name), columnCount);
            header.TableRowProperties = new TableRowProperties(new TableHeader());
            table.AppendChild(header);
        }
        foreach (var row in DxpFieldEval.SelectDatabaseRows(result.Rows, request))
            table.AppendChild(BuildRow(row.Select(_eval.FormatDatabaseCellValue), columnCount));
        return table;
    }

    private OpenXmlElement BuildTabbedResult(
        DxpDatabaseRequest request,
        DxpDatabaseResult result,
        int columnCount)
    {
        var paragraphs = new List<Paragraph>();
        if (request.IncludeColumnHeadings && result.Columns.Count > 0)
            paragraphs.Add(BuildTabbedParagraph(result.Columns.Select(column => column.Name), columnCount));
        foreach (var row in DxpFieldEval.SelectDatabaseRows(result.Rows, request))
            paragraphs.Add(BuildTabbedParagraph(row.Select(_eval.FormatDatabaseCellValue), columnCount));

        // Word uses tab-separated columns rather than a table at 62+ columns.
        // A block buffer currently carries one root, so preserve the rows in a
        // single paragraph separated by line breaks.
        var combined = new Paragraph();
        for (int i = 0; i < paragraphs.Count; i++)
        {
            if (i > 0)
                combined.AppendChild(new Run(new Break()));
            foreach (var child in paragraphs[i].ChildElements)
                combined.AppendChild(child.CloneNode(true));
        }
        return combined;
    }

    private TableRow BuildRow(IEnumerable<string> values, int columnCount)
    {
        var cells = values.Take(columnCount).Select(value =>
            new TableCell(new Paragraph(BuildValueRun(value)))).ToList();
        while (cells.Count < columnCount)
            cells.Add(new TableCell(new Paragraph(new Run(new Text()))));
        return new TableRow(cells);
    }

    private static Paragraph BuildTabbedParagraph(IEnumerable<string> values, int columnCount)
    {
        var paragraph = new Paragraph();
        int index = 0;
        foreach (string value in values.Take(columnCount))
        {
            if (index++ > 0)
                paragraph.AppendChild(new Run(new TabChar()));
            paragraph.AppendChild(BuildValueRun(value));
        }
        return paragraph;
    }

    private static T NewBorder<T>() where T : BorderType, new()
        => new() { Val = BorderValues.Single, Size = 4U };

    private static Text NewText(string? value)
    {
        string text = value ?? string.Empty;
        return new Text(text) {
            Space = text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]))
                ? SpaceProcessingModeValues.Preserve
                : null
        };
    }

    private static Run BuildValueRun(string? value)
    {
        string text = value ?? string.Empty;
        var run = new Run();
        int segmentStart = 0;
        for (int index = 0; index < text.Length; index++)
        {
            char ch = text[index];
            if (ch is not ('\r' or '\n' or '\t'))
                continue;
            if (index > segmentStart)
                run.AppendChild(NewText(text.Substring(segmentStart, index - segmentStart)));
            if (ch == '\t')
                run.AppendChild(new TabChar());
            else
            {
                run.AppendChild(new Break());
                if (ch == '\r' && index + 1 < text.Length && text[index + 1] == '\n')
                    index++;
            }
            segmentStart = index + 1;
        }
        if (segmentStart < text.Length)
            run.AppendChild(NewText(text.Substring(segmentStart)));
        else if (text.Length == 0)
            run.AppendChild(new Text());
        return run;
    }
}
