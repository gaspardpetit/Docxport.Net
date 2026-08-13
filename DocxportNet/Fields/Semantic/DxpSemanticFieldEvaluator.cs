using DocxportNet.API;
using DocxportNet.Fields.Resolution;
using DocxportNet.Fields.Eval;
using System.Globalization;

namespace DocxportNet.Fields.Semantic;

internal sealed class DxpSemanticFieldEvaluator
{
    private const int MaxDepth = 32;
    private readonly DxpFieldEval _eval;
    private readonly DxpFieldParser _parser = new();

    public DxpSemanticFieldEvaluator(DxpFieldEval eval) => _eval = eval;

    public Task<DxpSemanticFieldResult> EvaluateAsync(
        string instructionText,
        DxpIDocumentContext? documentContext = null,
        CancellationToken cancellationToken = default)
        => EvaluateFieldAsync(instructionText, documentContext, cancellationToken, 0);

    internal async Task<DxpSemanticFieldResult> EvaluateBranchAsync(
        IReadOnlyList<DxpSemanticBranchPart> parts,
        DxpIDocumentContext? documentContext = null,
        CancellationToken cancellationToken = default)
    {
        var content = new DxpSemanticContentBuilder();
        DxpSemanticContentBuilder? paragraph = null;

        void FlushParagraph()
        {
            if (paragraph == null)
                return;
            content.Append(new DxpSemanticParagraph(paragraph.Build()));
            paragraph = null;
        }

        foreach (DxpSemanticBranchPart part in parts)
        {
            if (part is DxpSemanticBranchParagraphStart)
            {
                FlushParagraph();
                paragraph = new DxpSemanticContentBuilder();
                continue;
            }
            if (part is DxpSemanticBranchText text)
            {
                (paragraph ?? content).AppendTextWithControls(text.Text);
                continue;
            }

            var field = ((DxpSemanticBranchField)part).Field;
            DxpSemanticFieldResult nested = await EvaluateDeferredAsync(
                field, documentContext, cancellationToken);
            if (nested.Status == DxpFieldEvalStatus.Failed)
                return nested;
            if (paragraph != null && nested.Content.HasBlocks)
                FlushParagraph();
            (paragraph ?? content).Append(nested.Content);
        }
        FlushParagraph();
        return new DxpSemanticFieldResult(DxpFieldEvalStatus.Resolved, content.Build());
    }

    private async Task<DxpSemanticFieldResult> EvaluateDeferredAsync(
        DxpDeferredField field,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken)
    {
        var parse = _parser.Parse(field.InstructionText);
        if (parse.Ast.FieldType?.Equals("SET", StringComparison.OrdinalIgnoreCase) == true &&
            field.CapturedScalar.HasValue &&
            DxpFieldTokenization.TryGetFirstToken(parse.Ast.ArgumentsText, out string bookmark))
        {
            DxpFieldValue value = field.CapturedScalar.Value;
            string text = ValueToText(value);
            _eval.Context.SetBookmarkValue(bookmark, value);
            _eval.Context.SetBookmarkNodes(bookmark, DxpFieldNodeBuffer.FromText(text));
            return DxpSemanticFieldResult.Empty();
        }

        return await EvaluateFieldAsync(
            field.InstructionText, documentContext, cancellationToken, 0);
    }

    private async Task<DxpSemanticFieldResult> EvaluateFieldAsync(
        string instructionText,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken,
        int depth)
    {
        if (depth >= MaxDepth)
            return new DxpSemanticFieldResult(
                DxpFieldEvalStatus.Failed,
                DxpSemanticContent.Empty,
                Error: new InvalidOperationException("Maximum nested field depth exceeded."));

        var parse = _parser.Parse(instructionText);
        string? fieldType = parse.Ast.FieldType;
        if (fieldType?.Equals("IF", StringComparison.OrdinalIgnoreCase) == true)
        {
            var condition = await _eval.EvaluateIfConditionAsync(instructionText, documentContext);
            if (condition == null || !condition.Value.Success)
                return await EvaluateScalarAsync(instructionText, documentContext, cancellationToken);
            string branch = condition.Value.Condition
                ? condition.Value.TrueText ?? string.Empty
                : condition.Value.FalseText ?? string.Empty;
            return await EvaluateTemplateAsync(branch, documentContext, cancellationToken, depth + 1);
        }

        if (fieldType?.Equals("DATABASE", StringComparison.OrdinalIgnoreCase) == true)
            return await EvaluateDatabaseAsync(instructionText, documentContext, cancellationToken);

        if (fieldType?.Equals("SET", StringComparison.OrdinalIgnoreCase) == true)
        {
            DxpFieldEvalResult set = await _eval.EvalAsync(
                new DxpFieldInstruction(instructionText), documentContext, cancellationToken);
            return DxpSemanticFieldResult.Empty(set.Status);
        }

        return await EvaluateScalarAsync(instructionText, documentContext, cancellationToken);
    }

    private async Task<DxpSemanticFieldResult> EvaluateTemplateAsync(
        string template,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken,
        int depth)
    {
        var content = new DxpSemanticContentBuilder();
        int literalStart = 0;
        for (int index = 0; index < template.Length; index++)
        {
            if (template[index] != '{')
                continue;
            if (!TryReadNestedField(template, index, out int end, out string instruction))
                continue;

            content.AppendTextWithControls(template.Substring(literalStart, index - literalStart));
            var nested = await EvaluateFieldAsync(instruction, documentContext, cancellationToken, depth);
            if (nested.Status == DxpFieldEvalStatus.Failed)
                return nested;
            content.Append(nested.Content);
            index = end;
            literalStart = end + 1;
        }
        content.AppendTextWithControls(template.Substring(literalStart));
        return new DxpSemanticFieldResult(DxpFieldEvalStatus.Resolved, content.Build());
    }

    private async Task<DxpSemanticFieldResult> EvaluateDatabaseAsync(
        string instructionText,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken)
    {
        var execution = await _eval.ExecuteDatabaseAsync(
            new DxpFieldInstruction(instructionText), documentContext, cancellationToken);
        if (execution?.Result == null)
            return DxpSemanticFieldResult.Empty(DxpFieldEvalStatus.Skipped);

        DxpDatabaseRequest request = execution.Value.Request;
        DxpDatabaseResult result = execution.Value.Result;
        int columnCount = result.Columns.Count > 0
            ? result.Columns.Count
            : result.Rows.Select(static row => row.Count).DefaultIfEmpty().Max();
        if (columnCount == 0)
            return DxpSemanticFieldResult.Empty();

        var rows = new List<DxpSemanticTableRow>();
        if (request.IncludeColumnHeadings && result.Columns.Count > 0)
            rows.Add(BuildRow(result.Columns.Select(static column => column.Name), columnCount));
        foreach (var row in DxpFieldEval.SelectDatabaseRows(result.Rows, request))
            rows.Add(BuildRow(row.Select(_eval.FormatDatabaseCellValue), columnCount));
        if (rows.Count == 0)
            return DxpSemanticFieldResult.Empty();

        var table = new DxpSemanticTable(
            rows,
            HasHeader: request.IncludeColumnHeadings && result.Columns.Count > 0,
            ShowBorders: (request.TableFormatAttributes.GetValueOrDefault() & 1) != 0,
            AutoFit: (request.TableFormatAttributes.GetValueOrDefault() & 16) != 0);
        return new DxpSemanticFieldResult(
            DxpFieldEvalStatus.Resolved,
            new DxpSemanticContent(new DxpSemanticNode[] { table }));
    }

    private static DxpSemanticTableRow BuildRow(IEnumerable<string> values, int columnCount)
    {
        var cells = values.Take(columnCount)
            .Select(value => new DxpSemanticTableCell(ContentFromText(value)))
            .ToList();
        while (cells.Count < columnCount)
            cells.Add(new DxpSemanticTableCell(DxpSemanticContent.Empty));
        return new DxpSemanticTableRow(cells);
    }

    private static DxpSemanticContent ContentFromText(string? text)
    {
        var builder = new DxpSemanticContentBuilder();
        builder.AppendTextWithControls(text);
        return builder.Build();
    }

    private async Task<DxpSemanticFieldResult> EvaluateScalarAsync(
        string instructionText,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken)
    {
        DxpFieldEvalResult result = await _eval.EvalAsync(
            new DxpFieldInstruction(instructionText), documentContext, cancellationToken);
        var builder = new DxpSemanticContentBuilder();
        if (result.Text != null)
            builder.AppendTextWithControls(result.Text);
        return new DxpSemanticFieldResult(result.Status, builder.Build(), result.Value, result.Error);
    }

    private static bool TryReadNestedField(
        string text,
        int start,
        out int end,
        out string instruction)
    {
        int depth = 0;
        bool inQuote = false;
        for (int index = start; index < text.Length; index++)
        {
            char ch = text[index];
            if (ch == '"' && (index == 0 || text[index - 1] != '\\'))
                inQuote = !inQuote;
            if (inQuote)
                continue;
            if (ch == '{')
                depth++;
            else if (ch == '}' && --depth == 0)
            {
                end = index;
                instruction = text.Substring(start + 1, index - start - 1).Trim();
                return instruction.Length > 0;
            }
        }
        end = start;
        instruction = string.Empty;
        return false;
    }

    private string ValueToText(DxpFieldValue value)
        => value.Kind switch
        {
            DxpFieldValueKind.Number => value.NumberValue.GetValueOrDefault().ToString(
                _eval.Context.Culture ?? CultureInfo.CurrentCulture),
            DxpFieldValueKind.DateTime => value.DateTimeValue.GetValueOrDefault().ToString(
                _eval.Context.Culture ?? CultureInfo.CurrentCulture),
            _ => value.StringValue ?? string.Empty
        };
}
