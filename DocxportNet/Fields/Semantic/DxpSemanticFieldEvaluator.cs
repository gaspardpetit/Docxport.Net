using DocxportNet.API;
using DocxportNet.Fields.Resolution;
using DocxportNet.Fields.Eval;
using System.Globalization;
using System.Text.RegularExpressions;

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

    internal async Task<DxpSemanticFieldResult> EvaluateExpressionAsync(
        DxpFieldExpression expression,
        DxpIDocumentContext? documentContext = null,
        CancellationToken cancellationToken = default)
    {
        DxpSemanticFieldResult result = await EvaluateExpressionAsync(
            expression, documentContext, cancellationToken, 0);
        return result with { Source = expression.Source };
    }

    private async Task<DxpSemanticFieldResult> EvaluateExpressionAsync(
        DxpFieldExpression expression,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken,
        int depth)
    {
        if (depth >= MaxDepth)
            return new DxpSemanticFieldResult(
                DxpFieldEvalStatus.Failed,
                DxpSemanticContent.Empty,
                Error: new InvalidOperationException("Maximum nested field depth exceeded."));

        var tokens = DxpFieldExpressionTokenizer.Tokenize(expression).ToList();
        if (tokens.Count == 0 || string.IsNullOrWhiteSpace(tokens[0].LiteralText))
            return DxpSemanticFieldResult.Empty(DxpFieldEvalStatus.Skipped);
        string fieldType = tokens[0].LiteralText!;

        if (fieldType.Equals("IF", StringComparison.OrdinalIgnoreCase))
            return await EvaluateIfExpressionAsync(tokens, documentContext, cancellationToken, depth);
        if (fieldType.Equals("SET", StringComparison.OrdinalIgnoreCase))
            return await EvaluateSetExpressionAsync(tokens, documentContext, cancellationToken, depth);
        if (fieldType.Equals("DATABASE", StringComparison.OrdinalIgnoreCase))
            return await EvaluateDatabaseAsync(
                await EvaluateInstructionTextAsync(expression, documentContext, cancellationToken, depth + 1),
                documentContext, cancellationToken, expression.CachedResult);
        if (fieldType.Equals("INCLUDETEXT", StringComparison.OrdinalIgnoreCase))
            return await EvaluateIncludeTextExpressionAsync(
                expression, tokens, documentContext, cancellationToken, depth + 1);

        // Word accepts a field containing only a bookmark name as the compact
        // equivalent of REF bookmark. Legacy templates use this form heavily,
        // including inside paths and conditions.
        if (tokens.Count == 1 && DxpFieldInstructionClassifier.TryGetImplicitRefName(
                fieldType, _eval.Context, out string bookmark))
            return await EvaluateScalarAsync(
                " REF " + bookmark + " ", documentContext, cancellationToken);

        return await EvaluateScalarExpressionAsync(expression, documentContext, cancellationToken);
    }

    private async Task<DxpSemanticFieldResult> EvaluateIfExpressionAsync(
        List<DxpFieldExpressionToken> tokens,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken,
        int depth)
    {
        NormalizeCompactComparison(tokens, operatorIndex: 2);
        if (tokens.Count < 5 || tokens[2].LiteralText == null)
            return new DxpSemanticFieldResult(
                DxpFieldEvalStatus.Failed,
                DxpSemanticContent.Empty,
                Error: new InvalidOperationException("IF field has invalid arguments."));

        DxpFieldValue left = await EvaluateTemplateValueAsync(tokens[1], documentContext, cancellationToken, depth + 1);
        DxpFieldValue right = await EvaluateTemplateValueAsync(tokens[3], documentContext, cancellationToken, depth + 1);
        bool condition = _eval.EvaluateSemanticComparison(left, tokens[2].LiteralText!, right);
        DxpFieldExpressionToken selected = condition
            ? tokens[4]
            : tokens.Count > 5 ? tokens[5] : new DxpFieldExpressionToken(Array.Empty<DxpFieldTemplatePart>());
        DxpSemanticContent content = await EvaluateTemplateContentAsync(
            selected, documentContext, cancellationToken, depth + 1);
        return new DxpSemanticFieldResult(DxpFieldEvalStatus.Resolved, content);
    }

    private async Task<DxpSemanticFieldResult> EvaluateSetExpressionAsync(
        IReadOnlyList<DxpFieldExpressionToken> tokens,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken,
        int depth)
    {
        if (tokens.Count < 2 || string.IsNullOrWhiteSpace(tokens[1].LiteralText))
            return new DxpSemanticFieldResult(
                DxpFieldEvalStatus.Failed,
                DxpSemanticContent.Empty,
                Error: new InvalidOperationException("SET field has invalid arguments."));

        DxpFieldValue value = tokens.Count > 2
            ? await EvaluateTemplateValueAsync(tokens[2], documentContext, cancellationToken, depth + 1)
            : new DxpFieldValue(string.Empty);
        string text = ValueToText(value);
        string bookmark = tokens[1].LiteralText!;
        _eval.Context.SetBookmarkValue(bookmark, value);
        _eval.Context.SetBookmarkNodes(bookmark, DxpFieldNodeBuffer.FromText(text));
        return DxpSemanticFieldResult.Empty();
    }

    private async Task<DxpSemanticFieldResult> EvaluateScalarExpressionAsync(
        DxpFieldExpression expression,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken)
    {
        return await EvaluateScalarAsync(
            await EvaluateInstructionTextAsync(expression, documentContext, cancellationToken, 1),
            documentContext,
            cancellationToken,
            expression.CachedResult);
    }

    private async Task<string> EvaluateInstructionTextAsync(
        DxpFieldExpression expression,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken,
        int depth)
    {
        var text = new System.Text.StringBuilder();
        foreach (DxpFieldExpressionPart part in expression.Parts)
        {
            switch (part)
            {
                case DxpFieldExpressionText literal:
                    text.Append(literal.Text);
                    break;
                case DxpFieldExpressionParagraph:
                    text.Append(' ');
                    break;
                case DxpFieldExpressionChild child:
                    DxpSemanticFieldResult result = await EvaluateExpressionAsync(
                        child.Expression, documentContext, cancellationToken, depth);
                    text.Append(result.Value.HasValue
                        ? ValueToText(result.Value.Value)
                        : ToPlainText(result.Content));
                    break;
            }
        }
        return text.ToString();
    }

    private async Task<DxpSemanticFieldResult> EvaluateIncludeTextExpressionAsync(
        DxpFieldExpression expression,
        IReadOnlyList<DxpFieldExpressionToken> tokens,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken,
        int depth)
    {
        if (tokens.Count is < 2 or > 3)
            return IncludeFallback(expression, error: false);
        string path = ValueToText(await EvaluateTemplateValueAsync(
            tokens[1], documentContext, cancellationToken, depth));
        string? bookmark = tokens.Count == 3
            ? ValueToText(await EvaluateTemplateValueAsync(
                tokens[2], documentContext, cancellationToken, depth))
            : null;
        if (string.IsNullOrWhiteSpace(path) ||
            bookmark != null && string.IsNullOrWhiteSpace(bookmark))
            return IncludeFallback(expression, error: false);

        IDxpIncludeTextResolver? resolver = _eval.Context.IncludeTextResolver;
        if (resolver == null)
            return IncludeFallback(expression, error: false);

        DxpIncludeTextSource? source;
        try
        {
            source = await resolver.ResolveAsync(
                new DxpIncludeTextRequest(path), _eval.Context, cancellationToken);
        }
        catch (Exception exception) when (_eval.Options.UseCacheOnError)
        {
            return IncludeFallback(expression, error: true, exception);
        }
        if (source == null || source.Content.Length == 0)
            return IncludeFallback(expression, error: false);

        byte[] content = source.Content;
        string rawText = string.Concat(expression.Parts
            .OfType<DxpFieldExpressionText>()
            .Select(static part => part.Text));
        bool isHtml = Regex.IsMatch(rawText, @"\\c\s+(?:\""?HTML\""?)", RegexOptions.IgnoreCase)
            || source.Format == DxpIncludeTextSourceFormat.Html
            || source.Format == DxpIncludeTextSourceFormat.Auto &&
                (path.EndsWith(".htm", StringComparison.OrdinalIgnoreCase) ||
                 path.EndsWith(".html", StringComparison.OrdinalIgnoreCase));
        if (isHtml)
        {
            try
            {
                content = await _eval.Context.ConvertHtmlIncludeAsync(content, cancellationToken);
            }
            catch (Exception exception) when (_eval.Options.UseCacheOnError)
            {
                return IncludeFallback(expression, error: true, exception);
            }
        }

        return new DxpSemanticFieldResult(
            DxpFieldEvalStatus.Resolved,
            new DxpSemanticContent(new DxpSemanticNode[]
            {
                new DxpSemanticInclude(path, source.Identity, content, bookmark)
            }));
    }

    private DxpSemanticFieldResult IncludeFallback(
        DxpFieldExpression expression,
        bool error,
        Exception? exception = null)
    {
        bool useCache = error ? _eval.Options.UseCacheOnError : _eval.Options.UseCacheOnNull;
        if (useCache && expression.CachedResult != null)
        {
            return new DxpSemanticFieldResult(
                DxpFieldEvalStatus.UsedCache,
                new DxpSemanticContent(new DxpSemanticNode[]
                {
                    new DxpSemanticText(expression.CachedResult)
                }),
                new DxpFieldValue(expression.CachedResult),
                exception,
                expression.Source);
        }
        return new DxpSemanticFieldResult(
            error ? DxpFieldEvalStatus.Failed : DxpFieldEvalStatus.Skipped,
            DxpSemanticContent.Empty,
            Error: exception,
            Source: expression.Source);
    }

    private async Task<DxpFieldValue> EvaluateTemplateValueAsync(
        DxpFieldExpressionToken template,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken,
        int depth)
    {
        if (template.Parts.Count == 1 && template.Parts[0] is DxpFieldTemplateChild child)
        {
            DxpSemanticFieldResult result = await EvaluateExpressionAsync(
                child.Expression, documentContext, cancellationToken, depth);
            return result.Value ?? new DxpFieldValue(ToPlainText(result.Content));
        }

        DxpSemanticContent content = await EvaluateTemplateContentAsync(
            template, documentContext, cancellationToken, depth);
        string text = ToPlainText(content);
        if (template.Parts.All(static part => part is DxpFieldTemplateText))
            return await _eval.ResolveSemanticValueAsync(text, documentContext);
        return new DxpFieldValue(text);
    }

    private async Task<DxpSemanticContent> EvaluateTemplateContentAsync(
        DxpFieldExpressionToken template,
        DxpIDocumentContext? documentContext,
        CancellationToken cancellationToken,
        int depth)
    {
        var root = new DxpSemanticContentBuilder();
        DxpSemanticContentBuilder? paragraph = null;
        DxpSemanticParagraphFormat? paragraphFormat = null;

        void FlushParagraph()
        {
            if (paragraph == null)
                return;
            root.Append(new DxpSemanticParagraph(paragraph.Build(), paragraphFormat));
            paragraph = null;
            paragraphFormat = null;
        }

        foreach (DxpFieldTemplatePart part in template.Parts)
        {
            if (part is DxpFieldTemplateParagraph paragraphBoundary)
            {
                FlushParagraph();
                paragraph = new DxpSemanticContentBuilder();
                paragraphFormat = paragraphBoundary.Format;
                continue;
            }
            if (part is DxpFieldTemplateText text)
            {
                (paragraph ?? root).AppendTextWithControls(text.Text, text.Format);
                continue;
            }

            var child = (DxpFieldTemplateChild)part;
            DxpSemanticFieldResult result = await EvaluateExpressionAsync(
                child.Expression, documentContext, cancellationToken, depth);
            if (paragraph != null && result.Content.HasBlocks)
                FlushParagraph();
            (paragraph ?? root).Append(result.Content);
        }
        FlushParagraph();
        return root.Build();
    }

    private static void NormalizeCompactComparison(List<DxpFieldExpressionToken> tokens, int operatorIndex)
    {
        if (tokens.Count <= operatorIndex || tokens[operatorIndex].LiteralText is not string compact)
            return;
        string? op = compact.StartsWith("<>", StringComparison.Ordinal) ||
            compact.StartsWith(">=", StringComparison.Ordinal) ||
            compact.StartsWith("<=", StringComparison.Ordinal)
                ? compact.Substring(0, 2)
                : compact.Length > 0 && compact[0] is '=' or '>' or '<'
                    ? compact.Substring(0, 1)
                    : null;
        if (op == null || compact.Length == op.Length)
            return;
        tokens[operatorIndex] = new DxpFieldExpressionToken(
            new DxpFieldTemplatePart[] { new DxpFieldTemplateText(op) });
        tokens.Insert(operatorIndex + 1, new DxpFieldExpressionToken(
            new DxpFieldTemplatePart[] { new DxpFieldTemplateText(compact.Substring(op.Length)) }));
    }

    private static string ToPlainText(DxpSemanticContent content)
    {
        var text = new System.Text.StringBuilder();
        foreach (DxpSemanticNode node in content.Nodes)
        {
            switch (node)
            {
                case DxpSemanticText value:
                    text.Append(value.Text);
                    break;
                case DxpSemanticBreak:
                    text.AppendLine();
                    break;
                case DxpSemanticTab:
                    text.Append('\t');
                    break;
                case DxpSemanticParagraph paragraph:
                    if (text.Length > 0)
                        text.AppendLine();
                    text.Append(ToPlainText(paragraph.Content));
                    break;
                case DxpSemanticTable table:
                    foreach (DxpSemanticTableRow row in table.Rows)
                    {
                        if (text.Length > 0)
                            text.AppendLine();
                        text.Append(string.Join("\t", row.Cells.Select(static cell => ToPlainText(cell.Content))));
                    }
                    break;
                case DxpSemanticInclude:
                    break;
            }
        }
        return text.ToString();
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
        CancellationToken cancellationToken,
        string? cachedResult = null)
    {
        var execution = await _eval.ExecuteDatabaseAsync(
            new DxpFieldInstruction(instructionText, cachedResult), documentContext, cancellationToken);
        if (execution?.Result == null)
            return await EvaluateScalarAsync(
                instructionText, documentContext, cancellationToken, cachedResult);

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
        CancellationToken cancellationToken,
        string? cachedResult = null)
    {
        DxpFieldEvalResult result = await _eval.EvalAsync(
            new DxpFieldInstruction(instructionText, cachedResult), documentContext, cancellationToken);
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
