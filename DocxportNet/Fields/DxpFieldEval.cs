using Microsoft.Extensions.Logging;
using System.Globalization;

namespace DocxportNet.Fields;

public sealed class DxpFieldEval
{
    private const string RefNotFoundError = "Error! Reference source not found.";
    private const string DocVariableMissingError = "Error! No document variable supplied.";
    private const string DocPropertyUnknownError = "Error! Unknown document property name.";
    private readonly DxpFieldParser _parser = new();
    private readonly DxpFieldFormatter _formatter = new();
    private readonly DxpFieldEvalDelegates _delegates;
    private readonly DxpFieldEvalOptions _options;
    private readonly Resolution.IDxpFieldValueResolver _resolver;
    private readonly ILogger? _logger;

    public DxpFieldEvalContext Context { get; } = new();

    public DxpFieldEval(DxpFieldEvalDelegates? delegates = null, DxpFieldEvalOptions? options = null, ILogger? logger = null)
    {
        _delegates = delegates ?? new DxpFieldEvalDelegates();
        _options = options ?? new DxpFieldEvalOptions();
        _logger = logger;
        _resolver = new Resolution.DxpChainedFieldValueResolver(
            new Resolution.DxpContextFieldValueResolver(),
            new Resolution.DxpDelegateFieldValueResolver(_delegates)
        );
    }

    public async Task<DxpFieldEvalResult> EvalAsync(DxpFieldInstruction instruction, CancellationToken cancellationToken = default)
    {
        _ = cancellationToken;
        _ = _delegates;
        try
        {
            if (_logger?.IsEnabled(LogLevel.Debug) == true)
                _logger.LogDebug("Evaluating field instruction '{Instruction}'.", instruction.InstructionText);

            var parse = _parser.Parse(instruction.InstructionText);
            if (!parse.Success)
            {
                _logger?.LogError("Field parse failed for instruction '{Instruction}'.", instruction.InstructionText);
                if (_options.ErrorOnUnsupported)
                {
                    return new DxpFieldEvalResult(
                        DxpFieldEvalStatus.Failed,
                        null,
                        new InvalidOperationException("Field parse failed."));
                }

                return FallbackToCacheOrSkip(instruction);
            }

            if (parse.Ast.FieldType != null)
            {
                string fieldType = parse.Ast.FieldType;
                bool knownFieldType = IsKnownFieldType(fieldType);
                if (fieldType.Equals("IF", StringComparison.OrdinalIgnoreCase))
                {
                    var ifResult = await EvalIfAsync(parse.Ast);
                    if (ifResult != null)
                        return ifResult;
                }
                else if (fieldType.Equals("COMPARE", StringComparison.OrdinalIgnoreCase))
                {
                    var compareResult = await EvalCompareAsync(parse.Ast);
                    if (compareResult != null)
                        return compareResult;
                }
                else if (fieldType.Equals("SKIPIF", StringComparison.OrdinalIgnoreCase) ||
                    fieldType.Equals("NEXTIF", StringComparison.OrdinalIgnoreCase))
                {
                    var skipResult = await EvalSkipIfAsync(parse.Ast);
                    if (skipResult != null)
                        return skipResult;
                }
                else
                {
                    var builtIn = await TryEvalBuiltInAsync(fieldType, parse.Ast);
                    if (builtIn.success)
                    {
                        string text = fieldType.Equals("MERGEFIELD", StringComparison.OrdinalIgnoreCase)
                            ? FormatMergeField(builtIn.value, parse.Ast)
                            : _formatter.Format(builtIn.value, parse.Ast.FormatSpecs, Context);
                        _logger?.LogInformation("Resolved field '{FieldType}'.", fieldType);
                        return new DxpFieldEvalResult(DxpFieldEvalStatus.Resolved, text);
                    }
                    if (knownFieldType)
                    {
                        _logger?.LogInformation("Field '{FieldType}' not resolved; skipping.", fieldType);
                    }
                }
            }
            else
            {
                _logger?.LogWarning("Field instruction missing field type; skipping. Instruction='{Instruction}'.", instruction.InstructionText);
            }

            if (_options.ErrorOnUnsupported)
            {
                _logger?.LogError("Field evaluation not implemented for instruction '{Instruction}'.", instruction.InstructionText);
                return new DxpFieldEvalResult(
                    DxpFieldEvalStatus.Failed,
                    null,
                    new NotSupportedException("Field evaluation is not implemented yet."));
            }

            if (parse.Ast.FieldType != null && !IsKnownFieldType(parse.Ast.FieldType))
                _logger?.LogWarning("Unsupported field type '{FieldType}'; skipping.", parse.Ast.FieldType);
            return FallbackToCacheOrSkip(instruction);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Field evaluation failed for instruction '{Instruction}'.", instruction.InstructionText);
            throw;
        }
    }

    private async Task<(bool success, DxpFieldValue value)> TryEvalBuiltInAsync(string fieldType, DxpFieldAst ast)
    {
        var value = default(DxpFieldValue);
        switch (fieldType.ToUpperInvariant())
        {
            case "=":
                if (ast.ArgumentsText != null)
                {
                    char listSeparator = ResolveListSeparator();
                    var parser = new Expressions.DxpFormulaParser(ast.ArgumentsText, listSeparator);
                    var expr = parser.ParseExpression();
                    var evaluator = new Expressions.DxpFormulaEvaluator(Context, EvalNestedFieldAsync, ResolveIdentifierValueAsync);
                    double result = await evaluator.EvaluateAsync(expr);
                    value = new DxpFieldValue(result);
                    if (_logger?.IsEnabled(LogLevel.Debug) == true)
                        _logger.LogDebug("Formula field evaluated to {Result}.", result);
                    return (true, value);
                }
                _logger?.LogWarning("Formula field '=' missing arguments.");
                return (false, value);
            case "DATE":
                value = new DxpFieldValue(Context.NowProvider());
                return (true, value);
            case "TIME":
                value = new DxpFieldValue(Context.NowProvider());
                return (true, value);
            case "CREATEDATE":
                if (Context.CreatedDate != null)
                {
                    value = new DxpFieldValue(Context.CreatedDate.Value);
                    return (true, value);
                }
                if (_logger?.IsEnabled(LogLevel.Debug) == true)
                    _logger.LogDebug("CREATEDATE not available in context.");
                return (false, value);
            case "SAVEDATE":
                if (Context.SavedDate != null)
                {
                    value = new DxpFieldValue(Context.SavedDate.Value);
                    return (true, value);
                }
                if (_logger?.IsEnabled(LogLevel.Debug) == true)
                    _logger.LogDebug("SAVEDATE not available in context.");
                return (false, value);
            case "PRINTDATE":
                value = new DxpFieldValue(Context.PrintDate ?? Context.NowProvider());
                return (true, value);
            case "SET":
                if (ast.ArgumentsText != null)
                {
                    var tokens = TokenizeArgs(ast.ArgumentsText);
                    if (tokens.Count > 0)
                    {
                        string name = tokens[0];
                        string rawValue = tokens.Count > 1 ? tokens[1] : string.Empty;
                        string resolved = await ResolveValueAsync(rawValue);
                        Context.SetBookmarkNodes(name, DxpFieldNodeBuffer.FromText(resolved));
                        value = new DxpFieldValue(resolved);
                        return (true, value);
                    }
                }
                _logger?.LogWarning("SET field missing arguments.");
                return (false, value);
            case "REF":
                if (ast.ArgumentsText != null)
                {
                    var tokens = TokenizeArgs(ast.ArgumentsText);
                    if (tokens.Count > 0)
                    {
                        string bookmark = tokens[0];
                        bookmark = await ExpandNestedTextAsync(bookmark);
                        var switches = ParseNonFormatSwitches(ast.RawText);
                        bool hasRefSwitches = HasRefSwitches(switches);
                        if (Context.RefResolver != null)
                        {
                            var request = new Resolution.DxpRefRequest(
                                Bookmark: bookmark,
                                Separator: switches.ContainsKey('d'),
                                Footnote: switches.ContainsKey('f'),
                                Hyperlink: switches.ContainsKey('h'),
                                ParagraphNumber: switches.ContainsKey('n'),
                                AboveBelow: switches.ContainsKey('p'),
                                RelativeParagraphNumber: switches.ContainsKey('r'),
                                SuppressNonNumeric: switches.ContainsKey('t'),
                                FullContextParagraphNumber: switches.ContainsKey('w'),
                                SeparatorText: switches.TryGetValue('d', out var sep) ? sep : null);

                            var result = await Context.RefResolver.ResolveAsync(request, Context);
                            if (result != null)
                            {
                                string text = result.Text ?? string.Empty;
                                if (request.SuppressNonNumeric)
                                    text = StripNonNumeric(text);
                                if (!string.IsNullOrWhiteSpace(result.HyperlinkTarget))
                                    Context.RefHyperlinks.Add(new Resolution.DxpRefHyperlink(bookmark, result.HyperlinkTarget!, text));
                                if (!string.IsNullOrWhiteSpace(result.FootnoteText))
                                    Context.RefFootnotes.Add(new Resolution.DxpRefFootnote(bookmark, result.FootnoteText!, result.FootnoteMark));
                                value = new DxpFieldValue(text);
                                return (true, value);
                            }
                        }

                        if (Context.TryGetBookmarkNodes(bookmark, out var bmNodes))
                        {
                            string text = bmNodes.ToPlainText();
                            if (hasRefSwitches && switches.ContainsKey('t'))
                                text = StripNonNumeric(text);
                            value = new DxpFieldValue(text);
                            return (true, value);
                        }

                        _logger?.LogWarning("REF could not resolve bookmark '{Bookmark}'.", bookmark);
                        value = new DxpFieldValue(RefNotFoundError);
                        return (true, value);
                    }
                }
                else
                {
                    _logger?.LogWarning("REF field missing arguments.");
                }
                value = new DxpFieldValue(RefNotFoundError);
                return (true, value);
            case "DOCVARIABLE":
                if (ast.ArgumentsText != null)
                {
                    var tokens = TokenizeArgs(ast.ArgumentsText);
                    if (tokens.Count > 0)
                    {
                        var resolver = Context.ValueResolver ?? _resolver;
                        var name = await ExpandNestedTextAsync(tokens[0]);
                        DxpFieldNodeBuffer? resolvedNodes = null;
                        if (_delegates.ResolveDocVariableNodesAsync != null)
                        {
                            resolvedNodes = await _delegates.ResolveDocVariableNodesAsync(name, Context);
                            if (resolvedNodes != null)
                            {
                                Context.SetDocVariableNodes(name, resolvedNodes);
                                value = new DxpFieldValue(resolvedNodes.ToPlainText());
                                return (true, value);
                            }
                        }

                        var resolved = await resolver.ResolveAsync(name, Resolution.DxpFieldValueKindHint.DocVariable, Context);
                        if (resolved == null)
                        {
                            value = new DxpFieldValue(DocVariableMissingError);
                            if (_logger?.IsEnabled(LogLevel.Debug) == true)
                                _logger.LogDebug("DOCVARIABLE '{Name}' not resolved; using error text.", name);
                        }
                        else
                        {
                            value = resolved.Value;
                        }

                        Context.SetDocVariableNodes(name, DxpFieldNodeBuffer.FromText(ToDefaultString(value)));
                        return (true, value);
                    }
                }
                _logger?.LogWarning("DOCVARIABLE field missing arguments.");
                value = new DxpFieldValue(DocVariableMissingError);
                return (true, value);
            case "DOCPROPERTY":
                if (ast.ArgumentsText != null)
                {
                    var tokens = TokenizeArgs(ast.ArgumentsText);
                    if (tokens.Count > 0)
                    {
                        var resolver = Context.ValueResolver ?? _resolver;
                        var name = await ExpandNestedTextAsync(tokens[0]);
                        var resolved = await resolver.ResolveAsync(name, Resolution.DxpFieldValueKindHint.DocumentProperty, Context);
                        if (resolved == null)
                        {
                            value = new DxpFieldValue(DocPropertyUnknownError);
                            if (_logger?.IsEnabled(LogLevel.Debug) == true)
                                _logger.LogDebug("DOCPROPERTY '{Name}' not resolved; using error text.", name);
                        }
                        else
                        {
                            value = resolved.Value;
                        }
                        return (true, value);
                    }
                }
                _logger?.LogWarning("DOCPROPERTY field missing arguments.");
                value = new DxpFieldValue(DocPropertyUnknownError);
                return (true, value);
            case "MERGEFIELD":
                if (ast.ArgumentsText != null)
                {
                    var tokens = TokenizeArgs(ast.ArgumentsText);
                    if (tokens.Count > 0)
                    {
                        string name = tokens[0];
                        name = await ExpandNestedTextAsync(name);
                        if (TryResolveMergeFieldName(name, ast, out var mapped))
                            name = mapped;
                        var resolver = Context.ValueResolver ?? _resolver;
                        var resolved = await resolver.ResolveAsync(name, Resolution.DxpFieldValueKindHint.MergeField, Context);
                        value = resolved ?? new DxpFieldValue(string.Empty);
                        if (resolved == null && _logger?.IsEnabled(LogLevel.Debug) == true)
                            _logger.LogDebug("MERGEFIELD '{Name}' not resolved; using empty string.", name);
                        return (true, value);
                    }
                }
                _logger?.LogWarning("MERGEFIELD field missing arguments.");
                return (false, value);
            case "SEQ":
                if (ast.ArgumentsText != null)
                {
                    var tokens = TokenizeArgs(ast.ArgumentsText);
                    if (tokens.Count > 0)
                    {
                        string identifier = tokens[0];
                        string? bookmark = tokens.Count > 1 ? tokens[1] : null;
                        identifier = await ExpandNestedTextAsync(identifier);
                        if (!string.IsNullOrEmpty(bookmark))
                        {
                            var expandedBookmark = await ExpandNestedTextAsync(bookmark!);
                            bookmark = expandedBookmark;
                        }
                        var switches = ParseNonFormatSwitches(ast.RawText);
                        bool repeat = switches.ContainsKey('c');
                        bool hide = switches.ContainsKey('h');
                        bool hasStar = ast.RawText.IndexOf("\\*", StringComparison.Ordinal) >= 0;

                        int? reset = ParseIntSwitchValue(switches, 'r');
                        if (reset == null && switches.ContainsKey('r'))
                            _logger?.LogWarning("SEQ switch '\\r' is invalid; expected integer.");
                        if (reset.HasValue)
                            Context.SetSequence(identifier, reset.Value);

                        if (switches.TryGetValue('s', out var levelText) && !string.IsNullOrWhiteSpace(levelText) &&
                            !int.TryParse(levelText, out _))
                        {
                            _logger?.LogWarning("SEQ switch '\\s' is invalid; expected integer.");
                        }
                        ApplySequenceResetByHeading(identifier, switches);

                        bool bookmarkReset = false;
                        if (!string.IsNullOrWhiteSpace(bookmark))
                        {
                            var bookmarkName = bookmark!;
                            if (Context.TryGetBookmarkNodes(bookmarkName, out var bmNodes) &&
                            TryParseNumber(bmNodes.ToPlainText(), out var bmNumber))
                            {
                                Context.SetSequence(identifier, (int)Math.Floor(bmNumber));
                                bookmarkReset = true;
                            }
                        }

                        int seqValue;
                        if (repeat || reset.HasValue || bookmarkReset)
                            seqValue = Context.GetSequence(identifier);
                        else
                            seqValue = Context.NextSequence(identifier);

                        // Track most recent sequence value for numbereditem lookups.
                        Context.SetNumberedItem(identifier, seqValue.ToString(System.Globalization.CultureInfo.InvariantCulture));

                        bool shouldHide = hide && !hasStar;
                        value = shouldHide ? new DxpFieldValue(string.Empty) : new DxpFieldValue(seqValue);
                        return (true, value);
                    }
                }
                _logger?.LogWarning("SEQ field missing arguments.");
                return (false, value);
            case "ASK":
                if (ast.ArgumentsText != null)
                {
                    var tokens = TokenizeArgs(ast.ArgumentsText);
                    if (tokens.Count > 0)
                    {
                        string bookmark = tokens[0];
                        string prompt = tokens.Count > 1 ? tokens[1] : string.Empty;
                        var switches = ParseNonFormatSwitches(ast.RawText);
                        bool onlyOnce = switches.ContainsKey('o');
                        string? defaultValue = switches.TryGetValue('d', out var def) ? def : null;

                        if (onlyOnce && Context.TryGetBookmarkNodes(bookmark, out _))
                        {
                            value = new DxpFieldValue(string.Empty);
                            return (true, value);
                        }

                        if (!string.IsNullOrEmpty(prompt))
                            prompt = await ResolveValueAsync(prompt);
                        if (!string.IsNullOrEmpty(defaultValue))
                        {
                            var resolvedDefault = await ResolveValueAsync(defaultValue!);
                            defaultValue = resolvedDefault;
                        }

                        DxpFieldValue? response = null;
                        if (_delegates.AskAsync != null)
                            response = await _delegates.AskAsync(prompt, Context);
                        else if (_logger?.IsEnabled(LogLevel.Debug) == true)
                            _logger.LogDebug("ASK delegate not configured; using default value.");

                        string resolved = response?.StringValue
                            ?? (response?.NumberValue?.ToString(Context.Culture ?? System.Globalization.CultureInfo.CurrentCulture))
                            ?? defaultValue
                            ?? string.Empty;

                        Context.SetBookmarkNodes(bookmark, DxpFieldNodeBuffer.FromText(resolved));
                        value = new DxpFieldValue(string.Empty);
                        return (true, value);
                    }
                }
                _logger?.LogWarning("ASK field missing arguments.");
                return (false, value);
            default:
                return (false, value);
        }
    }

    private async Task<DxpFieldEvalResult?> EvalIfAsync(DxpFieldAst ast)
    {
        if (ast.ArgumentsText == null)
            return null;
        if (!TryParseIfArgs(ast.ArgumentsText, out var left, out var op, out var right, out var trueText, out var falseText))
        {
            _logger?.LogWarning("IF field has invalid arguments.");
            return null;
        }

        string leftValue = await ResolveValueAsync(left);
        string rightValue = await ResolveValueAsync(right);
        bool condition = EvaluateComparison(leftValue, op, rightValue);
        _logger?.LogInformation("IF evaluated to {Condition}.", condition);
        string branch = condition ? trueText : (falseText ?? string.Empty);
        string resolvedBranch = await ResolveValueAsync(branch);
        var value = new DxpFieldValue(resolvedBranch);
        string text = _formatter.Format(value, ast.FormatSpecs, Context);
        return new DxpFieldEvalResult(DxpFieldEvalStatus.Resolved, text);
    }

    private async Task<DxpFieldEvalResult?> EvalCompareAsync(DxpFieldAst ast)
    {
        if (ast.ArgumentsText == null)
            return null;
        if (!TryParseCompareArgs(ast.ArgumentsText, out var left, out var op, out var right))
        {
            _logger?.LogWarning("COMPARE field has invalid arguments.");
            return null;
        }

        string leftValue = await ResolveValueAsync(left);
        string rightValue = await ResolveValueAsync(right);
        bool condition = EvaluateComparison(leftValue, op, rightValue);
        _logger?.LogInformation("COMPARE evaluated to {Condition}.", condition);
        var value = new DxpFieldValue(condition ? 1.0 : 0.0);
        string text = _formatter.Format(value, ast.FormatSpecs, Context);
        return new DxpFieldEvalResult(DxpFieldEvalStatus.Resolved, text);
    }

    private async Task<DxpFieldEvalResult?> EvalSkipIfAsync(DxpFieldAst ast)
    {
        if (ast.ArgumentsText == null)
            return null;
        if (!TryParseCompareArgs(ast.ArgumentsText, out var left, out var op, out var right))
        {
            _logger?.LogWarning("SKIPIF/NEXTIF field has invalid arguments.");
            return null;
        }

        string leftValue = await ResolveValueAsync(left);
        string rightValue = await ResolveValueAsync(right);
        bool condition = EvaluateComparison(leftValue, op, rightValue);
        _logger?.LogInformation("SKIPIF evaluated to {Condition}.", condition);
        if (condition)
            return new DxpFieldEvalResult(DxpFieldEvalStatus.Skipped, null);
        return new DxpFieldEvalResult(DxpFieldEvalStatus.Resolved, string.Empty);
    }

    private string ToDefaultString(DxpFieldValue value)
    {
        switch (value.Kind)
        {
            case DxpFieldValueKind.String:
                return value.StringValue ?? string.Empty;
            case DxpFieldValueKind.Number:
            {
                if (!value.NumberValue.HasValue)
                    return string.Empty;
                var culture = Context.Culture ?? CultureInfo.CurrentCulture;
                return value.NumberValue.Value.ToString(culture);
            }
            case DxpFieldValueKind.DateTime:
            {
                if (!value.DateTimeValue.HasValue)
                    return string.Empty;
                var culture = Context.Culture ?? CultureInfo.CurrentCulture;
                return value.DateTimeValue.Value.ToString(culture);
            }
            default:
                return string.Empty;
        }
    }

    private bool TryParseIfArgs(
        string argsText,
        out string left,
        out string op,
        out string right,
        out string trueText,
        out string? falseText)
    {
        left = op = right = trueText = string.Empty;
        falseText = null;
        var tokens = TokenizeArgs(argsText);
        if (tokens.Count < 4)
            return false;

        left = tokens[0];
        op = tokens[1];
        right = tokens[2];
        trueText = tokens[3];
        if (tokens.Count > 4)
            falseText = tokens[4];
        return true;
    }

    private bool TryParseCompareArgs(string argsText, out string left, out string op, out string right)
    {
        left = op = right = string.Empty;
        var tokens = TokenizeArgs(argsText);
        if (tokens.Count < 3)
            return false;
        left = tokens[0];
        op = tokens[1];
        right = tokens[2];
        return true;
    }

    private List<string> TokenizeArgs(string text)
    {
        var tokens = new List<string>();
        bool inQuote = false;
        bool justClosedQuote = false;
        var current = new System.Text.StringBuilder();
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (ch == '"')
            {
                if (inQuote && i > 0 && text[i - 1] == '\\')
                {
                    if (current.Length > 0)
                        current.Length -= 1; // remove escape backslash
                    current.Append('"');
                    continue;
                }
                inQuote = !inQuote;
                if (!inQuote)
                    justClosedQuote = true;
                continue;
            }

            if (!inQuote && ch == '{')
            {
                int start = i;
                int depth = 0;
                for (; i < text.Length; i++)
                {
                    if (text[i] == '{')
                        depth++;
                    else if (text[i] == '}')
                    {
                        depth--;
                        if (depth == 0)
                        {
                            i++;
                            break;
                        }
                    }
                }
                string token = text.Substring(start, i - start).Trim();
                if (current.Length > 0)
                {
                    tokens.Add(current.ToString());
                    current.Clear();
                }
                if (token.Length > 0)
                    tokens.Add(token);
                justClosedQuote = false;
                i--;
                continue;
            }

            if (!inQuote && char.IsWhiteSpace(ch))
            {
                if (current.Length > 0)
                {
                    tokens.Add(current.ToString());
                    current.Clear();
                    justClosedQuote = false;
                }
                else if (justClosedQuote)
                {
                    tokens.Add(string.Empty);
                    justClosedQuote = false;
                }
                continue;
            }

            current.Append(ch);
            justClosedQuote = false;
        }

        if (current.Length > 0)
            tokens.Add(current.ToString());
        else if (justClosedQuote)
            tokens.Add(string.Empty);
        return tokens;
    }

    private bool EvaluateComparison(string leftValue, string op, string rightValue)
    {
        bool leftNumeric = TryParseNumber(leftValue, out var leftNum);
        bool rightNumeric = TryParseNumber(rightValue, out var rightNum);
        bool knownOp = op is "=" or "<>" or ">" or "<" or ">=" or "<=";

        bool isNumeric = leftNumeric && rightNumeric;
        if ((op == "=" || op == "<>") && ContainsWildcards(rightValue))
        {
            bool match = WildcardMatch(leftValue, rightValue);
            return op == "=" ? match : !match;
        }

        if (isNumeric)
        {
            if (!knownOp)
                _logger?.LogWarning("Unsupported comparison operator '{Operator}'.", op);
            return op switch {
                "=" => leftNum == rightNum,
                "<>" => leftNum != rightNum,
                ">" => leftNum > rightNum,
                "<" => leftNum < rightNum,
                ">=" => leftNum >= rightNum,
                "<=" => leftNum <= rightNum,
                _ => false
            };
        }

        int cmp = string.Compare(leftValue, rightValue, StringComparison.Ordinal);
        if (!knownOp)
            _logger?.LogWarning("Unsupported comparison operator '{Operator}'.", op);
        return op switch {
            "=" => cmp == 0,
            "<>" => cmp != 0,
            ">" => cmp > 0,
            "<" => cmp < 0,
            ">=" => cmp >= 0,
            "<=" => cmp <= 0,
            _ => false
        };
    }

    private async Task<string> ResolveValueAsync(string token)
    {
        string unquoted = token;
        if (token.Length >= 2 && token[0] == '"' && token[token.Length - 1] == '"')
            unquoted = token.Substring(1, token.Length - 2);

        string expanded = await ExpandNestedFieldsAsync(unquoted);

        var resolver = Context.ValueResolver ?? _resolver;
        DxpFieldValue? resolvedValue = await resolver.ResolveAsync(expanded, Resolution.DxpFieldValueKindHint.Any, Context);
        if (resolvedValue.HasValue)
        {
            var value = resolvedValue.Value;
            if (value.StringValue != null)
                return value.StringValue;
            if (value.NumberValue != null)
                return value.NumberValue.Value.ToString(Context.Culture ?? System.Globalization.CultureInfo.CurrentCulture);
        }

        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("Value resolver miss for '{Token}'.", expanded);
        return expanded;
    }

    private async Task<string> ExpandNestedTextAsync(string token)
    {
        string unquoted = token;
        if (token.Length >= 2 && token[0] == '"' && token[token.Length - 1] == '"')
            unquoted = token.Substring(1, token.Length - 2);
        return await ExpandNestedFieldsAsync(unquoted);
    }

    private async Task<string> ExpandNestedFieldsAsync(string text)
    {
        int start = text.IndexOf('{');
        if (start < 0)
            return text;

        var sb = new System.Text.StringBuilder();
        int i = 0;
        while (i < text.Length)
        {
            if (text[i] != '{')
            {
                sb.Append(text[i]);
                i++;
                continue;
            }

            int depth = 0;
            int begin = i;
            for (; i < text.Length; i++)
            {
                if (text[i] == '{')
                    depth++;
                else if (text[i] == '}')
                {
                    depth--;
                    if (depth == 0)
                    {
                        i++;
                        break;
                    }
                }
            }

            if (depth != 0)
            {
                _logger?.LogWarning("Unbalanced braces in nested field text; using literal remainder.");
                sb.Append(text.Substring(begin));
                break;
            }

            string inner = text.Substring(begin + 1, i - begin - 2).Trim();
            var nested = await EvalAsync(new DxpFieldInstruction(inner));
            sb.Append(nested.Text ?? string.Empty);
        }

        return sb.ToString();
    }

    private bool TryParseNumber(string text, out double number)
    {
        var culture = Context.Culture ?? System.Globalization.CultureInfo.CurrentCulture;
        if (double.TryParse(text, System.Globalization.NumberStyles.Any, culture, out number))
            return true;
        if (Context.AllowInvariantNumericFallback &&
            double.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out number))
            return true;
        number = 0;
        return false;
    }

    private async Task<double> EvalNestedFieldAsync(string instruction)
    {
        var nested = await EvalAsync(new DxpFieldInstruction(instruction));
        if (nested.Text == null)
            return 0;
        if (TryParseNumber(nested.Text, out var num))
            return num;
        return 0;
    }

    private async Task<DxpFieldValue?> ResolveIdentifierValueAsync(string name)
    {
        var resolver = Context.ValueResolver ?? _resolver;
        return await resolver.ResolveAsync(name, Resolution.DxpFieldValueKindHint.Any, Context);
    }

    private string FormatMergeField(DxpFieldValue value, DxpFieldAst ast)
    {
        string formatted = _formatter.Format(value, ast.FormatSpecs, Context);
        if (string.IsNullOrEmpty(formatted))
            return string.Empty;

        var switches = ParseNonFormatSwitches(ast.RawText);
        if (switches.TryGetValue('v', out var vertical) && vertical == null)
            formatted = string.Join("\n", formatted.ToCharArray());

        switches.TryGetValue('b', out var before);
        switches.TryGetValue('f', out var after);
        return string.Concat(before ?? string.Empty, formatted, after ?? string.Empty);
    }

    private bool TryResolveMergeFieldName(string name, DxpFieldAst ast, out string resolved)
    {
        resolved = name;
        var switches = ParseNonFormatSwitches(ast.RawText);
        if (!switches.ContainsKey('m'))
            return false;
        if (Context.TryGetMergeFieldAlias(name, out var alias) && !string.IsNullOrWhiteSpace(alias))
        {
            resolved = alias!;
            return true;
        }
        return false;
    }

    private Dictionary<char, string?> ParseNonFormatSwitches(string rawText)
    {
        var switches = new Dictionary<char, string?>();
        bool inQuote = false;
        int braceDepth = 0;
        var switchStarts = new List<int>();
        for (int i = 0; i < rawText.Length; i++)
        {
            char ch = rawText[i];
            if (inQuote && ch == '\\' && i + 1 < rawText.Length && rawText[i + 1] == '"')
            {
                i++;
                continue;
            }
            if (ch == '"')
            {
                inQuote = !inQuote;
                continue;
            }
            if (!inQuote)
            {
                if (ch == '{')
                {
                    braceDepth++;
                    continue;
                }
                if (ch == '}' && braceDepth > 0)
                {
                    braceDepth--;
                    continue;
                }
            }
            if (!inQuote && braceDepth == 0 && ch == '\\')
                switchStarts.Add(i);
        }

        for (int i = 0; i < switchStarts.Count; i++)
        {
            int start = switchStarts[i];
            int end = i + 1 < switchStarts.Count ? switchStarts[i + 1] : rawText.Length;
            string seg = rawText.Substring(start, end - start).Trim();
            int index = 0;
            while (index < seg.Length && seg[index] == '\\')
                index++;
            while (index < seg.Length && char.IsWhiteSpace(seg[index]))
                index++;
            if (index >= seg.Length)
                continue;
            char kind = seg[index];
            if (kind is '*' or '#' or '@')
                continue;
            string? arg = index + 1 < seg.Length ? seg.Substring(index + 1).Trim() : null;
            string? unquoted = UnquoteSwitchArg(arg);
            switches[kind] = unquoted;
        }

        return switches;
    }

    private static string? UnquoteSwitchArg(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return value;
        string trimmed = value!.Trim();
        if (trimmed.Length >= 2 && trimmed[0] == '"' && trimmed[trimmed.Length - 1] == '"')
            return trimmed.Substring(1, trimmed.Length - 2);
        return trimmed;
    }

    private static bool HasRefSwitches(Dictionary<char, string?> switches)
    {
        foreach (var key in switches.Keys)
        {
            switch (key)
            {
                case 'd':
                case 'f':
                case 'h':
                case 'n':
                case 'p':
                case 'r':
                case 't':
                case 'w':
                    return true;
            }
        }
        return false;
    }

    private static string StripNonNumeric(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;
        var sb = new System.Text.StringBuilder(text.Length);
        foreach (char ch in text)
        {
            if (char.IsDigit(ch) || ch == '.' || ch == ',' || ch == '-' || ch == '–' || ch == '−')
                sb.Append(ch);
        }
        return sb.ToString();
    }

    private static int? ParseIntSwitchValue(Dictionary<char, string?> switches, char key)
    {
        if (!switches.TryGetValue(key, out var value) || string.IsNullOrWhiteSpace(value))
            return null;
        if (int.TryParse(value, out var parsed))
            return parsed;
        return null;
    }

    private void ApplySequenceResetByHeading(string identifier, Dictionary<char, string?> switches)
    {
        if (!switches.TryGetValue('s', out var value) || string.IsNullOrWhiteSpace(value))
            return;
        if (!int.TryParse(value, out var level))
            return;

        var current = Context.CurrentOutlineLevelProvider?.Invoke() ?? 0;
        if (current <= 0)
            return;

        string key = $"SEQ:{identifier}:LEVEL:{level}";
        if (!Context.TryGetSequenceResetKey(key, out var lastLevel) || lastLevel != current)
        {
            Context.SetSequence(identifier, 0);
            Context.SetSequenceResetKey(key, current);
        }
    }

    private char ResolveListSeparator()
    {
        var listSeparator = Context.ListSeparator;
        if (!string.IsNullOrEmpty(listSeparator))
            return listSeparator![0];

        var culture = Context.Culture ?? System.Globalization.CultureInfo.CurrentCulture;
        string separator = culture.TextInfo.ListSeparator;
        return string.IsNullOrEmpty(separator) ? ',' : separator[0];
    }

    private static bool ContainsWildcards(string text) => text.IndexOfAny(new[] { '*', '?' }) >= 0;

    private static bool WildcardMatch(string input, string pattern)
    {
        int i = 0, p = 0, star = -1, match = 0;
        while (i < input.Length)
        {
            if (p < pattern.Length && (pattern[p] == '?' || pattern[p] == input[i]))
            {
                i++;
                p++;
            }
            else if (p < pattern.Length && pattern[p] == '*')
            {
                star = p++;
                match = i;
            }
            else if (star != -1)
            {
                p = star + 1;
                i = ++match;
            }
            else
            {
                return false;
            }
        }
        while (p < pattern.Length && pattern[p] == '*')
            p++;
        return p == pattern.Length;
    }

    private DxpFieldEvalResult FallbackToCacheOrSkip(DxpFieldInstruction instruction)
    {
        if (_options.UseCacheOnNull && instruction.CachedResult != null)
        {
            _logger?.LogInformation("Using cached result for instruction '{Instruction}'.", instruction.InstructionText);
            return new DxpFieldEvalResult(DxpFieldEvalStatus.UsedCache, instruction.CachedResult);
        }
        _logger?.LogInformation("Skipping field instruction '{Instruction}'.", instruction.InstructionText);
        return new DxpFieldEvalResult(DxpFieldEvalStatus.Skipped, null);
    }

    private static bool IsKnownFieldType(string fieldType)
    {
        return fieldType.Equals("IF", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("COMPARE", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("SKIPIF", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("NEXTIF", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("=", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("DATE", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("TIME", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("CREATEDATE", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("SAVEDATE", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("PRINTDATE", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("SET", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("REF", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("DOCVARIABLE", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("DOCPROPERTY", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("MERGEFIELD", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("SEQ", StringComparison.OrdinalIgnoreCase)
            || fieldType.Equals("ASK", StringComparison.OrdinalIgnoreCase);
    }
}
