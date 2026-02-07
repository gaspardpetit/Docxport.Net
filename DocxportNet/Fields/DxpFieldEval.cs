namespace DocxportNet.Fields;

public sealed class DxpFieldEval
{
	private readonly DxpFieldParser _parser = new();
	private readonly DxpFieldFormatter _formatter = new();
	private readonly DxpFieldEvalDelegates _delegates;
	private readonly DxpFieldEvalOptions _options;
	private readonly Resolution.IDxpFieldValueResolver _resolver;

	public DxpFieldEvalContext Context { get; } = new();

	public DxpFieldEval(DxpFieldEvalDelegates? delegates = null, DxpFieldEvalOptions? options = null)
	{
		_delegates = delegates ?? new DxpFieldEvalDelegates();
		_options = options ?? new DxpFieldEvalOptions();
		_resolver = new Resolution.DxpChainedFieldValueResolver(
			new Resolution.DxpContextFieldValueResolver(),
			new Resolution.DxpDelegateFieldValueResolver(_delegates)
		);
	}

	public async Task<DxpFieldEvalResult> EvalAsync(DxpFieldInstruction instruction, CancellationToken cancellationToken = default)
	{
		_ = cancellationToken;
		_ = _delegates;
		var parse = _parser.Parse(instruction.InstructionText);
		if (!parse.Success)
		{
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
					return new DxpFieldEvalResult(DxpFieldEvalStatus.Resolved, text);
				}
			}
		}

		if (_options.ErrorOnUnsupported)
		{
			return new DxpFieldEvalResult(
				DxpFieldEvalStatus.Failed,
				null,
				new NotSupportedException("Field evaluation is not implemented yet."));
		}

		return FallbackToCacheOrSkip(instruction);
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
					return (true, value);
				}
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
				return (false, value);
			case "SAVEDATE":
				if (Context.SavedDate != null)
				{
					value = new DxpFieldValue(Context.SavedDate.Value);
					return (true, value);
				}
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
						Context.SetBookmark(name, resolved);
						value = new DxpFieldValue(resolved);
						return (true, value);
					}
				}
				return (false, value);
			case "REF":
				if (ast.ArgumentsText != null)
				{
					var tokens = TokenizeArgs(ast.ArgumentsText);
					if (tokens.Count > 0)
					{
						string bookmark = tokens[0];
						var switches = ParseNonFormatSwitches(ast.RawText);
						bool hasRefSwitches = HasRefSwitches(switches);
						if (hasRefSwitches && Context.RefResolver != null)
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

						if (Context.TryGetBookmark(bookmark, out var bm))
						{
							string text = bm ?? string.Empty;
							if (hasRefSwitches && switches.ContainsKey('t'))
								text = StripNonNumeric(text);
							value = new DxpFieldValue(text);
							return (true, value);
						}
					}
				}
				return (false, value);
			case "DOCVARIABLE":
				if (ast.ArgumentsText != null)
				{
					var tokens = TokenizeArgs(ast.ArgumentsText);
					if (tokens.Count > 0)
					{
						var resolver = Context.ValueResolver ?? _resolver;
						var resolved = await resolver.ResolveAsync(tokens[0], Resolution.DxpFieldValueKindHint.DocVariable, Context);
						value = resolved ?? new DxpFieldValue(string.Empty);
						return (true, value);
					}
				}
				return (false, value);
			case "DOCPROPERTY":
				if (ast.ArgumentsText != null)
				{
					var tokens = TokenizeArgs(ast.ArgumentsText);
					if (tokens.Count > 0)
					{
						var resolver = Context.ValueResolver ?? _resolver;
						var resolved = await resolver.ResolveAsync(tokens[0], Resolution.DxpFieldValueKindHint.DocumentProperty, Context);
						value = resolved ?? new DxpFieldValue(string.Empty);
						return (true, value);
					}
				}
				return (false, value);
			case "MERGEFIELD":
				if (ast.ArgumentsText != null)
				{
					var tokens = TokenizeArgs(ast.ArgumentsText);
					if (tokens.Count > 0)
					{
						string name = tokens[0];
						if (TryResolveMergeFieldName(name, ast, out var mapped))
							name = mapped;
						var resolver = Context.ValueResolver ?? _resolver;
						var resolved = await resolver.ResolveAsync(name, Resolution.DxpFieldValueKindHint.MergeField, Context);
						value = resolved ?? new DxpFieldValue(string.Empty);
						return (true, value);
					}
				}
				return (false, value);
			case "SEQ":
				if (ast.ArgumentsText != null)
				{
					var tokens = TokenizeArgs(ast.ArgumentsText);
					if (tokens.Count > 0)
					{
						string identifier = tokens[0];
						string? bookmark = tokens.Count > 1 ? tokens[1] : null;
						var switches = ParseNonFormatSwitches(ast.RawText);
						bool repeat = switches.ContainsKey('c');
						bool hide = switches.ContainsKey('h');
						bool hasStar = ast.RawText.IndexOf("\\*", StringComparison.Ordinal) >= 0;

						int? reset = ParseIntSwitchValue(switches, 'r');
						if (reset.HasValue)
							Context.SetSequence(identifier, reset.Value);

						ApplySequenceResetByHeading(identifier, switches);

							bool bookmarkReset = false;
							if (!string.IsNullOrWhiteSpace(bookmark))
							{
								var bookmarkName = bookmark!;
								if (Context.TryGetBookmark(bookmarkName, out var bmValue) &&
								bmValue != null &&
								TryParseNumber(bmValue, out var bmNumber))
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

						bool shouldHide = hide && !hasStar;
						value = shouldHide ? new DxpFieldValue(string.Empty) : new DxpFieldValue(seqValue);
						return (true, value);
					}
				}
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

						if (onlyOnce && Context.TryGetBookmark(bookmark, out var existing) && existing != null)
						{
							value = new DxpFieldValue(string.Empty);
							return (true, value);
						}

						DxpFieldValue? response = null;
						if (_delegates.AskAsync != null)
							response = await _delegates.AskAsync(prompt, Context);

						string resolved = response?.StringValue
							?? (response?.NumberValue?.ToString(Context.Culture ?? System.Globalization.CultureInfo.CurrentCulture))
							?? defaultValue
							?? string.Empty;

						Context.SetBookmark(bookmark, resolved);
						value = new DxpFieldValue(string.Empty);
						return (true, value);
					}
				}
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
			return null;

		string leftValue = await ResolveValueAsync(left);
		string rightValue = await ResolveValueAsync(right);
		bool condition = EvaluateComparison(leftValue, op, rightValue);
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
			return null;

		string leftValue = await ResolveValueAsync(left);
		string rightValue = await ResolveValueAsync(right);
		bool condition = EvaluateComparison(leftValue, op, rightValue);
		var value = new DxpFieldValue(condition ? 1.0 : 0.0);
		string text = _formatter.Format(value, ast.FormatSpecs, Context);
		return new DxpFieldEvalResult(DxpFieldEvalStatus.Resolved, text);
	}

	private async Task<DxpFieldEvalResult?> EvalSkipIfAsync(DxpFieldAst ast)
	{
		if (ast.ArgumentsText == null)
			return null;
		if (!TryParseCompareArgs(ast.ArgumentsText, out var left, out var op, out var right))
			return null;

		string leftValue = await ResolveValueAsync(left);
		string rightValue = await ResolveValueAsync(right);
		bool condition = EvaluateComparison(leftValue, op, rightValue);
		if (condition)
			return new DxpFieldEvalResult(DxpFieldEvalStatus.Skipped, null);
		return new DxpFieldEvalResult(DxpFieldEvalStatus.Resolved, string.Empty);
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
				i--;
				continue;
			}

			if (!inQuote && char.IsWhiteSpace(ch))
			{
				if (current.Length > 0)
				{
					tokens.Add(current.ToString());
					current.Clear();
				}
				continue;
			}

			current.Append(ch);
		}

		if (current.Length > 0)
			tokens.Add(current.ToString());
		return tokens;
	}

	private bool EvaluateComparison(string leftValue, string op, string rightValue)
	{
		bool leftNumeric = TryParseNumber(leftValue, out var leftNum);
		bool rightNumeric = TryParseNumber(rightValue, out var rightNum);

		bool isNumeric = leftNumeric && rightNumeric;
		if ((op == "=" || op == "<>") && ContainsWildcards(rightValue))
		{
			bool match = WildcardMatch(leftValue, rightValue);
			return op == "=" ? match : !match;
		}

		if (isNumeric)
		{
			return op switch
			{
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
		return op switch
		{
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

		return expanded;
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
			resolved = alias;
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
				return listSeparator[0];

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
			return new DxpFieldEvalResult(DxpFieldEvalStatus.UsedCache, instruction.CachedResult);
		return new DxpFieldEvalResult(DxpFieldEvalStatus.Skipped, null);
	}
}
