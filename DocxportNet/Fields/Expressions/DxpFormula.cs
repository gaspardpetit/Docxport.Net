using System.Globalization;

namespace DocxportNet.Fields.Expressions;

internal enum DxpFormulaTokenKind
{
	Number,
	Identifier,
	Field,
	Operator,
	LParen,
	RParen,
	Comma,
	End
}

internal sealed record DxpFormulaToken(DxpFormulaTokenKind Kind, string Text);

internal abstract record DxpFormulaNode;
internal sealed record DxpNumberNode(string Text) : DxpFormulaNode;
internal sealed record DxpIdentifierNode(string Name) : DxpFormulaNode;
internal sealed record DxpFieldNode(string Instruction) : DxpFormulaNode;
internal sealed record DxpRangeNode(string RangeText, Resolution.DxpTableRangeDirection? Direction) : DxpFormulaNode;
internal sealed record DxpUnaryNode(string Op, DxpFormulaNode Expr) : DxpFormulaNode;
internal sealed record DxpBinaryNode(string Op, DxpFormulaNode Left, DxpFormulaNode Right) : DxpFormulaNode;
internal sealed record DxpCallNode(string Name, IReadOnlyList<DxpFormulaNode> Args) : DxpFormulaNode;

internal sealed class DxpFormulaParser
{
	private readonly List<DxpFormulaToken> _tokens;
	private int _pos;

	private readonly char _listSeparator;

	public DxpFormulaParser(string text, char listSeparator = ',')
	{
		_listSeparator = listSeparator;
		_tokens = Tokenize(text);
		_pos = 0;
	}

	public DxpFormulaNode ParseExpression()
	{
		return ParseComparison();
	}

	private DxpFormulaNode ParseComparison()
	{
		var left = ParseAdd();
		while (MatchOp("=", "<>", ">=", "<=", ">", "<"))
		{
			string op = Previous().Text;
			var right = ParseAdd();
			left = new DxpBinaryNode(op, left, right);
		}
		return left;
	}

	private DxpFormulaNode ParseAdd()
	{
		var left = ParseMul();
		while (MatchOp("+", "-"))
		{
			string op = Previous().Text;
			var right = ParseMul();
			left = new DxpBinaryNode(op, left, right);
		}
		return left;
	}

	private DxpFormulaNode ParseMul()
	{
		var left = ParsePower();
		while (MatchOp("*", "/"))
		{
			string op = Previous().Text;
			var right = ParsePower();
			left = new DxpBinaryNode(op, left, right);
		}
		return left;
	}

	private DxpFormulaNode ParsePower()
	{
		var left = ParsePostfix();
		if (MatchOp("^"))
		{
			string op = Previous().Text;
			var right = ParsePower();
			return new DxpBinaryNode(op, left, right);
		}
		return left;
	}

	private DxpFormulaNode ParseUnary()
	{
		return ParsePostfix();
	}

	private DxpFormulaNode ParseSignedPrimary()
	{
		if (MatchOp("+", "-"))
		{
			string op = Previous().Text;
			return new DxpUnaryNode(op, ParseSignedPrimary());
		}
		return ParsePrimary();
	}

	private DxpFormulaNode ParsePostfix()
	{
		var expr = ParseSignedPrimary();
		while (MatchOp("%"))
		{
			expr = new DxpUnaryNode("%", expr);
		}
		return expr;
	}

	private DxpFormulaNode ParsePrimary()
	{
		if (Match(DxpFormulaTokenKind.Number))
			return new DxpNumberNode(Previous().Text);
		if (Match(DxpFormulaTokenKind.Field))
			return new DxpFieldNode(Previous().Text);
		if (Match(DxpFormulaTokenKind.Identifier))
		{
			string name = Previous().Text;
			if (TryParseDirectionalRange(name, out var dir))
				return new DxpRangeNode(name, dir);
			if (Match(DxpFormulaTokenKind.LParen))
			{
				var args = new List<DxpFormulaNode>();
				if (!Check(DxpFormulaTokenKind.RParen))
				{
					do
					{
						args.Add(ParseExpression());
					} while (Match(DxpFormulaTokenKind.Comma));
				}
				Consume(DxpFormulaTokenKind.RParen);
				return new DxpCallNode(name, args);
			}
			if (TryParseCellReference(name))
			{
				if (MatchOp(":") && Match(DxpFormulaTokenKind.Identifier))
				{
					string end = Previous().Text;
					if (TryParseCellReference(end))
						return new DxpRangeNode($"{name}:{end}", null);
					_pos -= 2;
				}
				return new DxpRangeNode(name, null);
			}
			return new DxpIdentifierNode(name);
		}
		if (Match(DxpFormulaTokenKind.LParen))
		{
			var expr = ParseExpression();
			Consume(DxpFormulaTokenKind.RParen);
			return expr;
		}
		return new DxpNumberNode("0");
	}

	private bool MatchOp(params string[] ops)
	{
		if (Check(DxpFormulaTokenKind.Operator))
		{
			string text = Peek().Text;
			foreach (var op in ops)
			{
				if (text == op)
				{
					Advance();
					return true;
				}
			}
		}
		return false;
	}

	private bool Match(DxpFormulaTokenKind kind)
	{
		if (Check(kind))
		{
			Advance();
			return true;
		}
		return false;
	}

	private void Consume(DxpFormulaTokenKind kind)
	{
		if (Check(kind))
		{
			Advance();
			return;
		}
	}

	private bool Check(DxpFormulaTokenKind kind)
	{
		return Peek().Kind == kind;
	}

	private DxpFormulaToken Advance()
	{
		if (!IsAtEnd())
			_pos++;
		return Previous();
	}

	private bool IsAtEnd() => Peek().Kind == DxpFormulaTokenKind.End;
	private DxpFormulaToken Peek() => _tokens[_pos];
	private DxpFormulaToken Previous() => _tokens[_pos - 1];

	private static bool TryParseCellReference(string name)
	{
		if (string.IsNullOrEmpty(name))
			return false;
		int i = 0;
		while (i < name.Length && char.IsLetter(name[i]))
			i++;
		if (i == 0 || i == name.Length)
			return false;
		for (int j = i; j < name.Length; j++)
		{
			if (!char.IsDigit(name[j]))
				return false;
		}
		return true;
	}

	private static bool TryParseDirectionalRange(string name, out Resolution.DxpTableRangeDirection direction)
	{
		switch (name.ToUpperInvariant())
		{
			case "ABOVE":
				direction = Resolution.DxpTableRangeDirection.Above;
				return true;
			case "BELOW":
				direction = Resolution.DxpTableRangeDirection.Below;
				return true;
			case "LEFT":
				direction = Resolution.DxpTableRangeDirection.Left;
				return true;
			case "RIGHT":
				direction = Resolution.DxpTableRangeDirection.Right;
				return true;
			default:
				direction = Resolution.DxpTableRangeDirection.Above;
				return false;
		}
	}

	private List<DxpFormulaToken> Tokenize(string text)
	{
		var tokens = new List<DxpFormulaToken>();
		int i = 0;
		while (i < text.Length)
		{
			if (text[i] == '{')
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
				string inner = text.Substring(start + 1, i - start - 2).Trim();
				tokens.Add(new DxpFormulaToken(DxpFormulaTokenKind.Field, inner));
				continue;
			}
			char ch = text[i];
			if (char.IsWhiteSpace(ch))
			{
				i++;
				continue;
			}
			if (char.IsDigit(ch) || ch == '.')
			{
				int start = i++;
				while (i < text.Length && (char.IsDigit(text[i]) || text[i] == '.'))
					i++;
				tokens.Add(new DxpFormulaToken(DxpFormulaTokenKind.Number, text.Substring(start, i - start)));
				continue;
			}
			if (char.IsLetter(ch) || ch == '_')
			{
				int start = i++;
				while (i < text.Length && (char.IsLetterOrDigit(text[i]) || text[i] == '_' || text[i] == '.'))
					i++;
				tokens.Add(new DxpFormulaToken(DxpFormulaTokenKind.Identifier, text.Substring(start, i - start)));
				continue;
			}
			switch (ch)
			{
				case '(':
					tokens.Add(new DxpFormulaToken(DxpFormulaTokenKind.LParen, "("));
					i++;
					break;
				case ')':
					tokens.Add(new DxpFormulaToken(DxpFormulaTokenKind.RParen, ")"));
					i++;
					break;
				case ',':
					tokens.Add(new DxpFormulaToken(DxpFormulaTokenKind.Comma, ","));
					i++;
					break;
				case ';':
					if (_listSeparator == ';')
					{
						tokens.Add(new DxpFormulaToken(DxpFormulaTokenKind.Comma, ";"));
						i++;
						break;
					}
					goto default;
				case var _ when ch == _listSeparator && ch != ',':
					tokens.Add(new DxpFormulaToken(DxpFormulaTokenKind.Comma, ","));
					i++;
					break;
				default:
				{
					string op = ch.ToString();
					if (i + 1 < text.Length)
					{
						string two = text.Substring(i, 2);
						if (two is ">=" or "<=" or "<>")
						{
							op = two;
							i += 2;
							tokens.Add(new DxpFormulaToken(DxpFormulaTokenKind.Operator, op));
							break;
						}
					}
					tokens.Add(new DxpFormulaToken(DxpFormulaTokenKind.Operator, op));
					i++;
					break;
				}
			}
		}
		tokens.Add(new DxpFormulaToken(DxpFormulaTokenKind.End, string.Empty));
		return tokens;
	}
}

internal sealed class DxpFormulaEvaluator
{
	private readonly DxpFieldEvalContext _context;
	private readonly Func<string, Task<double>> _evalNestedFieldAsync;
	private readonly Func<string, Task<DxpFieldValue?>> _resolveIdentifierValueAsync;

	public DxpFormulaEvaluator(
		DxpFieldEvalContext context,
		Func<string, Task<double>> evalNestedFieldAsync,
		Func<string, Task<DxpFieldValue?>> resolveIdentifierValueAsync)
	{
		_context = context;
		_evalNestedFieldAsync = evalNestedFieldAsync;
		_resolveIdentifierValueAsync = resolveIdentifierValueAsync;
	}

	public async Task<double> EvaluateAsync(DxpFormulaNode node)
	{
		switch (node)
		{
			case DxpNumberNode n:
				return ParseNumber(n.Text);
			case DxpIdentifierNode id:
				return await ResolveIdentifierAsync(id.Name);
			case DxpFieldNode f:
				return await _evalNestedFieldAsync(f.Instruction);
			case DxpRangeNode r:
				return await ResolveRangeAsNumberAsync(r);
			case DxpUnaryNode u:
				return await EvalUnaryAsync(u);
			case DxpBinaryNode b:
				return await EvalBinaryAsync(b);
			case DxpCallNode c:
				return await EvalCallAsync(c);
			default:
				return 0;
		}
	}

	private async Task<double> EvalUnaryAsync(DxpUnaryNode u)
	{
		double v = await EvaluateAsync(u.Expr);
		return u.Op switch
		{
			"+" => v,
			"-" => -v,
			"%" => v / 100.0,
			_ => v
		};
	}

	private async Task<double> EvalBinaryAsync(DxpBinaryNode b)
	{
		double left = await EvaluateAsync(b.Left);
		double right = await EvaluateAsync(b.Right);
		return b.Op switch
		{
			"+" => left + right,
			"-" => left - right,
			"*" => left * right,
			"/" => right == 0 ? 0 : left / right,
			"^" => Math.Pow(left, right),
			"=" => left == right ? 1 : 0,
			"<>" => left != right ? 1 : 0,
			">" => left > right ? 1 : 0,
			"<" => left < right ? 1 : 0,
			">=" => left >= right ? 1 : 0,
			"<=" => left <= right ? 1 : 0,
			_ => 0
		};
	}

	private async Task<double> EvalCallAsync(DxpCallNode call)
	{
		if (call.Name.Equals("IF", StringComparison.OrdinalIgnoreCase))
			return await EvalIfFunctionAsync(call);
		if (call.Name.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
			return await EvalTrueFunctionAsync(call);
		if (call.Name.Equals("DEFINED", StringComparison.OrdinalIgnoreCase))
			return await EvalDefinedAsync(call);

		var args = await EvaluateArgsAsync(call.Args);
		if (_context.FormulaFunctions.TryResolve(call.Name, out var fn))
			return fn(args);
		return 0;
	}

	private async Task<double> EvalIfFunctionAsync(DxpCallNode call)
	{
		if (call.Args.Count == 0)
			return 0;
		double test = await EvaluateAsync(call.Args[0]);
		if (test != 0)
			return call.Args.Count > 1 ? await EvaluateAsync(call.Args[1]) : 0;
		return call.Args.Count > 2 ? await EvaluateAsync(call.Args[2]) : 0;
	}

	private async Task<double> EvalTrueFunctionAsync(DxpCallNode call)
	{
		if (call.Args.Count == 0)
			return 1;
		double test = await EvaluateAsync(call.Args[0]);
		return test != 0 ? 1 : 0;
	}

	private async Task<double[]> EvaluateArgsAsync(IReadOnlyList<DxpFormulaNode> args)
	{
		if (args.Count == 0)
			return Array.Empty<double>();
		var values = new List<double>(args.Count);
		foreach (var arg in args)
		{
			if (arg is DxpRangeNode range)
			{
				var rangeValues = await ResolveRangeValuesAsync(range);
				values.AddRange(rangeValues);
			}
			else
			{
				values.Add(await EvaluateAsync(arg));
			}
		}
		return values.ToArray();
	}

	private async Task<double> ResolveRangeAsNumberAsync(DxpRangeNode range)
	{
		var values = await ResolveRangeValuesAsync(range);
		if (values.Count == 0)
			return 0;
		if (values.Count == 1)
			return values[0];
		double sum = 0;
		foreach (var v in values)
			sum += v;
		return sum;
	}

	private async Task<IReadOnlyList<double>> ResolveRangeValuesAsync(DxpRangeNode range)
	{
		var resolver = _context.TableResolver;
		if (resolver == null)
			return Array.Empty<double>();
		if (range.Direction.HasValue)
			return await resolver.ResolveDirectionalRangeAsync(range.Direction.Value, _context);
		return await resolver.ResolveRangeAsync(range.RangeText, _context);
	}

	private async Task<double> ResolveIdentifierAsync(string name)
	{
		var resolved = await _resolveIdentifierValueAsync(name);
		if (resolved.HasValue)
		{
			var value = resolved.Value;
			if (value.NumberValue != null)
				return value.NumberValue.Value;
			if (value.StringValue != null && TryParse(value.StringValue, out var n1))
				return n1;
		}
		return 0;
	}

	private async Task<double> EvalDefinedAsync(DxpCallNode call)
	{
		if (call.Args.Count == 0)
			return 0;

		try
		{
			var arg = call.Args[0];
			switch (arg)
			{
				case DxpIdentifierNode id:
				{
					var resolved = await _resolveIdentifierValueAsync(id.Name);
					return resolved.HasValue ? 1 : 0;
				}
				case DxpFieldNode f:
				{
					await _evalNestedFieldAsync(f.Instruction);
					return 1;
				}
				default:
					await EvaluateAsync(arg);
					return 1;
			}
		}
		catch
		{
			return 0;
		}
	}

	private double ParseNumber(string text)
	{
		if (TryParse(text, out var number))
			return number;
		return 0;
	}

	private bool TryParse(string text, out double number)
	{
		var culture = _context.Culture ?? CultureInfo.CurrentCulture;
		if (double.TryParse(text, NumberStyles.Any, culture, out number))
			return true;
		if (_context.AllowInvariantNumericFallback &&
			double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out number))
			return true;
		number = 0;
		return false;
	}
}
