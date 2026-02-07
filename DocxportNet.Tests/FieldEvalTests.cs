using System.Globalization;
using DocxportNet.Fields;
using Xunit;

namespace DocxportNet.Tests;

public class FieldEvalTests
{
	[Fact]
	public async Task EvalAsync_UsesCache_WhenAllowed()
	{
		var eval = new DxpFieldEval(options: new DxpFieldEvalOptions { UseCacheOnNull = true });
		var instruction = new DxpFieldInstruction(" FOO ", "cached");

		var result = await eval.EvalAsync(instruction);

		Assert.Equal(DxpFieldEvalStatus.UsedCache, result.Status);
		Assert.Equal("cached", result.Text);
		Assert.Null(result.Error);
	}

	[Fact]
	public async Task EvalAsync_Skips_WhenNoCacheAndUnsupported()
	{
		var eval = new DxpFieldEval(options: new DxpFieldEvalOptions { UseCacheOnNull = true, ErrorOnUnsupported = false });
		var instruction = new DxpFieldInstruction("FOO", null);

		var result = await eval.EvalAsync(instruction);

		Assert.Equal(DxpFieldEvalStatus.Skipped, result.Status);
		Assert.Null(result.Text);
		Assert.Null(result.Error);
	}

	[Fact]
	public async Task EvalAsync_Fails_WhenConfiguredToErrorOnUnsupported()
	{
		var eval = new DxpFieldEval(options: new DxpFieldEvalOptions { ErrorOnUnsupported = true });
		var instruction = new DxpFieldInstruction("FOO", "cached");

		var result = await eval.EvalAsync(instruction);

		Assert.Equal(DxpFieldEvalStatus.Failed, result.Status);
		Assert.Null(result.Text);
		Assert.IsType<NotSupportedException>(result.Error);
	}

	[Fact]
	public async Task EvalAsync_DateAndTimeUseNowProvider()
	{
		var eval = new DxpFieldEval();
		var now = new DateTimeOffset(2026, 2, 7, 10, 30, 0, TimeSpan.Zero);
		eval.Context.SetNow(() => now);

		var date = await eval.EvalAsync(new DxpFieldInstruction("DATE \\@ \"yyyy-MM-dd\""));
		var time = await eval.EvalAsync(new DxpFieldInstruction("TIME \\@ \"HH:mm\""));

		Assert.Equal(DxpFieldEvalStatus.Resolved, date.Status);
		Assert.Equal("2026-02-07", date.Text);
		Assert.Equal(DxpFieldEvalStatus.Resolved, time.Status);
		Assert.Equal("10:30", time.Text);
	}

	[Fact]
	public async Task EvalAsync_CreatedAndSavedDatesUseContextOrCache()
	{
		var eval = new DxpFieldEval();
		eval.Context.CreatedDate = new DateTimeOffset(2026, 1, 1, 12, 0, 0, TimeSpan.Zero);

		var created = await eval.EvalAsync(new DxpFieldInstruction("CREATEDATE \\@ \"yyyy-MM-dd\""));
		var saved = await eval.EvalAsync(new DxpFieldInstruction("SAVEDATE \\@ \"yyyy-MM-dd\"", "cached"));

		Assert.Equal(DxpFieldEvalStatus.Resolved, created.Status);
		Assert.Equal("2026-01-01", created.Text);
		Assert.Equal(DxpFieldEvalStatus.UsedCache, saved.Status);
		Assert.Equal("cached", saved.Text);
	}

	[Fact]
	public async Task EvalAsync_PrintDateDefaultsToNowIfUnset()
	{
		var eval = new DxpFieldEval();
		var now = new DateTimeOffset(2026, 2, 7, 8, 0, 0, TimeSpan.Zero);
		eval.Context.SetNow(() => now);

		var result = await eval.EvalAsync(new DxpFieldInstruction("PRINTDATE \\@ \"yyyy-MM-dd\""));

		Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
		Assert.Equal("2026-02-07", result.Text);
	}

	[Fact]
	public async Task EvalAsync_IfNumericComparison()
	{
		var eval = new DxpFieldEval();
		eval.Context.SetBookmark("Order", "120");

		var result = await eval.EvalAsync(new DxpFieldInstruction("IF Order >= 100 \"Thanks\" \"No\""));

		Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
		Assert.Equal("Thanks", result.Text);
	}

	[Fact]
	public async Task EvalAsync_IfStringComparisonAndWildcard()
	{
		var eval = new DxpFieldEval();
		eval.Context.SetDocVariable("Status", "Approved");

		var result = await eval.EvalAsync(new DxpFieldInstruction("IF Status = \"App*\" \"Yes\" \"No\""));

		Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
		Assert.Equal("Yes", result.Text);
	}

	[Fact]
	public async Task EvalAsync_IfNestedFieldComparison()
	{
		var eval = new DxpFieldEval();
		eval.Context.CreatedDate = new DateTimeOffset(2026, 2, 7, 0, 0, 0, TimeSpan.Zero);

		var result = await eval.EvalAsync(new DxpFieldInstruction("IF { CREATEDATE \\@ \"yyyy\" } = \"2026\" \"Y\" \"N\""));

		Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
		Assert.Equal("Y", result.Text);
	}

	[Fact]
	public async Task EvalAsync_IfNestedFieldInResultText()
	{
		var eval = new DxpFieldEval();
		eval.Context.SetNow(() => new DateTimeOffset(2026, 2, 7, 9, 0, 0, TimeSpan.Zero));

		var result = await eval.EvalAsync(new DxpFieldInstruction("IF 1 = 1 \"Today: { DATE \\@ \\\"yyyy-MM-dd\\\" }\" \"No\""));

		Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
		Assert.Equal("Today: 2026-02-07", result.Text);
	}

	[Fact]
	public async Task EvalAsync_FormulaArithmeticAndFunctions()
	{
		var eval = new DxpFieldEval();
		eval.Context.Culture = new CultureInfo("en-US");
		eval.Context.SetBookmark("A", "10");
		eval.Context.SetBookmark("B", "5");

		var result = await eval.EvalAsync(new DxpFieldInstruction("= (A + B) * 2"));
		var sum = await eval.EvalAsync(new DxpFieldInstruction("= SUM(1, 2, 3)"));
		var round = await eval.EvalAsync(new DxpFieldInstruction("= ROUND(3.14159, 2)"));

		Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
		Assert.Equal("30", result.Text);
		Assert.Equal("6", sum.Text);
		Assert.Equal("3.14", round.Text);
	}

	[Fact]
	public async Task EvalAsync_FormulaUsesListSeparatorFromContext()
	{
		var eval = new DxpFieldEval();
		eval.Context.Culture = new CultureInfo("fr-FR");
		eval.Context.ListSeparator = ";";

		var result = await eval.EvalAsync(new DxpFieldInstruction("= SUM(1; 2; 3)"));

		Assert.Equal("6", result.Text);
	}

	[Fact]
	public async Task EvalAsync_FormulaCustomFunction()
	{
		var eval = new DxpFieldEval();
		eval.Context.Culture = new CultureInfo("en-US");
		eval.Context.FormulaFunctions.Register("DOUBLE", args => args.Count > 0 ? args[0] * 2 : 0);

		var result = await eval.EvalAsync(new DxpFieldInstruction("= DOUBLE(4)"));

		Assert.Equal("8", result.Text);
	}

	[Fact]
	public async Task EvalAsync_ResolvesVariableViaDelegate()
	{
		var eval = new DxpFieldEval(new DxpFieldEvalDelegates
		{
			ResolveDocVariableAsync = (name, ctx) => Task.FromResult<DxpFieldValue?>(name == "X" ? new DxpFieldValue(5) : null)
		});
		eval.Context.Culture = new CultureInfo("en-US");

		var result = await eval.EvalAsync(new DxpFieldInstruction("= X + 1"));

		Assert.Equal("6", result.Text);
	}

	[Fact]
	public async Task EvalAsync_ResolvesVariableViaCustomResolver()
	{
		var eval = new DxpFieldEval();
		eval.Context.ValueResolver = new DocxportNet.Fields.Resolution.DxpChainedFieldValueResolver(
			new DocxportNet.Fields.Resolution.DxpContextFieldValueResolver(),
			new CustomResolver());
		eval.Context.Culture = new CultureInfo("en-US");

		var result = await eval.EvalAsync(new DxpFieldInstruction("= Y + 1"));

		Assert.Equal("10", result.Text);
	}

	[Fact]
	public async Task EvalAsync_SetAndRefBookmark()
	{
		var eval = new DxpFieldEval();

		var set = await eval.EvalAsync(new DxpFieldInstruction("SET Total \"42\""));
		var get = await eval.EvalAsync(new DxpFieldInstruction("REF Total"));

		Assert.Equal("42", set.Text);
		Assert.Equal("42", get.Text);
	}

	[Fact]
	public async Task EvalAsync_DocPropertyUsesContextValue()
	{
		var eval = new DxpFieldEval();
		eval.Context.SetDocumentPropertyValue("Title", new DxpFieldValue("Doc Title"));

		var result = await eval.EvalAsync(new DxpFieldInstruction("DOCPROPERTY Title"));

		Assert.Equal("Doc Title", result.Text);
	}

	[Fact]
	public async Task EvalAsync_DocPropertyDateFormats()
	{
		var eval = new DxpFieldEval();
		eval.Context.SetDocumentPropertyValue("CreateDate", new DxpFieldValue(new DateTimeOffset(2026, 2, 7, 0, 0, 0, TimeSpan.Zero)));

		var result = await eval.EvalAsync(new DxpFieldInstruction("DOCPROPERTY CreateDate \\@ \"yyyy-MM-dd\""));

		Assert.Equal("2026-02-07", result.Text);
	}

	[Fact]
	public async Task EvalAsync_DocVariableUsesResolver()
	{
		var eval = new DxpFieldEval(new DxpFieldEvalDelegates
		{
			ResolveDocVariableAsync = (name, ctx) => Task.FromResult<DxpFieldValue?>(name == "X" ? new DxpFieldValue("ok") : null)
		});

		var result = await eval.EvalAsync(new DxpFieldInstruction("DOCVARIABLE X"));
		var missing = await eval.EvalAsync(new DxpFieldInstruction("DOCVARIABLE Missing"));

		Assert.Equal("ok", result.Text);
		Assert.Equal(string.Empty, missing.Text);
	}

	[Fact]
	public async Task EvalAsync_MergeFieldUsesResolverAndSwitches()
	{
		var eval = new DxpFieldEval(new DxpFieldEvalDelegates
		{
			ResolveMergeFieldAsync = (name, ctx) => Task.FromResult<DxpFieldValue?>(name == "FirstName" ? new DxpFieldValue("Ana") : null)
		});

		var result = await eval.EvalAsync(new DxpFieldInstruction("MERGEFIELD FirstName \\b \"Hello \" \\f \"!\""));
		var missing = await eval.EvalAsync(new DxpFieldInstruction("MERGEFIELD Missing \\b \"Hello \" \\f \"!\""));

		Assert.Equal("Hello Ana!", result.Text);
		Assert.Equal(string.Empty, missing.Text);
	}

	[Fact]
	public async Task EvalAsync_MergeFieldMapsWithM()
	{
		var eval = new DxpFieldEval(new DxpFieldEvalDelegates
		{
			ResolveMergeFieldAsync = (name, ctx) => Task.FromResult<DxpFieldValue?>(name == "GivenName" ? new DxpFieldValue("Ana") : null)
		});
		eval.Context.SetMergeFieldAlias("FirstName", "GivenName");

		var result = await eval.EvalAsync(new DxpFieldInstruction("MERGEFIELD FirstName \\m"));

		Assert.Equal("Ana", result.Text);
	}

	[Fact]
	public async Task EvalAsync_RefUsesResolverAndSwitches()
	{
		var eval = new DxpFieldEval();
		eval.Context.RefResolver = new MockRefResolver();

		var result = await eval.EvalAsync(new DxpFieldInstruction("REF Note1 \\f"));
		var numeric = await eval.EvalAsync(new DxpFieldInstruction("REF Section \\t"));
		var hyperlink = await eval.EvalAsync(new DxpFieldInstruction("REF Link \\h"));

		Assert.Equal("[1]", result.Text);
		Assert.Equal("1.01", numeric.Text);
		Assert.Equal("LinkText", hyperlink.Text);
		Assert.Single(eval.Context.RefFootnotes);
		Assert.Single(eval.Context.RefHyperlinks);
	}

	[Fact]
	public async Task EvalAsync_AskSetsBookmarkAndUsesDefault()
	{
		var eval = new DxpFieldEval(new DxpFieldEvalDelegates
		{
			AskAsync = (prompt, ctx) => Task.FromResult<DxpFieldValue?>(prompt == "Name?" ? new DxpFieldValue("Ana") : null)
		});

		var asked = await eval.EvalAsync(new DxpFieldInstruction("ASK Name \"Name?\" \\d \"Unknown\""));
		var askedDefault = await eval.EvalAsync(new DxpFieldInstruction("ASK City \"City?\" \\d \"Paris\""));
		var name = await eval.EvalAsync(new DxpFieldInstruction("REF Name"));
		var city = await eval.EvalAsync(new DxpFieldInstruction("REF City"));

		Assert.Equal(string.Empty, asked.Text);
		Assert.Equal(string.Empty, askedDefault.Text);
		Assert.Equal("Ana", name.Text);
		Assert.Equal("Paris", city.Text);
	}

	[Fact]
	public async Task EvalAsync_AskWithORespectsExisting()
	{
		var eval = new DxpFieldEval(new DxpFieldEvalDelegates
		{
			AskAsync = (prompt, ctx) => Task.FromResult<DxpFieldValue?>(new DxpFieldValue("New"))
		});
		eval.Context.SetBookmark("Answer", "Existing");

		var asked = await eval.EvalAsync(new DxpFieldInstruction("ASK Answer \"Prompt\" \\o"));
		var result = await eval.EvalAsync(new DxpFieldInstruction("REF Answer"));

		Assert.Equal(string.Empty, asked.Text);
		Assert.Equal("Existing", result.Text);
	}

	private sealed class CustomResolver : DocxportNet.Fields.Resolution.IDxpFieldValueResolver
	{
		public Task<DxpFieldValue?> ResolveAsync(string name, DocxportNet.Fields.Resolution.DxpFieldValueKindHint kind, DxpFieldEvalContext context)
		{
			if (name == "Y")
				return Task.FromResult<DxpFieldValue?>(new DxpFieldValue(9));
			return Task.FromResult<DxpFieldValue?>(null);
		}
	}

	private sealed class MockRefResolver : DocxportNet.Fields.Resolution.IDxpRefResolver
	{
		public Task<DocxportNet.Fields.Resolution.DxpRefResult?> ResolveAsync(DocxportNet.Fields.Resolution.DxpRefRequest request, DxpFieldEvalContext context)
		{
			if (request.Bookmark == "Note1" && request.Footnote)
			{
				return Task.FromResult<DocxportNet.Fields.Resolution.DxpRefResult?>(
					new DocxportNet.Fields.Resolution.DxpRefResult("[1]", FootnoteText: "Footnote text", FootnoteMark: "1"));
			}
			if (request.Bookmark == "Section")
			{
				return Task.FromResult<DocxportNet.Fields.Resolution.DxpRefResult?>(
					new DocxportNet.Fields.Resolution.DxpRefResult("Section 1.01"));
			}
			if (request.Bookmark == "Link" && request.Hyperlink)
			{
				return Task.FromResult<DocxportNet.Fields.Resolution.DxpRefResult?>(
					new DocxportNet.Fields.Resolution.DxpRefResult("LinkText", HyperlinkTarget: "#target"));
			}
			return Task.FromResult<DocxportNet.Fields.Resolution.DxpRefResult?>(null);
		}
	}

	private sealed class MockTableResolver : DocxportNet.Fields.Resolution.IDxpTableValueResolver
	{
		private readonly Dictionary<string, IReadOnlyList<double>> _ranges = new(StringComparer.OrdinalIgnoreCase);
		private readonly Dictionary<DocxportNet.Fields.Resolution.DxpTableRangeDirection, IReadOnlyList<double>> _directions = new();

		public MockTableResolver Add(string range, params double[] values)
		{
			_ranges[range] = values;
			return this;
		}

		public MockTableResolver AddDirection(DocxportNet.Fields.Resolution.DxpTableRangeDirection direction, params double[] values)
		{
			_directions[direction] = values;
			return this;
		}

		public Task<IReadOnlyList<double>> ResolveRangeAsync(string range, DxpFieldEvalContext context)
		{
			if (_ranges.TryGetValue(range, out var values))
				return Task.FromResult(values);
			return Task.FromResult<IReadOnlyList<double>>(Array.Empty<double>());
		}

		public Task<IReadOnlyList<double>> ResolveDirectionalRangeAsync(DocxportNet.Fields.Resolution.DxpTableRangeDirection direction, DxpFieldEvalContext context)
		{
			if (_directions.TryGetValue(direction, out var values))
				return Task.FromResult(values);
			return Task.FromResult<IReadOnlyList<double>>(Array.Empty<double>());
		}
	}

	[Fact]
	public async Task EvalAsync_FormulaTrivialFunctions()
	{
		var eval = new DxpFieldEval();
		eval.Context.Culture = new CultureInfo("en-US");

		var avg = await eval.EvalAsync(new DxpFieldInstruction("= AVERAGE(2, 4, 6)"));
		var avgEmpty = await eval.EvalAsync(new DxpFieldInstruction("= AVERAGE()"));
		var count = await eval.EvalAsync(new DxpFieldInstruction("= COUNT(1, 2, 3, 4)"));
		var countEmpty = await eval.EvalAsync(new DxpFieldInstruction("= COUNT()"));
		var mod = await eval.EvalAsync(new DxpFieldInstruction("= MOD(7, 3)"));
		var modZero = await eval.EvalAsync(new DxpFieldInstruction("= MOD(7, 0)"));
		var intf = await eval.EvalAsync(new DxpFieldInstruction("= INT(3.9)"));
		var intNeg = await eval.EvalAsync(new DxpFieldInstruction("= INT(-3.1)"));
		var notf = await eval.EvalAsync(new DxpFieldInstruction("= NOT(0)"));
		var notf2 = await eval.EvalAsync(new DxpFieldInstruction("= NOT(5)"));
		var andf = await eval.EvalAsync(new DxpFieldInstruction("= AND(1, 2, 3)"));
		var andf2 = await eval.EvalAsync(new DxpFieldInstruction("= AND(1, 0, 3)"));
		var orf = await eval.EvalAsync(new DxpFieldInstruction("= OR(0, 0, 5)"));
		var orf2 = await eval.EvalAsync(new DxpFieldInstruction("= OR(0, 0, 0)"));
		var t = await eval.EvalAsync(new DxpFieldInstruction("= TRUE()"));
		var f = await eval.EvalAsync(new DxpFieldInstruction("= FALSE()"));

		Assert.Equal("4", avg.Text);
		Assert.Equal("0", avgEmpty.Text);
		Assert.Equal("4", count.Text);
		Assert.Equal("0", countEmpty.Text);
		Assert.Equal("1", mod.Text);
		Assert.Equal("0", modZero.Text);
		Assert.Equal("3", intf.Text);
		Assert.Equal("-4", intNeg.Text);
		Assert.Equal("1", notf.Text);
		Assert.Equal("0", notf2.Text);
		Assert.Equal("1", andf.Text);
		Assert.Equal("0", andf2.Text);
		Assert.Equal("1", orf.Text);
		Assert.Equal("0", orf2.Text);
		Assert.Equal("1", t.Text);
		Assert.Equal("0", f.Text);
	}

	[Fact]
	public async Task EvalAsync_FormulaIfAndTrueFunctions()
	{
		var eval = new DxpFieldEval();
		eval.Context.Culture = new CultureInfo("en-US");

		var ifTrue = await eval.EvalAsync(new DxpFieldInstruction("= IF(1, 10, 5)"));
		var ifFalse = await eval.EvalAsync(new DxpFieldInstruction("= IF(0, 10, 5)"));
		var trueZero = await eval.EvalAsync(new DxpFieldInstruction("= TRUE(0)"));
		var trueNonZero = await eval.EvalAsync(new DxpFieldInstruction("= TRUE(2)"));

		Assert.Equal("10", ifTrue.Text);
		Assert.Equal("5", ifFalse.Text);
		Assert.Equal("0", trueZero.Text);
		Assert.Equal("1", trueNonZero.Text);
	}

	[Fact]
	public async Task EvalAsync_FormulaDefinedFunction()
	{
		var eval = new DxpFieldEval();
		eval.Context.Culture = new CultureInfo("en-US");
		eval.Context.SetBookmark("A", "5");

		var defined = await eval.EvalAsync(new DxpFieldInstruction("= DEFINED(A)"));
		var undefined = await eval.EvalAsync(new DxpFieldInstruction("= DEFINED(Unknown)"));

		Assert.Equal("1", defined.Text);
		Assert.Equal("0", undefined.Text);
	}

	[Fact]
	public async Task EvalAsync_FormulaWithNestedField()
	{
		var eval = new DxpFieldEval();
		eval.Context.Culture = new CultureInfo("en-US");
		eval.Context.SetNow(() => new DateTimeOffset(2026, 2, 7, 0, 0, 0, TimeSpan.Zero));

		var result = await eval.EvalAsync(new DxpFieldInstruction("= { DATE \\@ \"yyyy\" } + 1"));

		Assert.Equal("2027", result.Text);
	}

	[Fact]
	public async Task EvalAsync_FormulaComparisonReturnsOneOrZero()
	{
		var eval = new DxpFieldEval();
		eval.Context.Culture = new CultureInfo("en-US");
		var result = await eval.EvalAsync(new DxpFieldInstruction("= 3 > 2"));
		var result2 = await eval.EvalAsync(new DxpFieldInstruction("= 2 > 3"));

		Assert.Equal("1", result.Text);
		Assert.Equal("0", result2.Text);
	}

	[Fact]
	public async Task EvalAsync_FormulaPrecedenceAndPercent()
	{
		var eval = new DxpFieldEval();
		eval.Context.Culture = new CultureInfo("en-US");

		var result = await eval.EvalAsync(new DxpFieldInstruction("= 2 + 3 * 4"));
		var result2 = await eval.EvalAsync(new DxpFieldInstruction("= (2 + 3) * 4"));
		var result3 = await eval.EvalAsync(new DxpFieldInstruction("= 50% + 25"));

		Assert.Equal("14", result.Text);
		Assert.Equal("20", result2.Text);
		Assert.Equal("25.5", result3.Text);
	}

	[Fact]
	public async Task EvalAsync_FormulaUnaryMinusAndPower()
	{
		var eval = new DxpFieldEval();
		eval.Context.Culture = new CultureInfo("en-US");

		var result = await eval.EvalAsync(new DxpFieldInstruction("= -2^2"));
		var result2 = await eval.EvalAsync(new DxpFieldInstruction("= (-2)^2"));

		Assert.Equal("4", result.Text);
		Assert.Equal("4", result2.Text);
	}

	[Fact]
	public async Task EvalAsync_FormulaUnaryOddities()
	{
		var eval = new DxpFieldEval();
		eval.Context.Culture = new CultureInfo("en-US");

		var r1 = await eval.EvalAsync(new DxpFieldInstruction("= 0-2^2"));
		var r2 = await eval.EvalAsync(new DxpFieldInstruction("= -2^2"));
		var r3 = await eval.EvalAsync(new DxpFieldInstruction("= -(2^2)"));
		var r4 = await eval.EvalAsync(new DxpFieldInstruction("= 0+-2^2"));
		var r5 = await eval.EvalAsync(new DxpFieldInstruction("= 2++++++++2"));
		var r6 = await eval.EvalAsync(new DxpFieldInstruction("= 2------2"));

		Assert.Equal("-4", r1.Text);
		Assert.Equal("4", r2.Text);
		Assert.Equal("-4", r3.Text);
		Assert.Equal("4", r4.Text);
		Assert.Equal("4", r5.Text);
		Assert.Equal("4", r6.Text);
	}

	[Fact]
	public async Task EvalAsync_CompareReturnsOneOrZero()
	{
		var eval = new DxpFieldEval();
		eval.Context.SetBookmark("Value", "5");

		var result = await eval.EvalAsync(new DxpFieldInstruction("COMPARE Value >= 5"));
		var result2 = await eval.EvalAsync(new DxpFieldInstruction("COMPARE Value < 5"));

		Assert.Equal("1", result.Text);
		Assert.Equal("0", result2.Text);
	}

	[Fact]
	public async Task EvalAsync_SkipIfAndNextIfReturnSkippedStatus()
	{
		var eval = new DxpFieldEval();
		eval.Context.SetBookmark("Value", "5");

		var skip = await eval.EvalAsync(new DxpFieldInstruction("SKIPIF Value >= 5"));
		var next = await eval.EvalAsync(new DxpFieldInstruction("NEXTIF Value >= 5"));
		var skipFalse = await eval.EvalAsync(new DxpFieldInstruction("SKIPIF Value < 5"));

		Assert.Equal(DxpFieldEvalStatus.Skipped, skip.Status);
		Assert.Equal(DxpFieldEvalStatus.Skipped, next.Status);
		Assert.Equal(DxpFieldEvalStatus.Resolved, skipFalse.Status);
		Assert.Equal(string.Empty, skipFalse.Text);
	}

	[Fact]
	public async Task EvalAsync_SeqIncrementsAndResets()
	{
		var eval = new DxpFieldEval();

		var first = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure"));
		var second = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure"));
		var repeat = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure \\c"));

		eval.Context.SetBookmark("Start", "10");
		var reset = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure Start"));
		var afterReset = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure"));

		Assert.Equal("1", first.Text);
		Assert.Equal("2", second.Text);
		Assert.Equal("2", repeat.Text);
		Assert.Equal("10", reset.Text);
		Assert.Equal("11", afterReset.Text);
	}

	[Fact]
	public async Task EvalAsync_SeqHiddenReturnsEmpty()
	{
		var eval = new DxpFieldEval();

		var first = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure"));
		var hidden = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure \\h"));
		var next = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure"));

		Assert.Equal("1", first.Text);
		Assert.Equal(string.Empty, hidden.Text);
		Assert.Equal("3", next.Text);
	}

	[Fact]
	public async Task EvalAsync_SeqResetAndHideRespectStar()
	{
		var eval = new DxpFieldEval();

		var reset = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure \\r 3"));
		var hidden = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure \\h"));
		var notHidden = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure \\h \\* Arabic"));

		Assert.Equal("3", reset.Text);
		Assert.Equal(string.Empty, hidden.Text);
		Assert.Equal("5", notHidden.Text);
	}

	[Fact]
	public async Task EvalAsync_SeqResetsPerHeadingLevel()
	{
		var eval = new DxpFieldEval();
		int outline = 1;
		eval.Context.CurrentOutlineLevelProvider = () => outline;

		var first = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure \\s 1"));
		var second = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure \\s 1"));

		outline = 2;
		var reset = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure \\s 1"));
		var next = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure \\s 1"));

		Assert.Equal("1", first.Text);
		Assert.Equal("2", second.Text);
		Assert.Equal("1", reset.Text);
		Assert.Equal("2", next.Text);
	}

	[Fact]
	public async Task EvalAsync_FormulaResolvesTableCellRanges()
	{
		var eval = new DxpFieldEval();
		eval.Context.Culture = new CultureInfo("en-US");
		eval.Context.TableResolver = new MockTableResolver()
			.Add("A1", 2)
			.Add("B1", 3)
			.Add("A1:B1", 2, 3)
			.AddDirection(DocxportNet.Fields.Resolution.DxpTableRangeDirection.Left, 5, 7);

		var cell = await eval.EvalAsync(new DxpFieldInstruction("= A1 + B1"));
		var range = await eval.EvalAsync(new DxpFieldInstruction("= SUM(A1:B1)"));
		var dir = await eval.EvalAsync(new DxpFieldInstruction("= SUM(LEFT)"));

		Assert.Equal("5", cell.Text);
		Assert.Equal("5", range.Text);
		Assert.Equal("12", dir.Text);
	}

	[Fact]
	public void Context_AllowsCaseInsensitiveDocVariableLookup()
	{
		var ctx = new DxpFieldEvalContext();
		ctx.SetDocVariable("Answer", "42");

		Assert.True(ctx.TryGetDocVariable("answer", out var value));
		Assert.Equal("42", value);
	}

	[Fact]
	public void Context_AllowsCaseInsensitiveDocumentPropertyLookup()
	{
		var ctx = new DxpFieldEvalContext();
		ctx.SetDocumentProperty("Title", "Doc");

		Assert.True(ctx.TryGetDocumentProperty("title", out var value));
		Assert.Equal("Doc", value);
	}

	[Fact]
	public void Context_AllowsCaseInsensitiveBookmarkLookup()
	{
		var ctx = new DxpFieldEvalContext();
		ctx.SetBookmark("TotalCost", "123.45");

		Assert.True(ctx.TryGetBookmark("totalcost", out var value));
		Assert.Equal("123.45", value);
	}

	[Fact]
	public void Context_SetNow_ThrowsOnNullProvider()
	{
		var ctx = new DxpFieldEvalContext();

		Assert.Throws<ArgumentNullException>(() => ctx.SetNow(null!));
	}
}
