using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Visitors;
using DocxportNet.Visitors.PlainText;
using DocxportNet.Walker;
using System.Globalization;
using System.Xml.Linq;
using Xunit.Abstractions;
using DocxportNet.Tests.Utils;
using Xunit.Sdk;

namespace DocxportNet.Tests;

public class FieldEvalTests : TestBase<FieldEvalTests>
{
    public FieldEvalTests(ITestOutputHelper output) : base(output)
    {
    }

    [Fact]
    public async Task EvalAsync_UsesCache_WhenAllowed()
    {
        var eval = new DxpFieldEval(options: new DxpFieldEvalOptions { UseCacheOnNull = true }, logger: Logger);
        var instruction = new DxpFieldInstruction(" FOO ", "cached");

        var result = await eval.EvalAsync(instruction);

        Assert.Equal(DxpFieldEvalStatus.UsedCache, result.Status);
        Assert.Equal("cached", result.Text);
        Assert.Null(result.Error);
    }

    [Fact]
    public async Task EvalAsync_Skips_WhenNoCacheAndUnsupported()
    {
        var eval = new DxpFieldEval(options: new DxpFieldEvalOptions { UseCacheOnNull = true, ErrorOnUnsupported = false }, logger: Logger);
        var instruction = new DxpFieldInstruction("FOO", null);

        var result = await eval.EvalAsync(instruction);

        Assert.Equal(DxpFieldEvalStatus.Skipped, result.Status);
        Assert.Null(result.Text);
        Assert.Null(result.Error);
    }

    [Fact]
    public async Task EvalAsync_Fails_WhenConfiguredToErrorOnUnsupported()
    {
        var eval = new DxpFieldEval(options: new DxpFieldEvalOptions { ErrorOnUnsupported = true }, logger: Logger);
        var instruction = new DxpFieldInstruction("FOO", "cached");

        var result = await eval.EvalAsync(instruction);

        Assert.Equal(DxpFieldEvalStatus.Failed, result.Status);
        Assert.Null(result.Text);
        Assert.IsType<NotSupportedException>(result.Error);
    }

    [Fact]
    public async Task EvalAsync_DateAndTimeUseNowProvider()
    {
        var eval = new DxpFieldEval(logger: Logger);
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
        var eval = new DxpFieldEval(logger: Logger);
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
        var eval = new DxpFieldEval(logger: Logger);
        var now = new DateTimeOffset(2026, 2, 7, 8, 0, 0, TimeSpan.Zero);
        eval.Context.SetNow(() => now);

        var result = await eval.EvalAsync(new DxpFieldInstruction("PRINTDATE \\@ \"yyyy-MM-dd\""));

        Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
        Assert.Equal("2026-02-07", result.Text);
    }

    [Fact]
    public async Task EvalAsync_IfNumericComparison()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.SetBookmark("Order", "120");

        var result = await eval.EvalAsync(new DxpFieldInstruction("IF Order >= 100 \"Thanks\" \"No\""));

        Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
        Assert.Equal("Thanks", result.Text);
    }

    [Fact]
    public async Task EvalAsync_IfStringComparisonAndWildcard()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.SetDocVariable("Status", "Approved");

        var result = await eval.EvalAsync(new DxpFieldInstruction("IF Status = \"App*\" \"Yes\" \"No\""));

        Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
        Assert.Equal("Yes", result.Text);
    }

    [Fact]
    public async Task EvalAsync_IfNestedFieldComparison()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.CreatedDate = new DateTimeOffset(2026, 2, 7, 0, 0, 0, TimeSpan.Zero);

        var result = await eval.EvalAsync(new DxpFieldInstruction("IF { CREATEDATE \\@ \"yyyy\" } = \"2026\" \"Y\" \"N\""));

        Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
        Assert.Equal("Y", result.Text);
    }

    [Fact]
    public async Task EvalAsync_IfNestedFieldInResultText()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.SetNow(() => new DateTimeOffset(2026, 2, 7, 9, 0, 0, TimeSpan.Zero));

        var result = await eval.EvalAsync(new DxpFieldInstruction("IF 1 = 1 \"Today: { DATE \\@ \\\"yyyy-MM-dd\\\" }\" \"No\""));

        Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
        Assert.Equal("Today: 2026-02-07", result.Text);
    }

    [Fact]
    public async Task EvalAsync_FormulaArithmeticAndFunctions()
    {
        var eval = new DxpFieldEval(logger: Logger);
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
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("fr-FR");
        eval.Context.ListSeparator = ";";

        var result = await eval.EvalAsync(new DxpFieldInstruction("= SUM(1; 2; 3)"));

        Assert.Equal("6", result.Text);
    }

    [Fact]
    public async Task EvalAsync_FormulaCustomFunction()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("en-US");
        eval.Context.FormulaFunctions.Register("DOUBLE", args => args.Count > 0 ? args[0] * 2 : 0);

        var result = await eval.EvalAsync(new DxpFieldInstruction("= DOUBLE(4)"));

        Assert.Equal("8", result.Text);
    }

    [Fact]
    public async Task EvalAsync_ResolvesVariableViaDelegate()
    {
        var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
            ResolveDocVariableAsync = (name, ctx) => Task.FromResult<DxpFieldValue?>(name == "X" ? new DxpFieldValue(5) : null)
        }, logger: Logger);
        eval.Context.Culture = new CultureInfo("en-US");

        var result = await eval.EvalAsync(new DxpFieldInstruction("= X + 1"));

        Assert.Equal("6", result.Text);
    }

    [Fact]
    public async Task EvalAsync_ResolvesVariableViaCustomResolver()
    {
        var eval = new DxpFieldEval(logger: Logger);
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
        var eval = new DxpFieldEval(logger: Logger);

        var set = await eval.EvalAsync(new DxpFieldInstruction("SET Total \"42\""));
        var get = await eval.EvalAsync(new DxpFieldInstruction("REF Total"));

        Assert.Equal("42", set.Text);
        Assert.Equal("42", get.Text);
    }

    [Fact]
    public async Task EvalAsync_DocPropertyUsesContextValue()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.SetDocumentPropertyValue("Title", new DxpFieldValue("Doc Title"));

        var result = await eval.EvalAsync(new DxpFieldInstruction("DOCPROPERTY Title"));

        Assert.Equal("Doc Title", result.Text);
    }

    [Fact]
    public async Task EvalAsync_DocPropertyExpandsNestedName()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.SetDocumentPropertyValue("Title", new DxpFieldValue("Doc Title"));
        eval.Context.SetBookmark("PropName", "Title");

        var result = await eval.EvalAsync(new DxpFieldInstruction("DOCPROPERTY { REF PropName }"));

        Assert.Equal("Doc Title", result.Text);
    }

    [Fact]
    public async Task EvalAsync_DocPropertyDateFormats()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.SetDocumentPropertyValue("CreateDate", new DxpFieldValue(new DateTimeOffset(2026, 2, 7, 0, 0, 0, TimeSpan.Zero)));

        var result = await eval.EvalAsync(new DxpFieldInstruction("DOCPROPERTY CreateDate \\@ \"yyyy-MM-dd\""));

        Assert.Equal("2026-02-07", result.Text);
    }

    [Fact]
    public async Task EvalAsync_DocVariableUsesResolver()
    {
        var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
            ResolveDocVariableAsync = (name, ctx) => Task.FromResult<DxpFieldValue?>(name == "X" ? new DxpFieldValue("ok") : null)
        }, logger: Logger);

        var result = await eval.EvalAsync(new DxpFieldInstruction("DOCVARIABLE X"));
        var missing = await eval.EvalAsync(new DxpFieldInstruction("DOCVARIABLE Missing"));

        Assert.Equal("ok", result.Text);
        Assert.Equal("Error! No document variable supplied.", missing.Text);
    }

    [Fact]
    public async Task EvalAsync_DocVariableExpandsNestedName()
    {
        var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
            ResolveDocVariableAsync = (name, ctx) => Task.FromResult<DxpFieldValue?>(name == "X" ? new DxpFieldValue("ok") : null)
        }, logger: Logger);
        eval.Context.SetBookmark("VarName", "X");

        var result = await eval.EvalAsync(new DxpFieldInstruction("DOCVARIABLE { REF VarName }"));

        Assert.Equal("ok", result.Text);
    }

    [Fact]
    public void Walker_EvalMode_SetSuppressesOutputAndSetsBookmark()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect 1: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SET Var1 "1" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF Var1 </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(bodyXml));
        var expected = TestCompare.Normalize("Expect 1: 1\n\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public async Task EvalAsync_MergeFieldUsesResolverAndSwitches()
    {
        var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
            ResolveMergeFieldAsync = (name, ctx) => Task.FromResult<DxpFieldValue?>(name == "FirstName" ? new DxpFieldValue("Ana") : null)
        }, logger: Logger);

        var result = await eval.EvalAsync(new DxpFieldInstruction("MERGEFIELD FirstName \\b \"Hello \" \\f \"!\""));
        var missing = await eval.EvalAsync(new DxpFieldInstruction("MERGEFIELD Missing \\b \"Hello \" \\f \"!\""));

        Assert.Equal("Hello Ana!", result.Text);
        Assert.Equal(string.Empty, missing.Text);
    }

    [Fact]
    public async Task EvalAsync_MergeFieldExpandsNestedName()
    {
        var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
            ResolveMergeFieldAsync = (name, ctx) => Task.FromResult<DxpFieldValue?>(name == "FirstName" ? new DxpFieldValue("Ana") : null)
        }, logger: Logger);
        eval.Context.SetBookmark("FieldName", "FirstName");

        var result = await eval.EvalAsync(new DxpFieldInstruction("MERGEFIELD { REF FieldName }"));

        Assert.Equal("Ana", result.Text);
    }

    [Fact]
    public async Task EvalAsync_MergeFieldMapsWithM()
    {
        var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
            ResolveMergeFieldAsync = (name, ctx) => Task.FromResult<DxpFieldValue?>(name == "GivenName" ? new DxpFieldValue("Ana") : null)
        }, logger: Logger);
        eval.Context.SetMergeFieldAlias("FirstName", "GivenName");

        var result = await eval.EvalAsync(new DxpFieldInstruction("MERGEFIELD FirstName \\m"));

        Assert.Equal("Ana", result.Text);
    }

    [Fact]
    public async Task EvalAsync_RefUsesResolverAndSwitches()
    {
        var eval = new DxpFieldEval(logger: Logger);
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
    public async Task EvalAsync_RefExpandsNestedBookmarkName()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.SetBookmark("TargetName", "Note1");
        eval.Context.RefResolver = new MockRefResolver();

        var result = await eval.EvalAsync(new DxpFieldInstruction("REF { REF TargetName } \\f"));

        Assert.Equal("[1]", result.Text);
        Assert.Single(eval.Context.RefFootnotes);
    }

    [Fact]
    public async Task EvalAsync_AskSetsBookmarkAndUsesDefault()
    {
        var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
            AskAsync = (prompt, ctx) => Task.FromResult<DxpFieldValue?>(prompt == "Name?" ? new DxpFieldValue("Ana") : null)
        }, logger: Logger);

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
        var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
            AskAsync = (prompt, ctx) => Task.FromResult<DxpFieldValue?>(new DxpFieldValue("New"))
        }, logger: Logger);
        eval.Context.SetBookmark("Answer", "Existing");

        var asked = await eval.EvalAsync(new DxpFieldInstruction("ASK Answer \"Prompt\" \\o"));
        var result = await eval.EvalAsync(new DxpFieldInstruction("REF Answer"));

        Assert.Equal(string.Empty, asked.Text);
        Assert.Equal("Existing", result.Text);
    }

    [Fact]
    public async Task EvalAsync_AskExpandsNestedPromptAndDefault()
    {
        string? capturedPrompt = null;
        var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
            AskAsync = (prompt, ctx) => {
                capturedPrompt = prompt;
                return Task.FromResult<DxpFieldValue?>(null);
            }
        }, logger: Logger);
        eval.Context.SetBookmark("Greeting", "Hi");
        eval.Context.SetBookmark("DefaultCity", "Rome");

        var asked = await eval.EvalAsync(new DxpFieldInstruction("ASK City \"{ REF Greeting } there?\" \\d \"{ REF DefaultCity }\""));
        var city = await eval.EvalAsync(new DxpFieldInstruction("REF City"));

        Assert.Equal(string.Empty, asked.Text);
        Assert.Equal("Hi there?", capturedPrompt);
        Assert.Equal("Rome", city.Text);
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
        private readonly Dictionary<DocxportNet.Fields.Resolution.DxpTableRangeDirection, IReadOnlyList<double>> _directions = [];

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
            return Task.FromResult<IReadOnlyList<double>>([]);
        }

        public Task<IReadOnlyList<double>> ResolveDirectionalRangeAsync(DocxportNet.Fields.Resolution.DxpTableRangeDirection direction, DxpFieldEvalContext context)
        {
            if (_directions.TryGetValue(direction, out var values))
                return Task.FromResult(values);
            return Task.FromResult<IReadOnlyList<double>>([]);
        }
    }

    [Fact]
    public async Task EvalAsync_FormulaTrivialFunctions()
    {
        var eval = new DxpFieldEval(logger: Logger);
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
        var eval = new DxpFieldEval(logger: Logger);
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
        var eval = new DxpFieldEval(logger: Logger);
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
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("en-US");
        eval.Context.SetNow(() => new DateTimeOffset(2026, 2, 7, 0, 0, 0, TimeSpan.Zero));

        var result = await eval.EvalAsync(new DxpFieldInstruction("= { DATE \\@ \"yyyy\" } + 1"));

        Assert.Equal("2027", result.Text);
    }

    [Fact]
    public async Task EvalAsync_FormulaComparisonReturnsOneOrZero()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("en-US");
        var result = await eval.EvalAsync(new DxpFieldInstruction("= 3 > 2"));
        var result2 = await eval.EvalAsync(new DxpFieldInstruction("= 2 > 3"));

        Assert.Equal("1", result.Text);
        Assert.Equal("0", result2.Text);
    }

    [Fact]
    public async Task EvalAsync_FormulaPrecedenceAndPercent()
    {
        var eval = new DxpFieldEval(logger: Logger);
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
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("en-US");

        var result = await eval.EvalAsync(new DxpFieldInstruction("= -2^2"));
        var result2 = await eval.EvalAsync(new DxpFieldInstruction("= (-2)^2"));

        Assert.Equal("4", result.Text);
        Assert.Equal("4", result2.Text);
    }

    [Fact]
    public async Task EvalAsync_FormulaUnaryOddities()
    {
        var eval = new DxpFieldEval(logger: Logger);
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
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.SetBookmark("Value", "5");

        var result = await eval.EvalAsync(new DxpFieldInstruction("COMPARE Value >= 5"));
        var result2 = await eval.EvalAsync(new DxpFieldInstruction("COMPARE Value < 5"));

        Assert.Equal("1", result.Text);
        Assert.Equal("0", result2.Text);
    }

    [Fact]
    public async Task EvalAsync_SkipIfAndNextIfReturnSkippedStatus()
    {
        var eval = new DxpFieldEval(logger: Logger);
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
        var eval = new DxpFieldEval(logger: Logger);

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
    public async Task EvalAsync_SeqExpandsNestedIdentifierAndBookmark()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.SetBookmark("SeqName", "Figure");
        eval.Context.SetBookmark("ResetName", "Start");
        eval.Context.SetBookmark("Start", "10");

        var first = await eval.EvalAsync(new DxpFieldInstruction("SEQ { REF SeqName }"));
        var reset = await eval.EvalAsync(new DxpFieldInstruction("SEQ { REF SeqName } { REF ResetName }"));
        var next = await eval.EvalAsync(new DxpFieldInstruction("SEQ { REF SeqName }"));

        Assert.Equal("1", first.Text);
        Assert.Equal("10", reset.Text);
        Assert.Equal("11", next.Text);
    }

    [Fact]
    public async Task EvalAsync_SeqHiddenReturnsEmpty()
    {
        var eval = new DxpFieldEval(logger: Logger);

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
        var eval = new DxpFieldEval(logger: Logger);

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
        var eval = new DxpFieldEval(logger: Logger);
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
        var eval = new DxpFieldEval(logger: Logger);
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

    [Fact]
    public void Walker_CacheMode_UsesCachedResults_AndSuppressesSet()
    {
        var bodyXml = $@"
<w:body xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
        xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"">
<w:p w14:paraId=""7F9887E7"" w14:textId=""31B65863"">
  <w:r><w:t xml:space=""preserve"">Expect 1: </w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""begin""/></w:r>
  <w:r><w:instrText xml:space=""preserve""> SET Var1 ""1"" </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType=""separate""/></w:r>
  <w:r><w:t>1</w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""end""/></w:r>
  <w:fldSimple w:instr="" REF Var1 "">
    <w:r><w:t>1</w:t></w:r>
  </w:fldSimple>
</w:p>
<w:p w14:paraId=""4F8CA4F2"" w14:textId=""0FE5363B"">
  <w:r><w:t xml:space=""preserve"">Expect Error: </w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""begin""/></w:r>
  <w:r><w:instrText xml:space=""preserve""> REF VarUnknown </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType=""separate""/></w:r>
  <w:r><w:t>Error! Reference source not found.</w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""end""/></w:r>
</w:p>
<w:p w14:paraId=""1B477A12"" w14:textId=""504973EB"">
  <w:r><w:t xml:space=""preserve"">Expect No Error: </w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""begin""/></w:r>
  <w:r><w:instrText xml:space=""preserve""> IF </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType=""begin""/></w:r>
  <w:r><w:instrText xml:space=""preserve""> REF VarUnknow </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType=""separate""/></w:r>
  <w:r><w:instrText>Error! Reference source not found.</w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType=""end""/></w:r>
  <w:r><w:instrText xml:space=""preserve""> = """" ""Empty"" ""Not Empty"" </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType=""separate""/></w:r>
  <w:r><w:t>Not Empty</w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""end""/></w:r>
</w:p>
<w:p w14:paraId=""1D7DCB2B"" w14:textId=""25F027AF"">
  <w:r><w:t>Expect one:</w:t></w:r>
  <w:r><w:t xml:space=""preserve""> </w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""begin""/></w:r>
  <w:r><w:instrText xml:space=""preserve""> SET Var1 ""1"" </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType=""separate""/></w:r>
  <w:r><w:t>1</w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""end""/></w:r>
  <w:r><w:fldChar w:fldCharType=""begin""/></w:r>
  <w:r><w:instrText xml:space=""preserve""> IF </w:instrText></w:r>
  <w:fldSimple w:instr="" REF Var1 "">
    <w:r><w:instrText>1</w:instrText></w:r>
  </w:fldSimple>
  <w:r><w:instrText xml:space=""preserve""> = ""1"" ""one"" ""not one"" </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType=""separate""/></w:r>
  <w:r><w:t>one</w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""end""/></w:r>
</w:p>
<w:p w14:paraId=""381F356C"" w14:textId=""1386C4C3"">
  <w:r><w:t xml:space=""preserve"">Expect </w:t></w:r>
  <w:r><w:t>one</w:t></w:r>
  <w:r><w:t xml:space=""preserve""> (bold)</w:t></w:r>
  <w:r><w:t xml:space=""preserve"">: </w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""begin""/></w:r>
  <w:r><w:instrText xml:space=""preserve""> SET Var1 ""</w:instrText></w:r>
  <w:r><w:instrText>1</w:instrText></w:r>
  <w:r><w:instrText xml:space=""preserve"">"" </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType=""separate""/></w:r>
  <w:r><w:t>1</w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""end""/></w:r>
  <w:r><w:fldChar w:fldCharType=""begin""/></w:r>
  <w:r><w:instrText xml:space=""preserve""> IF </w:instrText></w:r>
  <w:fldSimple w:instr="" REF Var1 "">
    <w:r><w:instrText>1</w:instrText></w:r>
  </w:fldSimple>
  <w:r><w:instrText xml:space=""preserve""> = ""1"" ""</w:instrText></w:r>
  <w:r><w:instrText>one</w:instrText></w:r>
  <w:r><w:instrText>"" ""</w:instrText></w:r>
  <w:r><w:instrText>not one</w:instrText></w:r>
  <w:r><w:instrText xml:space=""preserve"">"" </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType=""separate""/></w:r>
  <w:r><w:t>one</w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""end""/></w:r>
</w:p>
<w:p w14:paraId=""7B90ECD8"" w14:textId=""271E59D9"">
  <w:r><w:t xml:space=""preserve"">Expect </w:t></w:r>
  <w:r><w:t>1</w:t></w:r>
  <w:r><w:t>2</w:t></w:r>
  <w:r><w:t>3</w:t></w:r>
  <w:r><w:t xml:space=""preserve"">: </w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""begin""/></w:r>
  <w:r><w:instrText xml:space=""preserve""> IF 1 = 1 ""</w:instrText></w:r>
  <w:r><w:instrText>1</w:instrText></w:r>
  <w:r><w:instrText>2</w:instrText></w:r>
  <w:r><w:instrText>3</w:instrText></w:r>
  <w:r><w:instrText xml:space=""preserve"">"" ""error"" </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType=""separate""/></w:r>
  <w:r><w:t>1</w:t></w:r>
  <w:r><w:t>2</w:t></w:r>
  <w:r><w:t>3</w:t></w:r>
  <w:r><w:fldChar w:fldCharType=""end""/></w:r>
</w:p>
</w:body>";

        var expected = TestCompare.Normalize(string.Join("\n\n", new[] {
            "Expect 1: 1",
            "Expect Error: Error! Reference source not found.",
            "Expect No Error: Not Empty",
            "Expect one: one",
            "Expect one (bold): one",
            "Expect 123: 123"
        }) + "\n\n");

        var actual = TestCompare.Normalize(ExportPlainTextCachedFromBodyXml(bodyXml));
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Walker_EvalMode_IfWithMissingRefUsesErrorTextAsValue()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p>
    <w:r>
      <w:t xml:space="preserve">Expect No Error: </w:t>
    </w:r>

    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> IF { REF VarUnknow } = "" "Empty" "Not Empty" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Not Empty</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(bodyXml));
        var expected = TestCompare.Normalize("Expect No Error: Not Empty\n\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Walker_TableDirectionalRanges_ResolveThroughMiddleware()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(BuildTestTable()));
            main.Document.Save();
        }

        stream.Position = 0;

        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("en-US");
        var collector = new TableFieldCollector(eval);
        var visitor = DxpVisitorMiddleware.Chain(
            collector,
            next => new DxpFieldEvalMiddleware(next, eval, logger: Logger),
            next => new DxpContextTracker(next));

        using (var readDoc = WordprocessingDocument.Open(stream, false))
            new DxpWalker(Logger).Accept(readDoc, visitor);

        Assert.Equal("12", collector.Results["= SUM(BELOW)"]);
        Assert.Equal("1", collector.Results["= SUM(ABOVE)"]);
        Assert.Equal("6", collector.Results["= SUM(LEFT)"]);
        Assert.Equal("9", collector.Results["= SUM(ABOVE) + 0"]);
        Assert.Equal("16", collector.Results["= SUM(RIGHT)"]);
    }

    [Fact]
    public void Walker_RefResolvesBookmarkParagraphFootnoteAndHyperlink()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            AddMinimalNumberingDefinitions(main, numId: 1, abstractNumId: 1);
            var content = BuildRefTestContent(numId: 1).ToList();
            main.Document = new Document(new Body(content));

            var footnotesPart = main.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes = new Footnotes(
                new Footnote(
                    new Paragraph(new Run(new Text("Footnote text")))) { Id = 1 });

            main.Document.Save();
        }

        stream.Position = 0;

        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("en-US");
        var collector = new RefFieldCollector(eval);
        var visitor = DxpVisitorMiddleware.Chain(
            collector,
            next => new DxpFieldEvalMiddleware(next, eval, logger: Logger),
            next => new DxpContextTracker(next));

        using (var readDoc = WordprocessingDocument.Open(stream, false))
        {
            new DxpWalker(Logger).Accept(readDoc, visitor);
        }

        Assert.Equal("Bookmark Textlink", collector.Results["REF BM1"]);
        Assert.Equal("1", collector.Results["REF BM1 \\n"]);
        Assert.Equal("1 above", collector.Results["REF BM1 \\n \\p"]);
        Assert.Equal("Footnote text", collector.Results["REF BM1 \\f"]);
        Assert.Equal("Bookmark Textlink", collector.Results["REF BM1 \\h"]);
        Assert.Contains(eval.Context.RefHyperlinks, link => link.Bookmark == "BM1" && link.Target == "BM1");
        Assert.Contains(eval.Context.RefFootnotes, note => note.Bookmark == "BM1" && note.Text == "Footnote text");
    }

    private static Table BuildTestTable()
    {
        return new Table(
            new TableRow(
                CellWithText("1"),
                CellWithFormula("= SUM(BELOW)", "12"),
                CellWithText("3")
            ),
            new TableRow(
                CellWithFormula("= SUM(ABOVE)", "1"),
                CellWithText("5"),
                CellWithFormula("= SUM(LEFT)", "6")
            ),
            new TableRow(
                CellWithFormula("= SUM(RIGHT)", "16"),
                CellWithText("7"),
                CellWithFormula("= SUM(ABOVE) + 0", "9")
            )
        );
    }

    private static IEnumerable<OpenXmlElement> BuildRefTestContent(int? numId = null)
    {
        var bookmarkStart = new BookmarkStart { Name = "BM1", Id = "1" };
        var bookmarkEnd = new BookmarkEnd { Id = "1" };

        ParagraphProperties? properties = null;
        if (numId.HasValue)
        {
            properties = new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = numId.Value }));
        }

        var paragraphElements = new List<OpenXmlElement> {
            new Run(new Text("1.")),
            bookmarkStart,
            new Run(new Text("Bookmark Text")),
            new Run(new FootnoteReference { Id = 1 }),
            new Hyperlink(new Run(new Text("link"))) { Anchor = "BM1" },
            bookmarkEnd
        };
        if (properties != null)
            paragraphElements.Insert(0, properties);

        yield return new Paragraph(paragraphElements);

        yield return new Paragraph(new Run(new Text("After")));

        yield return new Paragraph(new SimpleField { Instruction = "REF BM1" });
        yield return new Paragraph(new SimpleField { Instruction = "REF BM1 \\n" });
        yield return new Paragraph(new SimpleField { Instruction = "REF BM1 \\n \\p" });
        yield return new Paragraph(new SimpleField { Instruction = "REF BM1 \\f" });
        yield return new Paragraph(new SimpleField { Instruction = "REF BM1 \\h" });
    }

    private static void AddMinimalNumberingDefinitions(MainDocumentPart main, int numId, int abstractNumId)
    {
        var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering(
            new AbstractNum(
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." }
                ) { LevelIndex = 0 }
            ) { AbstractNumberId = abstractNumId },
            new NumberingInstance(new AbstractNumId { Val = abstractNumId }) { NumberID = numId }
        );
    }

    private static TableCell CellWithText(string text)
    {
        return new TableCell(new Paragraph(new Run(new Text(text))));
    }

    private static TableCell CellWithFormula(string instruction, string cachedText)
    {
        var fld = new SimpleField { Instruction = instruction };
        fld.Append(new Run(new Text(cachedText)));
        return new TableCell(new Paragraph(fld));
    }

    private string ExportPlainTextCachedFromBodyXml(string bodyXml)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var xml = System.Xml.Linq.XDocument.Parse(bodyXml);
            var body = new Body();
            body.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            body.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            body.InnerXml = string.Concat(xml.Root!.Nodes());
            main.Document = new Document(body);
            main.Document.Save();
        }

        stream.Position = 0;

        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        if (visitor is not IDxpFieldEvalProvider provider)
            throw new XunitException("DxpPlainTextVisitor should provide field evaluation context.");

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => new DxpFieldEvalMiddleware(next, provider.FieldEval, DxpFieldEvalMode.Cache, logger: Logger),
            next => new DxpContextTracker(next));

        using (var readDoc = WordprocessingDocument.Open(stream, false))
            new DxpWalker(Logger).Accept(readDoc, pipeline);

        return writer.ToString();
    }

    private string ExportPlainTextEvaluatedFromBodyXml(string bodyXml)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var xml = System.Xml.Linq.XDocument.Parse(bodyXml);
            var body = new Body();
            body.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            body.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            body.InnerXml = string.Concat(xml.Root!.Nodes());
            main.Document = new Document(body);
            main.Document.Save();
        }

        stream.Position = 0;

        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        if (visitor is not IDxpFieldEvalProvider provider)
            throw new XunitException("DxpPlainTextVisitor should provide field evaluation context.");

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => new DxpFieldEvalMiddleware(next, provider.FieldEval, DxpFieldEvalMode.Evaluate, logger: Logger),
            next => new DxpContextTracker(next));

        using (var readDoc = WordprocessingDocument.Open(stream, false))
            new DxpWalker(Logger).Accept(readDoc, pipeline);

        return writer.ToString();
    }

    private sealed class TableFieldCollector : DxpVisitor
    {
        public Dictionary<string, string?> Results { get; } = new(StringComparer.OrdinalIgnoreCase);

        public TableFieldCollector(DxpFieldEval eval) : base(null)
        {
        }

        public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
        {
            var instruction = d.CurrentFields.Current?.InstructionText;
            if (string.IsNullOrWhiteSpace(instruction))
                return;

            var key = instruction.Trim();
            if (Results.TryGetValue(key, out var existing))
                Results[key] = (existing ?? string.Empty) + text;
            else
                Results[key] = text;
        }
    }

    private sealed class RefFieldCollector : DxpVisitor
    {
        public Dictionary<string, string?> Results { get; } = new(StringComparer.OrdinalIgnoreCase);

        public RefFieldCollector(DxpFieldEval eval) : base(null)
        {
        }

        public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
        {
            var instruction = d.CurrentFields.Current?.InstructionText;
            if (string.IsNullOrWhiteSpace(instruction))
                return;

            var key = instruction.Trim();
            if (Results.TryGetValue(key, out var existing))
                Results[key] = (existing ?? string.Empty) + text;
            else
                Results[key] = text;
        }
    }
}
