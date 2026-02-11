using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Visitors;
using DocxportNet.Visitors.PlainText;
using DocxportNet.Walker;
using System.Globalization;
using System.Threading;
using Xunit.Abstractions;
using DocxportNet.Tests.Utils;
using Xunit.Sdk;
using System.Text;
using DocxportNet.Core;
using Microsoft.Extensions.Logging;
using DocxportNet.Fields.Resolution;
using DocxportNet.Fields.Resolution.Impl;
using DocxportNet.Middleware;
using DocxportNet.Fields.Eval;
using DocxportNet.Visitors.Html;
using DocumentFormat.OpenXml.CustomProperties;
using ExtendedProperties = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.VariantTypes;
using DocxportNet;

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
        eval.Context.SetBookmarkNodes("Order", DxpFieldNodeBuffer.FromText("120"));

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
        eval.Context.SetBookmarkNodes("A", DxpFieldNodeBuffer.FromText("10"));
        eval.Context.SetBookmarkNodes("B", DxpFieldNodeBuffer.FromText("5"));

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
    public async Task EvalAsync_FormulaErrorsReturnErrorText()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("en-US");

        var divide = await eval.EvalAsync(new DxpFieldInstruction("= 1 / 0"));
        var unknown = await eval.EvalAsync(new DxpFieldInstruction("= LN(10)"));

        Assert.Equal("!Zero Divide", divide.Text);
        Assert.Equal("!Syntax Error, 10", unknown.Text);
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
        eval.Context.ValueResolver = new DxpChainedFieldValueResolver(
            new DxpContextFieldValueResolver(),
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
        eval.Context.SetBookmarkNodes("PropName", DxpFieldNodeBuffer.FromText("Title"));

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
    public void Walker_EvalMode_DocProperty_ResolvesBuiltInAndCustom()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Title: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> DOCPROPERTY Title </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Created: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> DOCPROPERTY Created \\@ "yyyy-MM-dd" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Custom: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> DOCPROPERTY "CustomProp1" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var eval = new DxpFieldEval(logger: Logger);
        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(
            bodyXml,
            eval,
            includeCustomProperties: true,
            configureDocument: doc => {
                doc.PackageProperties.Title = "My Title";
                doc.PackageProperties.Created = new DateTime(2026, 2, 7, 12, 0, 0, DateTimeKind.Utc);

                var customPart = doc.CustomFilePropertiesPart ?? doc.AddCustomFilePropertiesPart();
                customPart.Properties ??= new Properties();

                CustomDocumentProperty? existing = null;
                foreach (var prop in customPart.Properties.Elements<CustomDocumentProperty>())
                {
                    if (string.Equals(prop.Name?.Value, "CustomProp1", StringComparison.Ordinal))
                    {
                        existing = prop;
                        break;
                    }
                }

                var custom = existing ?? new CustomDocumentProperty {
                    Name = "CustomProp1",
                    PropertyId = 2,
                    FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
                };
                custom.VTLPWSTR = new VTLPWSTR("Custom Value");
                if (existing == null)
                    customPart.Properties.AppendChild(custom);
                customPart.Properties.Save();
            }));
        Assert.Contains("Title: My Title", actual, StringComparison.Ordinal);
        Assert.Contains("Created: 2026-02-07", actual, StringComparison.Ordinal);
        Assert.Contains("Custom: Custom Value", actual, StringComparison.Ordinal);
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
        eval.Context.SetBookmarkNodes("VarName", DxpFieldNodeBuffer.FromText("X"));

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
    public void Walker_EvalMode_SetWithNestedFieldValue_SetsBookmark()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Value: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SET Var1 "{ DATE \\@ yyyy }" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF Var1 </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.SetNow(() => new DateTimeOffset(2026, 2, 7, 0, 0, 0, TimeSpan.Zero));

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(bodyXml, eval));
        var expected = TestCompare.Normalize("Value: 2026\n\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Walker_EvalMode_ValueFields_EmitStructuredResults()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect: </w:t></w:r>
    <w:r><w:t xml:space="preserve">Title=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> DOCPROPERTY Title </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Merge=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> MERGEFIELD FirstName \\* Upper </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Seq=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SEQ Figure </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Date=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> DATE \\@ "yyyy-MM-dd" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Compare=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> COMPARE 5 > 3 </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Formula=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> = 2 + 3 </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
            ResolveMergeFieldAsync = (name, _) => Task.FromResult<DxpFieldValue?>(name == "FirstName" ? new DxpFieldValue("Ana") : null)
        }, logger: Logger);
        eval.Context.SetDocumentPropertyValue("Title", new DxpFieldValue("Doc"));
        eval.Context.SetNow(() => new DateTimeOffset(2026, 2, 7, 0, 0, 0, TimeSpan.Zero));
        eval.Context.Culture = new CultureInfo("en-US");

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(bodyXml, eval));
        var expected = TestCompare.Normalize("Expect: Title=Doc Merge=ANA Seq=1 Date=2026-02-07 Compare=1 Formula=5\n\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Walker_EvalMode_AskAndSkipIfSuppressOutput()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Answer: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> ASK Name "Name?" \\d "Unknown" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF Name </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Skip=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SKIPIF 1 = 1 </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>should-not-appear</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
            AskAsync = (prompt, _) => Task.FromResult<DxpFieldValue?>(prompt == "Name?" ? new DxpFieldValue("Ana") : null)
        }, logger: Logger);

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(bodyXml, eval));
        var expected = TestCompare.Normalize("Answer: Ana Skip=\n\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Walker_EvalMode_RefWithoutSwitch_ReplaysStructuredBookmark()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect one: </w:t></w:r>
    <w:bookmarkStart w:id="0" w:name="BM1"/>
    <w:del w:id="1" w:author="test">
      <w:r><w:rPr><w:b/></w:rPr><w:t>one</w:t></w:r>
    </w:del>
    <w:bookmarkEnd w:id="0"/>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF BM1 </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = ExportPlainTextEvaluatedFromBodyXml(bodyXml);
        var expected = TestCompare.Normalize("Expect one: one\n\n");
        Assert.Equal(expected, TestCompare.Normalize(actual));
    }

    [Fact]
    public void Walker_EvalMode_RefWithSwitch_FormatsFlattenedBookmark()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect 12: </w:t></w:r>
    <w:bookmarkStart w:id="0" w:name="BM2"/>
    <w:del w:id="1" w:author="test">
      <w:r><w:rPr><w:b/></w:rPr><w:t>1</w:t></w:r>
      <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>2</w:t></w:r>
    </w:del>
    <w:bookmarkEnd w:id="0"/>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF BM2 \\# "00" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = ExportPlainTextEvaluatedFromBodyXml(bodyXml);
        var expected = TestCompare.Normalize("Expect 12: 12\n\n");
        Assert.Equal(expected, TestCompare.Normalize(actual));
    }

    [Fact]
    public void Walker_EvalMode_RefWithoutSwitch_ReplaysStructuredBookmarkRuns()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect: </w:t></w:r>
    <w:bookmarkStart w:id="0" w:name="BM3"/>
    <w:r><w:rPr><w:b/></w:rPr><w:t>one</w:t></w:r>
    <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>two</w:t></w:r>
    <w:r><w:t>three</w:t></w:r>
    <w:bookmarkEnd w:id="0"/>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF BM3 </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = TestCompare.Normalize(ExportRunMarkupEvaluatedFromBodyXml(bodyXml));
        var expected = TestCompare.Normalize("Expect: <b>one</b><u>two</u>three<b>one</b><u>two</u>three\n\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Walker_EvalMode_RefWithFormatSwitch_FlattensBookmarkContent()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect: </w:t></w:r>
    <w:bookmarkStart w:id="0" w:name="BM3"/>
    <w:r><w:rPr><w:b/></w:rPr><w:t>one</w:t></w:r>
    <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>two</w:t></w:r>
    <w:r><w:t>three</w:t></w:r>
    <w:bookmarkEnd w:id="0"/>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF BM3 \\* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = TestCompare.Normalize(ExportRunMarkupEvaluatedFromBodyXml(bodyXml));
        var expected = TestCompare.Normalize("Expect: <b>one</b><u>two</u>threeonetwothree\n\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Walker_EvalMode_RefSwitches_EmitExpectedTextAndSideEffects()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Refs: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF Section \\n </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve">, </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF Section \\r </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve">, </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF Section \\w </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve">, </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF Section \\t </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve">, </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF Section \\d "-" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve">, </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF Section \\n \\p </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve">, </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF Note1 \\f </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve">, </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF Link \\h </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.RefResolver = new MockRefResolver();

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(bodyXml, eval));
        var expected = TestCompare.Normalize("Refs: 1.01, 1.01, Section 1.01, 1.01, Section 1-01, 1.01 below, [1], LinkText\n\n");
        Assert.Equal(expected, actual);
        Assert.Single(eval.Context.RefFootnotes);
        Assert.Single(eval.Context.RefHyperlinks);
    }

    [Fact]
    public void Walker_EvalMode_RefWithCharformat_UsesFieldCodeRunStyle()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect bold: </w:t></w:r>
    <w:bookmarkStart w:id="0" w:name="BM1"/>
    <w:del w:id="1" w:author="test">
      <w:r><w:t>one</w:t></w:r>
    </w:del>
    <w:bookmarkEnd w:id="0"/>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:rPr><w:b/></w:rPr><w:instrText xml:space="preserve"> REF BM1 \\* Charformat </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = ExportRunMarkupEvaluatedFromBodyXml(bodyXml);
        var expected = TestCompare.Normalize("Expect bold: one<b>one</b>\n\n");
        Assert.Equal(expected, TestCompare.Normalize(actual));
    }

    [Fact]
    public void Walker_EvalMode_DocVariableWithCharformat_UsesFieldCodeRunStyle()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect bold: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:rPr><w:b/></w:rPr><w:instrText xml:space="preserve"> DOCVARIABLE Var1 \\* Charformat </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var delegates = new DxpFieldEvalDelegates {
            ResolveDocVariableAsync = (name, _) => Task.FromResult<DxpFieldValue?>(name == "Var1" ? new DxpFieldValue("one") : null)
        };
        var eval = new DxpFieldEval(delegates, logger: Logger);
        eval.Context.SetDocVariableNodes("Var1", DxpFieldNodeBuffer.FromText("one"));

        var actual = ExportRunMarkupEvaluatedFromBodyXml(bodyXml, eval);
        var expected = TestCompare.Normalize("Expect bold: <b>one</b>\n\n");
        Assert.Equal(expected, TestCompare.Normalize(actual));
    }

    [Fact]
    public void Walker_EvalMode_RefWithMergeformat_UsesCachedResultRunStyles()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect merge: </w:t></w:r>
    <w:bookmarkStart w:id="0" w:name="BM1"/>
    <w:r><w:t>one</w:t></w:r>
    <w:bookmarkEnd w:id="0"/>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF BM1 \\* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:rPr><w:b/></w:rPr><w:t>cached</w:t></w:r>
    <w:r><w:t>result</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = ExportRunMarkupEvaluatedFromBodyXml(bodyXml);
        var expected = TestCompare.Normalize("Expect merge: one<b>on</b>e\n\n");
        Assert.Equal(expected, TestCompare.Normalize(actual));
    }

    [Fact]
    public void Walker_EvalMode_DocVariableWithMergeformat_UsesCachedResultRunStyles()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect merge: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> DOCVARIABLE Var1 \\* MERGEFORMAT </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:rPr><w:b/></w:rPr><w:t>cached</w:t></w:r>
    <w:r><w:t>result</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var delegates = new DxpFieldEvalDelegates {
            ResolveDocVariableAsync = (name, _) => Task.FromResult<DxpFieldValue?>(name == "Var1" ? new DxpFieldValue("one") : null)
        };
        var eval = new DxpFieldEval(delegates, logger: Logger);
        eval.Context.SetDocVariableNodes("Var1", DxpFieldNodeBuffer.FromText("one"));

        var actual = ExportRunMarkupEvaluatedFromBodyXml(bodyXml, eval);
        var expected = TestCompare.Normalize("Expect merge: <b>on</b>e\n\n");
        Assert.Equal(expected, TestCompare.Normalize(actual));
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
        eval.Context.SetBookmarkNodes("FieldName", DxpFieldNodeBuffer.FromText("FirstName"));

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
    public async Task EvalAsync_GreetingLineUsesMergeMacroProvider()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("en-US");
        eval.Context.MergeMacroProvider = new EnglishMergeMacroProvider();
        eval.Context.MergeCursor = new ListMergeRecordCursor(new[] {
            new Dictionary<string, DxpFieldValue>(StringComparer.OrdinalIgnoreCase) {
                ["Title"] = new DxpFieldValue("Dr."),
                ["LastName"] = new DxpFieldValue("Smith"),
                ["FirstName"] = new DxpFieldValue("Taylor")
            }
        });

        var result = await eval.EvalAsync(new DxpFieldInstruction("GREETINGLINE"));

        Assert.Equal("Dear Dr. Smith,", result.Text);
    }

    [Fact]
    public async Task EvalAsync_GreetingLineFallsBackWhenNamesMissing()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("en-US");
        eval.Context.MergeMacroProvider = new EnglishMergeMacroProvider();
        eval.Context.MergeCursor = new ListMergeRecordCursor(new[] {
            new Dictionary<string, DxpFieldValue>(StringComparer.OrdinalIgnoreCase) {
                ["FirstName"] = new DxpFieldValue("Ana")
            }
        });

        var result = await eval.EvalAsync(new DxpFieldInstruction("GREETINGLINE"));

        Assert.Equal("Dear Ana,", result.Text);

        var fallbackEval = new DxpFieldEval(logger: Logger);
        fallbackEval.Context.Culture = new CultureInfo("en-US");
        fallbackEval.Context.MergeMacroProvider = new EnglishMergeMacroProvider();
        fallbackEval.Context.MergeCursor = new ListMergeRecordCursor(new[] {
            new Dictionary<string, DxpFieldValue>(StringComparer.OrdinalIgnoreCase)
        });

        var fallback = await fallbackEval.EvalAsync(new DxpFieldInstruction("GREETINGLINE"));

        Assert.Equal("Hello,", fallback.Text);
    }

    [Fact]
    public async Task EvalAsync_AddressBlockFormatsAndOmitsEmptyParts()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("en-US");
        eval.Context.MergeMacroProvider = new EnglishMergeMacroProvider();
        eval.Context.MergeCursor = new ListMergeRecordCursor(new[] {
            new Dictionary<string, DxpFieldValue>(StringComparer.OrdinalIgnoreCase) {
                ["Company"] = new DxpFieldValue("Acme Co."),
                ["Address1"] = new DxpFieldValue("123 Main St"),
                ["Address2"] = new DxpFieldValue(string.Empty),
                ["City"] = new DxpFieldValue("Springfield"),
                ["State"] = new DxpFieldValue("IL"),
                ["PostalCode"] = new DxpFieldValue("62704"),
                ["Country"] = new DxpFieldValue("USA")
            }
        });

        var result = await eval.EvalAsync(new DxpFieldInstruction("ADDRESSBLOCK"));

        Assert.Equal("Acme Co.\n123 Main St\nSPRINGFIELD IL  62704\nUSA", result.Text);
    }

    [Fact]
    public async Task EvalAsync_AddressBlock_EnglishUsesUppercaseLocality()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("en-US");
        eval.Context.MergeCursor = new ListMergeRecordCursor(new[] {
            new Dictionary<string, DxpFieldValue>(StringComparer.OrdinalIgnoreCase) {
                ["Company"] = new DxpFieldValue("Acme Co."),
                ["Address1"] = new DxpFieldValue("123 Main St"),
                ["City"] = new DxpFieldValue("Springfield"),
                ["State"] = new DxpFieldValue("IL"),
                ["PostalCode"] = new DxpFieldValue("62704"),
                ["Country"] = new DxpFieldValue("United States")
            }
        });

        var result = await eval.EvalAsync(new DxpFieldInstruction("ADDRESSBLOCK"));

        Assert.Equal("Acme Co.\n123 Main St\nSPRINGFIELD IL  62704\nUNITED STATES", result.Text);
    }

    [Fact]
    public async Task EvalAsync_AddressBlock_FrenchUsesCedex()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("fr-FR");
        eval.Context.MergeCursor = new ListMergeRecordCursor(new[] {
            new Dictionary<string, DxpFieldValue>(StringComparer.OrdinalIgnoreCase) {
                ["Company"] = new DxpFieldValue("Entreprise SA"),
                ["Address1"] = new DxpFieldValue("12 Rue de la République"),
                ["PostalCode"] = new DxpFieldValue("75011"),
                ["City"] = new DxpFieldValue("PARIS"),
                ["Cedex"] = new DxpFieldValue("2"),
                ["Country"] = new DxpFieldValue("France")
            }
        });

        var result = await eval.EvalAsync(new DxpFieldInstruction("ADDRESSBLOCK"));

        Assert.Equal("Entreprise SA\n12 Rue de la République\n75011 PARIS CEDEX 2\nFrance", result.Text);
    }

    [Fact]
    public async Task EvalAsync_AddressBlock_ItalianIncludesProvinceCode()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("it-IT");
        eval.Context.MergeCursor = new ListMergeRecordCursor(new[] {
            new Dictionary<string, DxpFieldValue>(StringComparer.OrdinalIgnoreCase) {
                ["Company"] = new DxpFieldValue("Azienda SRL"),
                ["Address1"] = new DxpFieldValue("Via Nazionale 123"),
                ["PostalCode"] = new DxpFieldValue("00185"),
                ["City"] = new DxpFieldValue("ROMA"),
                ["State"] = new DxpFieldValue("RM"),
                ["Country"] = new DxpFieldValue("Italy")
            }
        });

        var result = await eval.EvalAsync(new DxpFieldInstruction("ADDRESSBLOCK"));

        Assert.Equal("Azienda SRL\nVia Nazionale 123\n00185 ROMA (RM)\nItaly", result.Text);
    }

    [Fact]
    public async Task EvalAsync_AddressBlock_JapaneseUsesPostalMarkAndOrder()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.Culture = new CultureInfo("ja-JP");
        eval.Context.MergeCursor = new ListMergeRecordCursor(new[] {
            new Dictionary<string, DxpFieldValue>(StringComparer.OrdinalIgnoreCase) {
                ["Company"] = new DxpFieldValue("株式会社テスト"),
                ["PostalCode"] = new DxpFieldValue("100-8994"),
                ["State"] = new DxpFieldValue("東京都"),
                ["City"] = new DxpFieldValue("中央区"),
                ["Address1"] = new DxpFieldValue("八重洲1-5-3"),
                ["Address2"] = new DxpFieldValue("ビル2F"),
                ["Country"] = new DxpFieldValue("日本")
            }
        });

        var result = await eval.EvalAsync(new DxpFieldInstruction("ADDRESSBLOCK"));

        Assert.Equal("株式会社テスト\n〒100-8994\n東京都中央区\n八重洲1-5-3\nビル2F\n日本", result.Text);
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
        eval.Context.SetBookmarkNodes("TargetName", DxpFieldNodeBuffer.FromText("Note1"));
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
        eval.Context.SetBookmarkNodes("Answer", DxpFieldNodeBuffer.FromText("Existing"));

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
        eval.Context.SetBookmarkNodes("Greeting", DxpFieldNodeBuffer.FromText("Hi"));
        eval.Context.SetBookmarkNodes("DefaultCity", DxpFieldNodeBuffer.FromText("Rome"));

        var asked = await eval.EvalAsync(new DxpFieldInstruction("ASK City \"{ REF Greeting } there?\" \\d \"{ REF DefaultCity }\""));
        var city = await eval.EvalAsync(new DxpFieldInstruction("REF City"));

        Assert.Equal(string.Empty, asked.Text);
        Assert.Equal("Hi there?", capturedPrompt);
        Assert.Equal("Rome", city.Text);
    }

    [Fact]
    public async Task EvalAsync_DatabaseUsesProviderAndRendersRows()
    {
        var provider = new MockDatabaseProvider {
            Result = new DxpDatabaseResult(
                new[] {
                    new DxpDatabaseColumn("Id"),
                    new DxpDatabaseColumn("Name")
                },
                new[] {
                    new DxpFieldValue?[] { new DxpFieldValue(1), new DxpFieldValue("Alpha") },
                    new DxpFieldValue?[] { new DxpFieldValue(2), new DxpFieldValue("Beta") }
                })
        };
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.DatabaseProvider = provider;

        var result = await eval.EvalAsync(new DxpFieldInstruction("DATABASE \"SELECT\" Id \"7\""));

        Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
        Assert.Equal("Id\tName\n1\tAlpha\n2\tBeta", result.Text);
        Assert.NotNull(provider.LastRequest);
        Assert.Equal("SELECT", provider.LastRequest!.QueryText);
        Assert.NotNull(provider.LastRequest.Parameters);
        Assert.True(provider.LastRequest.Parameters!.TryGetValue("Id", out var param));
        Assert.Equal("7", param.StringValue);
    }

    [Fact]
    public async Task EvalAsync_DatabaseWithoutProviderReturnsEmpty()
    {
        var eval = new DxpFieldEval(logger: Logger);

        var result = await eval.EvalAsync(new DxpFieldInstruction("DATABASE \"SELECT\""));

        Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
        Assert.Equal(string.Empty, result.Text);
    }

    private sealed class CustomResolver : IDxpFieldValueResolver
    {
        public Task<DxpFieldValue?> ResolveAsync(string name, DxpFieldValueKindHint kind, DxpFieldEvalContext context)
        {
            if (name == "Y")
                return Task.FromResult<DxpFieldValue?>(new DxpFieldValue(9));
            return Task.FromResult<DxpFieldValue?>(null);
        }
    }

    private sealed class MockRefResolver : IDxpRefResolver
    {
        public Task<DxpRefRecord?> ResolveAsync(
            DxpRefRequest request,
            DxpFieldEvalContext context,
            DxpIDocumentContext? documentContext)
        {
            _ = documentContext;
            if (request.Bookmark == "Note1" && request.Footnote)
            {
                return Task.FromResult<DxpRefRecord?>(
                    new DxpRefRecord(
                        Bookmark: request.Bookmark,
                        Nodes: null,
                        DocumentText: null,
                        DocumentOrder: null,
                        ParagraphNumber: null,
                        Footnote: new DxpRefFootnote(request.Bookmark, "1", "[1]"),
                        Endnote: null,
                        Hyperlink: null));
            }
            if (request.Bookmark == "Section")
            {
                return Task.FromResult<DxpRefRecord?>(
                    new DxpRefRecord(
                        Bookmark: request.Bookmark,
                        Nodes: null,
                        DocumentText: "Section 1.01",
                        DocumentOrder: 2,
                        ParagraphNumber: new DxpRefParagraphNumber(request.Bookmark, "Section 1.01", "1.01", "101"),
                        Footnote: null,
                        Endnote: null,
                        Hyperlink: null));
            }
            if (request.Bookmark == "Link" && request.Hyperlink)
            {
                return Task.FromResult<DxpRefRecord?>(
                    new DxpRefRecord(
                        Bookmark: request.Bookmark,
                        Nodes: null,
                        DocumentText: "LinkText",
                        DocumentOrder: null,
                        ParagraphNumber: null,
                        Footnote: null,
                        Endnote: null,
                        Hyperlink: new DxpRefHyperlink(request.Bookmark, "#target", null)));
            }
            return Task.FromResult<DxpRefRecord?>(null);
        }
    }

    private sealed class MockTableResolver : IDxpTableValueResolver
    {
        private readonly Dictionary<string, IReadOnlyList<double>> _ranges = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<DxpTableRangeDirection, IReadOnlyList<double>> _directions = [];

        public MockTableResolver Add(string range, params double[] values)
        {
            _ranges[range] = values;
            return this;
        }

        public MockTableResolver AddDirection(DxpTableRangeDirection direction, params double[] values)
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

        public Task<IReadOnlyList<double>> ResolveDirectionalRangeAsync(DxpTableRangeDirection direction, DxpFieldEvalContext context)
        {
            if (_directions.TryGetValue(direction, out var values))
                return Task.FromResult(values);
            return Task.FromResult<IReadOnlyList<double>>([]);
        }
    }

    private sealed class MockDatabaseProvider : IDatabaseFieldProvider
    {
        public DxpDatabaseRequest? LastRequest { get; private set; }
        public DxpDatabaseResult? Result { get; set; }

        public Task<DxpDatabaseResult?> ExecuteAsync(
            DxpDatabaseRequest request,
            CancellationToken cancellationToken)
        {
            _ = cancellationToken;
            LastRequest = request;
            return Task.FromResult(Result);
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
        eval.Context.SetBookmarkNodes("A", DxpFieldNodeBuffer.FromText("5"));

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
        eval.Context.SetBookmarkNodes("Value", DxpFieldNodeBuffer.FromText("5"));

        var result = await eval.EvalAsync(new DxpFieldInstruction("COMPARE Value >= 5"));
        var result2 = await eval.EvalAsync(new DxpFieldInstruction("COMPARE Value < 5"));

        Assert.Equal("1", result.Text);
        Assert.Equal("0", result2.Text);
    }

    [Fact]
    public async Task EvalAsync_SkipIfAndNextIfReturnSkippedStatus()
    {
        var eval = new DxpFieldEval(logger: Logger);
        eval.Context.SetBookmarkNodes("Value", DxpFieldNodeBuffer.FromText("5"));

        var skip = await eval.EvalAsync(new DxpFieldInstruction("SKIPIF Value >= 5"));
        var next = await eval.EvalAsync(new DxpFieldInstruction("NEXTIF Value >= 5"));
        var skipFalse = await eval.EvalAsync(new DxpFieldInstruction("SKIPIF Value < 5"));

        Assert.Equal(DxpFieldEvalStatus.Skipped, skip.Status);
        Assert.Equal(DxpFieldEvalStatus.Skipped, next.Status);
        Assert.Equal(DxpFieldEvalStatus.Resolved, skipFalse.Status);
        Assert.Equal(string.Empty, skipFalse.Text);
    }

    [Fact]
    public async Task EvalAsync_NextAdvancesCursor()
    {
        var eval = new DxpFieldEval(logger: Logger);
        var cursor = new TrackingMergeRecordCursor(2);
        eval.Context.MergeCursor = cursor;

        var result = await eval.EvalAsync(new DxpFieldInstruction("NEXT"));

        Assert.Equal(DxpFieldEvalStatus.Resolved, result.Status);
        Assert.Equal(1, cursor.Index);
        Assert.Equal(1, cursor.MoveNextCalls);
        Assert.Equal(DxpMergeRecordAction.Advance, eval.Context.MergeRecordAction);
    }

    [Fact]
    public async Task EvalAsync_NextIfAdvancesOnlyWhenTrue()
    {
        var evalTrue = new DxpFieldEval(logger: Logger);
        var cursorTrue = new TrackingMergeRecordCursor(2);
        evalTrue.Context.MergeCursor = cursorTrue;

        var resultTrue = await evalTrue.EvalAsync(new DxpFieldInstruction("NEXTIF 1 = 1"));

        Assert.Equal(DxpFieldEvalStatus.Skipped, resultTrue.Status);
        Assert.Equal(1, cursorTrue.Index);
        Assert.Equal(1, cursorTrue.MoveNextCalls);
        Assert.Equal(DxpMergeRecordAction.Advance, evalTrue.Context.MergeRecordAction);

        var evalFalse = new DxpFieldEval(logger: Logger);
        var cursorFalse = new TrackingMergeRecordCursor(2);
        evalFalse.Context.MergeCursor = cursorFalse;

        var resultFalse = await evalFalse.EvalAsync(new DxpFieldInstruction("NEXTIF 1 = 0"));

        Assert.Equal(DxpFieldEvalStatus.Resolved, resultFalse.Status);
        Assert.Equal(0, cursorFalse.Index);
        Assert.Equal(0, cursorFalse.MoveNextCalls);
        Assert.Equal(DxpMergeRecordAction.None, evalFalse.Context.MergeRecordAction);
    }

    [Fact]
    public async Task EvalAsync_SkipIfMarksSkipOutputAndAdvances()
    {
        var eval = new DxpFieldEval(logger: Logger);
        var cursor = new TrackingMergeRecordCursor(2);
        eval.Context.MergeCursor = cursor;

        var result = await eval.EvalAsync(new DxpFieldInstruction("SKIPIF 1 = 1"));

        Assert.Equal(DxpFieldEvalStatus.Skipped, result.Status);
        Assert.Equal(1, cursor.Index);
        Assert.Equal(1, cursor.MoveNextCalls);
        Assert.Equal(DxpMergeRecordAction.SkipOutput, eval.Context.MergeRecordAction);
    }

    [Fact]
    public void MergeLoop_AdvancesAndSkipsRecords()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Name=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> MERGEFIELD Name </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Next=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NEXTIF Flag = 1 </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> MERGEFIELD Name </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SKIPIF Skip = 1 </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var docBytes = BuildDocBytesFromBodyXml(bodyXml);
        var cursor = new ListMergeRecordCursor(new List<Dictionary<string, DxpFieldValue>>
        {
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Name"] = new DxpFieldValue("Alpha"),
                ["Flag"] = new DxpFieldValue(1.0),
                ["Skip"] = new DxpFieldValue(0.0)
            },
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Name"] = new DxpFieldValue("Beta"),
                ["Flag"] = new DxpFieldValue(0.0),
                ["Skip"] = new DxpFieldValue(0.0)
            },
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Name"] = new DxpFieldValue("Gamma"),
                ["Flag"] = new DxpFieldValue(0.0),
                ["Skip"] = new DxpFieldValue(1.0)
            },
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Name"] = new DxpFieldValue("Delta"),
                ["Flag"] = new DxpFieldValue(0.0),
                ["Skip"] = new DxpFieldValue(0.0)
            }
        });

        using var readDoc = WordprocessingDocument.Open(new MemoryStream(docBytes), false);
        var results = DxpMergeLoop.MergePlainText(readDoc, cursor, logger: Logger);

        Assert.Equal(3, results.Count);
        Assert.Equal(TestCompare.Normalize("Name=Alpha Next=Beta\n\n"), TestCompare.Normalize(results[0]));
        Assert.Equal(TestCompare.Normalize("Name=Beta Next=Beta\n\n"), TestCompare.Normalize(results[1]));
        Assert.Equal(TestCompare.Normalize("Name=Delta Next=Delta\n\n"), TestCompare.Normalize(results[2]));
    }

    [Fact]
    public void MergeLoop_HtmlMarkdownPlainText_ShareRecordCounts()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Name=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> MERGEFIELD Name </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SKIPIF Skip = 1 </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var docBytes = BuildDocBytesFromBodyXml(bodyXml);
        var records = new List<Dictionary<string, DxpFieldValue>>
        {
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Name"] = new DxpFieldValue("Alice"),
                ["Skip"] = new DxpFieldValue(0.0)
            },
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Name"] = new DxpFieldValue("Bob"),
                ["Skip"] = new DxpFieldValue(1.0)
            },
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Name"] = new DxpFieldValue("Cara"),
                ["Skip"] = new DxpFieldValue(0.0)
            }
        };

        using var plainDoc = WordprocessingDocument.Open(new MemoryStream(docBytes), false);
        using var htmlDoc = WordprocessingDocument.Open(new MemoryStream(docBytes), false);
        using var mdDoc = WordprocessingDocument.Open(new MemoryStream(docBytes), false);

        var plainResults = DxpMergeLoop.MergePlainText(plainDoc, new ListMergeRecordCursor(records), logger: Logger);
        var htmlResults = DxpMergeLoop.MergeHtml(htmlDoc, new ListMergeRecordCursor(records), logger: Logger);
        var markdownResults = DxpMergeLoop.MergeMarkdown(mdDoc, new ListMergeRecordCursor(records), logger: Logger);

        Assert.Equal(2, plainResults.Count);
        Assert.Equal(plainResults.Count, htmlResults.Count);
        Assert.Equal(plainResults.Count, markdownResults.Count);

        Assert.Contains("Alice", plainResults[0], StringComparison.Ordinal);
        Assert.Contains("Cara", plainResults[1], StringComparison.Ordinal);

        Assert.Contains("Alice", htmlResults[0], StringComparison.Ordinal);
        Assert.Contains("Cara", htmlResults[1], StringComparison.Ordinal);

        Assert.Contains("Alice", markdownResults[0], StringComparison.Ordinal);
        Assert.Contains("Cara", markdownResults[1], StringComparison.Ordinal);
    }

    [Fact]
    public void MergeLoop_RespectsAdvanceAndSkipOutput_WithoutDoc()
    {
        var records = new List<Dictionary<string, DxpFieldValue>>
        {
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Name"] = new DxpFieldValue("Alpha"),
                ["Flag"] = new DxpFieldValue(1.0),
                ["Skip"] = new DxpFieldValue(0.0)
            },
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Name"] = new DxpFieldValue("Beta"),
                ["Flag"] = new DxpFieldValue(0.0),
                ["Skip"] = new DxpFieldValue(0.0)
            },
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Name"] = new DxpFieldValue("Gamma"),
                ["Flag"] = new DxpFieldValue(0.0),
                ["Skip"] = new DxpFieldValue(1.0)
            },
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["Name"] = new DxpFieldValue("Delta"),
                ["Flag"] = new DxpFieldValue(0.0),
                ["Skip"] = new DxpFieldValue(0.0)
            }
        };

        var cursor = new ListMergeRecordCursor(records);
        var outputs = DxpExport.ExportToStrings(cursor, eval => {
            var name = eval.EvalAsync(new DxpFieldInstruction("MERGEFIELD Name")).GetAwaiter().GetResult().Text ?? string.Empty;
            _ = eval.EvalAsync(new DxpFieldInstruction("NEXTIF Flag = 1")).GetAwaiter().GetResult();
            var nextName = eval.EvalAsync(new DxpFieldInstruction("MERGEFIELD Name")).GetAwaiter().GetResult().Text ?? string.Empty;
            _ = eval.EvalAsync(new DxpFieldInstruction("SKIPIF Skip = 1")).GetAwaiter().GetResult();
            return $"{name}-{nextName}";
        }, logger: Logger);

        Assert.Equal(3, outputs.Count);
        Assert.Equal("Alpha-Beta", outputs[0]);
        Assert.Equal("Beta-Beta", outputs[1]);
        Assert.Equal("Delta-Delta", outputs[2]);
    }

    [Fact]
    public void MergeLoop_IsolatesBookmarksPerRecord()
    {
        var records = new List<Dictionary<string, DxpFieldValue>>
        {
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["SetToken"] = new DxpFieldValue("1")
            },
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["SetToken"] = new DxpFieldValue("0")
            }
        };

        var cursor = new ListMergeRecordCursor(records);
        var outputs = DxpExport.ExportToStrings(cursor, eval => {
            var setToken = eval.EvalAsync(new DxpFieldInstruction("MERGEFIELD SetToken")).GetAwaiter().GetResult().Text ?? string.Empty;
            if (string.Equals(setToken, "1", StringComparison.Ordinal))
                _ = eval.EvalAsync(new DxpFieldInstruction("SET Token \"Alpha\"")).GetAwaiter().GetResult();
            var value = eval.EvalAsync(new DxpFieldInstruction("REF Token")).GetAwaiter().GetResult().Text ?? string.Empty;
            return value;
        }, logger: Logger);

        Assert.Equal(2, outputs.Count);
        Assert.Equal("Alpha", outputs[0]);
        Assert.NotEqual("Alpha", outputs[1]);
    }

    [Fact]
    public void MergeLoop_MergeRecAndMergeSeq_RespectSkipOutput()
    {
        var records = new List<Dictionary<string, DxpFieldValue>>
        {
            new(StringComparer.OrdinalIgnoreCase) { ["Skip"] = new DxpFieldValue(0.0) },
            new(StringComparer.OrdinalIgnoreCase) { ["Skip"] = new DxpFieldValue(1.0) },
            new(StringComparer.OrdinalIgnoreCase) { ["Skip"] = new DxpFieldValue(0.0) }
        };

        var cursor = new ListMergeRecordCursor(records);
        var outputs = DxpExport.ExportToStrings(cursor, eval => {
            var rec = eval.EvalAsync(new DxpFieldInstruction("MERGEREC")).GetAwaiter().GetResult().Text ?? string.Empty;
            var seq = eval.EvalAsync(new DxpFieldInstruction("MERGESEQ")).GetAwaiter().GetResult().Text ?? string.Empty;
            _ = eval.EvalAsync(new DxpFieldInstruction("SKIPIF Skip = 1")).GetAwaiter().GetResult();
            return $"{rec}-{seq}";
        }, logger: Logger);

        Assert.Equal(2, outputs.Count);
        Assert.Equal("1-1", outputs[0]);
        Assert.Equal("3-2", outputs[1]);
    }

    [Fact]
    public void EvalAsync_DocumentMetrics_ResolveFromExtendedProperties()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Pages=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NUMPAGES \\# "0000" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Words=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NUMWORDS </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Chars=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NUMCHARS </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(
            bodyXml,
            configureDocument: doc => {
                var part = doc.AddNewPart<ExtendedFilePropertiesPart>();
                part.Properties = new ExtendedProperties.Properties(
                    new ExtendedProperties.Pages { Text = "12" },
                    new ExtendedProperties.Words { Text = "345" },
                    new ExtendedProperties.Characters { Text = "6789" });
                part.Properties.Save();
            }));

        var expected = TestCompare.Normalize("Pages=0012 Words=345 Chars=6789\n\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void EvalAsync_DocumentMetrics_MissingPropertiesYieldEmpty()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Pages=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Words=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NUMWORDS </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> Chars=</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NUMCHARS </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(bodyXml));
        var expected = TestCompare.Normalize("Pages= Words= Chars=\n\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public async Task EvalAsync_SeqIncrementsAndResets()
    {
        var eval = new DxpFieldEval(logger: Logger);

        var first = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure"));
        var second = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure"));
        var repeat = await eval.EvalAsync(new DxpFieldInstruction("SEQ Figure \\c"));

        eval.Context.SetBookmarkNodes("Start", DxpFieldNodeBuffer.FromText("10"));
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
        eval.Context.SetBookmarkNodes("SeqName", DxpFieldNodeBuffer.FromText("Figure"));
        eval.Context.SetBookmarkNodes("ResetName", DxpFieldNodeBuffer.FromText("Start"));
        eval.Context.SetBookmarkNodes("Start", DxpFieldNodeBuffer.FromText("10"));

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
            .AddDirection(DxpTableRangeDirection.Left, 5, 7);

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
        ctx.SetBookmarkNodes("TotalCost", DxpFieldNodeBuffer.FromText("123.45"));

        Assert.True(ctx.TryGetBookmarkNodes("totalcost", out var nodes));
        Assert.Equal("123.45", nodes.ToPlainText());
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
    public void Walker_EvalMode_NestedRefWithInnerFieldChars_Substitutes()
    {
        var bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SET X "Y" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Y</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SET Y "Z" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Z</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF X </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Y</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:instrText xml:space="preserve"> </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Z</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(bodyXml));
        var expected = TestCompare.Normalize("Expect: Z\n\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Walker_EvalMode_NestedRefAcrossInstructionRuns_Substitutes()
    {
        var bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SET X "Y" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Y</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SET Y "Z" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Z</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF { </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> REF X </w:instrText></w:r>
    <w:r><w:instrText xml:space="preserve"> } </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>?</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(bodyXml));
        var expected = TestCompare.Normalize("Expect: Z\n\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Walker_EvalMode_NestedRefSplitBetweenInstructionAndResult_DoesNotSubstitute()
    {
        var bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SET X "Y" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Y</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SET Y "Z" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Z</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF { </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t xml:space="preserve"> REF X } </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(bodyXml));
        Assert.DoesNotContain("Z", actual);
    }

    [Fact]
    public void Walker_EvalMode_NestedRefInlineInstruction_Substitutes()
    {
        var bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p>
    <w:r><w:t xml:space="preserve">Expect: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SET X "Y" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Y</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> SET Y "Z" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Z</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> REF { REF X } </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>?</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(bodyXml));
        var expected = TestCompare.Normalize("Expect: Z\n\n");
        Assert.Equal(expected, actual);
    }


    [Fact]
    public void Walker_EvalMode_IfWithNestedRefInTrueBranch_EmitsRefResult()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:bookmarkStart w:id="0" w:name="BM1"/>
    <w:r><w:t>one</w:t></w:r>
    <w:bookmarkEnd w:id="0"/>
  </w:p>
  <w:p>
    <w:r><w:t xml:space="preserve">Expect: </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> IF 1 = 1 "</w:instrText></w:r>
    <w:fldSimple w:instr=" REF BM1 ">
      <w:r><w:instrText>one</w:instrText></w:r>
    </w:fldSimple>
    <w:r><w:instrText xml:space="preserve">" "no" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>cached</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = TestCompare.Normalize(ExportPlainTextEvaluatedFromBodyXml(bodyXml));
        var expected = TestCompare.Normalize("one\n\nExpect: one\n\n");
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void Walker_EvalMode_InlineIfPreservesInlineFormatting()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p w14:paraId="7B90ECD8" w14:textId="271E59D9" w:rsidR="00065997" w:rsidRPr="00065997" w:rsidRDefault="00065997">
    <w:r><w:t xml:space="preserve">Expect </w:t></w:r>
    <w:r w:rsidRPr="00065997"><w:rPr><w:b/><w:bCs/></w:rPr><w:t>1</w:t></w:r>
    <w:r w:rsidRPr="00065997"><w:rPr><w:u w:val="single"/></w:rPr><w:t>2</w:t></w:r>
    <w:r w:rsidRPr="00065997"><w:rPr><w:b/><w:bCs/></w:rPr><w:t>3</w:t></w:r>
    <w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">: </w:t></w:r>
    <w:r w:rsidRPr="00065997"><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r w:rsidRPr="00065997"><w:instrText xml:space="preserve"> IF 1 = 1 "</w:instrText></w:r>
    <w:r w:rsidRPr="00065997"><w:rPr><w:b/><w:bCs/></w:rPr><w:instrText>1</w:instrText></w:r>
    <w:r w:rsidRPr="00065997"><w:rPr><w:u w:val="single"/></w:rPr><w:instrText>2</w:instrText></w:r>
    <w:r w:rsidRPr="00065997"><w:rPr><w:b/><w:bCs/></w:rPr><w:instrText>3</w:instrText></w:r>
    <w:r w:rsidRPr="00065997"><w:instrText xml:space="preserve">" "error" </w:instrText></w:r>
    <w:r w:rsidRPr="00065997"><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r w:rsidR="001D199F" w:rsidRPr="00065997"><w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr><w:t>1</w:t></w:r>
    <w:r w:rsidR="001D199F" w:rsidRPr="00065997"><w:rPr><w:noProof/><w:u w:val="single"/></w:rPr><w:t>2</w:t></w:r>
    <w:r w:rsidR="001D199F" w:rsidRPr="00065997"><w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr><w:t>3</w:t></w:r>
    <w:r w:rsidRPr="00065997"><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var actual = TestCompare.Normalize(ExportRunMarkupEvaluatedFromBodyXml(bodyXml));
        var expected = TestCompare.Normalize("Expect <b>1</b><u>2</u><b>3: 1</b><u>2</u><b>3</b>\n\n");
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData(DxpEvalFieldMode.Evaluate)]
    [InlineData(DxpEvalFieldMode.Cache)]
    public void Walker_FieldEval_InlineIfWithDocVariable_PreservesParagraphRunBackground(DxpEvalFieldMode mode)
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p w14:paraId="64F18825" w14:textId="6D3566A0" w:rsidR="000A7DB0" w:rsidRDefault="000A7DB0">
    <w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
    </w:pPr>
  </w:p>
  <w:p w14:paraId="2DEC30ED" w14:textId="03AB7D46" w:rsidR="00822BD9" w:rsidRPr="002C7D37" w:rsidRDefault="00822BD9" w:rsidP="00822BD9">
    <w:pPr>
      <w:jc w:val="right"/>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
    </w:pPr>
    <w:r w:rsidRPr="002C7D37">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r w:rsidRPr="002C7D37">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> IF </w:instrText>
    </w:r>
    <w:r w:rsidRPr="002C7D37">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r w:rsidRPr="002C7D37">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> DOCVARIABLE </w:instrText>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
      <w:instrText>GREENTECH</w:instrText>
    </w:r>
    <w:r w:rsidRPr="002C7D37">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r w:rsidRPr="002C7D37">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
      <w:instrText>OK</w:instrText>
    </w:r>
    <w:r w:rsidRPr="002C7D37">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
    <w:r w:rsidRPr="002C7D37">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> = "OK" "OK" "NO" </w:instrText>
    </w:r>
    <w:r w:rsidRPr="002C7D37">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r w:rsidRPr="002C7D37">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
      <w:t>OK</w:t>
    </w:r>
    <w:r w:rsidRPr="002C7D37">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:color w:val="FFFFFF" w:themeColor="background1"/>
        <w:sz w:val="23"/>
        <w:shd w:val="solid" w:color="auto" w:fill="000000"/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>
</w:body>
""";

        var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
            ResolveDocVariableAsync = (name, ctx) => Task.FromResult<DxpFieldValue?>(name == "GREENTECH" ? new DxpFieldValue("OK") : null)
        }, logger: Logger);

        var html = ExportHtmlFromBodyXml(bodyXml, mode, eval);

        Assert.Contains("OK", html, StringComparison.Ordinal);
        Assert.Contains("align-right", html, StringComparison.Ordinal);
        Assert.Contains("color:#ffffff", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("background-color:#000000", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("font-family:Arial", html, StringComparison.OrdinalIgnoreCase);
    }

	[Theory]
	[InlineData(DxpEvalFieldMode.Evaluate)]
	[InlineData(DxpEvalFieldMode.Cache)]
	public void Walker_FieldEval_InlineIfWithDocVariableDateFormat_FormatsDate(DxpEvalFieldMode mode)
	{
		const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p w14:paraId="68483D31" w14:textId="5C6B5B73" w:rsidR="00054542" w:rsidRPr="00A719A3" w:rsidRDefault="00DA27CD" w:rsidP="00C03949">
    <w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
    </w:pPr>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:bCs/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> IF </w:instrText>
    </w:r>
    <w:r w:rsidR="00245C11" w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r w:rsidR="00245C11" w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> DOCVARIABLE GrantNo </w:instrText>
    </w:r>
    <w:r w:rsidR="00245C11" w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r w:rsidR="000F128F" w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> </w:instrText>
    </w:r>
    <w:r w:rsidR="00245C11" w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> = " " "</w:instrText>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:bCs/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> DOCVARIABLE </w:instrText>
    </w:r>
    <w:r w:rsidR="00A67C8D" w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText>ApplicationD</w:instrText>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText>ate</w:instrText>
    </w:r>
    <w:r w:rsidR="00A67C8D" w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> \\@ "MMMM d, yyyy"</w:instrText>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> </w:instrText>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:bCs/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r w:rsidR="000F128F" w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
      </w:rPr>
      <w:instrText>Erreur ! Aucune variable de document fournie.</w:instrText>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:bCs/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText>" "</w:instrText>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:bCs/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> DOCVARIABLE GrantDate</w:instrText>
    </w:r>
    <w:r w:rsidR="00A67C8D" w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> \\@ "MMMM d, yyyy"</w:instrText>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> </w:instrText>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:bCs/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText>January 4, 2014</w:instrText>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:bCs/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      </w:rPr>
      <w:instrText xml:space="preserve">" </w:instrText>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:bCs/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r w:rsidR="00A81ACC" w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:b/>
        <w:noProof/>
      </w:rPr>
      <w:t>Erreur ! Aucune variable de document fournie.</w:t>
    </w:r>
    <w:r w:rsidRPr="00EB4FDB">
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:bCs/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>
</w:body>
""";

		var eval = new DxpFieldEval(new DxpFieldEvalDelegates {
			ResolveDocVariableAsync = (name, ctx) => name switch {
				"GrantNo" => Task.FromResult<DxpFieldValue?>(new DxpFieldValue("")),
				"ApplicationDate" => Task.FromResult<DxpFieldValue?>(new DxpFieldValue(new DateTimeOffset(2014, 1, 4, 0, 0, 0, TimeSpan.Zero))),
				"GrantDate" => Task.FromResult<DxpFieldValue?>(new DxpFieldValue(new DateTimeOffset(2014, 1, 4, 0, 0, 0, TimeSpan.Zero))),
				_ => Task.FromResult<DxpFieldValue?>(null)
			}
		}, logger: Logger);

		var html = ExportHtmlFromBodyXml(bodyXml, mode, eval);

		if (mode == DxpEvalFieldMode.Evaluate)
			Assert.Contains("January 4, 2014", html, StringComparison.Ordinal);
		else
			Assert.Contains("Erreur ! Aucune variable de document fournie.", html, StringComparison.Ordinal);
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
        var resolver = new CapturingTableResolver()
            .AddDirection(DxpTableRangeDirection.Below, 12)
            .AddDirection(DxpTableRangeDirection.Above, 1)
            .AddDirection(DxpTableRangeDirection.Left, 6)
            .AddDirection(DxpTableRangeDirection.Right, 16);
        eval.Context.TableResolver = resolver;
        var visitor = DxpVisitorMiddleware.Chain(
            new DxpVisitor(Logger),
            next => new DxpFieldEvalMiddleware(next, eval, logger: Logger),
            next => new DxpContextMiddleware(next, Logger));

        using (var readDoc = WordprocessingDocument.Open(stream, false))
            new DxpWalker(Logger).Accept(readDoc, visitor);

        Assert.Contains(DxpTableRangeDirection.Below, resolver.DirectionCalls);
        Assert.Contains(DxpTableRangeDirection.Above, resolver.DirectionCalls);
        Assert.Contains(DxpTableRangeDirection.Left, resolver.DirectionCalls);
        Assert.Contains(DxpTableRangeDirection.Right, resolver.DirectionCalls);
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
        var resolver = new CapturingRefResolver(new DxpRefIndexResolver(), Logger);
        var options = new DxpEvalFieldMiddlewareOptions
        {
            RefResolver = resolver
        };
        var visitor = DxpVisitorMiddleware.Chain(
            new DxpVisitor(Logger),
            next => new DxpFieldEvalMiddleware(next, eval, logger: Logger, options: options),
            next => new DxpContextMiddleware(next, Logger));

        using (var readDoc = WordprocessingDocument.Open(stream, false))
        {
            new DxpWalker(Logger).Accept(readDoc, visitor);
        }

        DxpRefResult? FindResult(
            bool paragraphNumber = false,
            bool aboveBelow = false,
            bool footnote = false,
            bool hyperlink = false)
        {
            return resolver.Calls
                .Where(call => call.request.Bookmark == "BM1")
                .Where(call => call.request.ParagraphNumber == paragraphNumber)
                .Where(call => call.request.AboveBelow == aboveBelow)
                .Where(call => call.request.Footnote == footnote)
                .Where(call => call.request.Hyperlink == hyperlink)
                .Select(call => call.result)
                .FirstOrDefault();
        }

        Assert.Equal(4, resolver.Calls.Count);
        Assert.Equal("1", FindResult(paragraphNumber: true)?.Text);
        Assert.Equal("1 above", FindResult(paragraphNumber: true, aboveBelow: true)?.Text);
        Assert.Equal("Footnote text", FindResult(footnote: true)?.Text);
        Assert.Equal("Bookmark Textlink", FindResult(hyperlink: true)?.Text);
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

    private static byte[] BuildDocBytesFromBodyXml(string bodyXml)
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

        return stream.ToArray();
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

        if (visitor is not DxpIFieldEvalProvider provider)
            throw new XunitException("DxpPlainTextVisitor should provide field evaluation context.");

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => new DxpFieldEvalMiddleware(next, provider.FieldEval, DxpEvalFieldMode.Cache, logger: Logger),
            next => new DxpContextMiddleware(next, Logger));

        using (var readDoc = WordprocessingDocument.Open(stream, false))
            new DxpWalker(Logger).Accept(readDoc, pipeline);

        return writer.ToString();
    }

    private string ExportPlainTextEvaluatedFromBodyXml(
        string bodyXml,
        DxpFieldEval? fieldEval = null,
        bool includeCustomProperties = false,
        Action<WordprocessingDocument>? configureDocument = null)
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
            configureDocument?.Invoke(doc);
        }

        stream.Position = 0;

        var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig(), Logger, fieldEval);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        if (visitor is not DxpIFieldEvalProvider provider)
            throw new XunitException("DxpPlainTextVisitor should provide field evaluation context.");

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => new DxpFieldEvalMiddleware(
                next,
                provider.FieldEval,
                DxpEvalFieldMode.Evaluate,
                includeDocumentProperties: true,
                includeCustomProperties: includeCustomProperties,
                logger: Logger),
            next => new DxpContextMiddleware(next, Logger));

        using (var readDoc = WordprocessingDocument.Open(stream, false))
            new DxpWalker(Logger).Accept(readDoc, pipeline);

        return writer.ToString();
    }

    private string ExportRunMarkupEvaluatedFromBodyXml(string bodyXml, DxpFieldEval? fieldEval = null)
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

        var visitor = new RunMarkupVisitor();
        var eval = fieldEval ?? new DxpFieldEval(logger: Logger);

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => new DxpFieldEvalMiddleware(next, eval, DxpEvalFieldMode.Evaluate, logger: Logger),
            next => new DxpContextMiddleware(next, Logger));

        using (var readDoc = WordprocessingDocument.Open(stream, false))
            new DxpWalker(Logger).Accept(readDoc, pipeline);

        return visitor.ToString();
    }

    private string ExportHtmlFromBodyXml(string bodyXml, DxpEvalFieldMode mode, DxpFieldEval? fieldEval = null)
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

        var config = DxpHtmlVisitorConfig.CreateRichConfig();
        var visitor = new DxpHtmlVisitor(config, Logger, fieldEval);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        if (visitor is not DxpIFieldEvalProvider provider)
            throw new XunitException("DxpHtmlVisitor should provide field evaluation context.");

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => new DxpFieldEvalMiddleware(next, provider.FieldEval, mode, logger: Logger),
            next => new DxpContextMiddleware(next, Logger));

        using (var readDoc = WordprocessingDocument.Open(stream, false))
            new DxpWalker(Logger).Accept(readDoc, pipeline);

        return TestCompare.Normalize(writer.ToString());
    }


    private sealed class RunMarkupVisitor : DxpVisitor, DxpITextVisitor
    {
        private readonly StringBuilder _builder = new();

        public RunMarkupVisitor() : base(null)
        {
        }

        public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
        {
            return DxpDisposable.Empty;
        }

        public override void VisitText(Text t, DxpIDocumentContext d)
        {
            _builder.Append(t.Text);
        }

        public override void StyleBoldBegin(DxpIDocumentContext d) => _builder.Append("<b>");
        public override void StyleBoldEnd(DxpIDocumentContext d) => _builder.Append("</b>");
        public override void StyleUnderlineBegin(DxpIDocumentContext d) => _builder.Append("<u>");
        public override void StyleUnderlineEnd(DxpIDocumentContext d) => _builder.Append("</u>");

        public void SetOutput(TextWriter writer)
        {
        }

        public override string ToString() => _builder.ToString();
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

        public override void VisitText(Text t, DxpIDocumentContext d)
        {
            var current = d.CurrentFields.Current;
            if (current?.InResult != true)
                return;
            var instruction = current.InstructionText;
            if (string.IsNullOrWhiteSpace(instruction))
                return;

            var key = instruction.Trim();
            if (Results.TryGetValue(key, out var existing))
                Results[key] = (existing ?? string.Empty) + t.Text;
            else
                Results[key] = t.Text;
        }
    }

    private sealed class CapturingRefResolver : IDxpRefResolver
    {
        private readonly IDxpRefResolver _inner;
        private readonly ILogger? _logger;
        public List<(DxpRefRequest request, DxpRefRecord? record, DxpRefResult? result)> Calls { get; } = new();

        public CapturingRefResolver(IDxpRefResolver inner, ILogger? logger = null)
        {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            _logger = logger;
        }

        public async Task<DxpRefRecord?> ResolveAsync(
            DxpRefRequest request,
            DxpFieldEvalContext context,
            DxpIDocumentContext? documentContext)
        {
            _logger?.LogDebug(
                "CapturingRefResolver: resolving {Bookmark} (n={ParagraphNumber}, p={AboveBelow}, f={Footnote}, h={Hyperlink})",
                request.Bookmark,
                request.ParagraphNumber,
                request.AboveBelow,
                request.Footnote,
                request.Hyperlink);
            var record = await _inner.ResolveAsync(request, context, documentContext);
            var result = record?.Format(request, context);
            Calls.Add((request, record, result));
            _logger?.LogDebug(
                "CapturingRefResolver: resolved {Bookmark} -> {Text}",
                request.Bookmark,
                result?.Text ?? "<null>");
            return record;
        }
    }

    private sealed class CapturingTableResolver : IDxpTableValueResolver
    {
        private readonly Dictionary<string, IReadOnlyList<double>> _ranges = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<DxpTableRangeDirection, IReadOnlyList<double>> _directions = [];

        public List<string> RangeCalls { get; } = new();
        public List<DxpTableRangeDirection> DirectionCalls { get; } = new();

        public CapturingTableResolver AddRange(string range, params double[] values)
        {
            _ranges[range] = values;
            return this;
        }

        public CapturingTableResolver AddDirection(DxpTableRangeDirection direction, params double[] values)
        {
            _directions[direction] = values;
            return this;
        }

        public Task<IReadOnlyList<double>> ResolveRangeAsync(string range, DxpFieldEvalContext context)
        {
            RangeCalls.Add(range);
            if (_ranges.TryGetValue(range, out var values))
                return Task.FromResult(values);
            return Task.FromResult<IReadOnlyList<double>>([]);
        }

        public Task<IReadOnlyList<double>> ResolveDirectionalRangeAsync(
            DxpTableRangeDirection direction,
            DxpFieldEvalContext context)
        {
            DirectionCalls.Add(direction);
            if (_directions.TryGetValue(direction, out var values))
                return Task.FromResult(values);
            return Task.FromResult<IReadOnlyList<double>>([]);
        }
    }

    private sealed class ListMergeRecordCursor : IDxpMergeRecordCursor
    {
        private readonly IReadOnlyList<Dictionary<string, DxpFieldValue>> _records;
        private int _index;

        public ListMergeRecordCursor(IReadOnlyList<Dictionary<string, DxpFieldValue>> records)
        {
            _records = records ?? throw new ArgumentNullException(nameof(records));
            _index = 0;
        }

        public bool HasCurrent => _index >= 0 && _index < _records.Count;
        public int RecordIndex => HasCurrent ? _index + 1 : 0;

        public bool MoveNext()
        {
            if (_index + 1 < _records.Count)
            {
                _index++;
                return true;
            }

            _index = _records.Count;
            return false;
        }

        public DxpFieldValue? GetValue(string fieldName)
        {
            if (!HasCurrent)
                return null;

            if (_records[_index].TryGetValue(fieldName, out var value))
                return value;

            return null;
        }
    }

    private sealed class TrackingMergeRecordCursor : IDxpMergeRecordCursor
    {
        private readonly int _count;

        public TrackingMergeRecordCursor(int count)
        {
            _count = count;
            Index = 0;
        }

        public int Index { get; private set; }
        public int MoveNextCalls { get; private set; }
        public bool HasCurrent => Index >= 0 && Index < _count;
        public int RecordIndex => HasCurrent ? Index + 1 : 0;

        public bool MoveNext()
        {
            MoveNextCalls++;
            if (Index + 1 < _count)
            {
                Index++;
                return true;
            }

            Index = _count;
            return false;
        }

        public DxpFieldValue? GetValue(string fieldName)
        {
            _ = fieldName;
            return null;
        }
    }
}
