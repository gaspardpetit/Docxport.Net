using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Resolution;
using DocxportNet.Middleware;
using DocxportNet.Tests.Utils;
using DocxportNet.Visitors.PlainText;
using DocxportNet.Walker;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace DocxportNet.Tests;

public class PlainTextExportTests : TestBase<PlainTextExportTests>
{
    private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
    private static readonly string SamplesDirectory = Path.Combine(ProjectRoot, "samples");

    public PlainTextExportTests(ITestOutputHelper output) : base(output)
    {
    }

    public sealed class Sample : IXunitSerializable
    {
        public Sample()
        {
            DocxPath = string.Empty;
        }

        public Sample(string docxPath)
        {
            DocxPath = docxPath;
        }

        public string DocxPath { get; set; }
        public string FileName => Path.GetFileName(DocxPath);

        public void Serialize(IXunitSerializationInfo info)
        {
            info.AddValue(nameof(DocxPath), DocxPath);
        }

        public void Deserialize(IXunitSerializationInfo info)
        {
            DocxPath = info.GetValue<string>(nameof(DocxPath));
        }

        public override string ToString() => FileName;
    }

    public static IEnumerable<object[]> SampleDocs() =>
        Directory.EnumerateFiles(SamplesDirectory, "*.docx", SearchOption.TopDirectoryOnly)
            .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
            .OrderBy(Path.GetFileName)
            .Select(path => new object[] { new Sample(path) });

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void AcceptMatchesFixture(Sample sample)
    {
        Verify(sample, DxpPlainTextVisitorConfig.CreateAcceptConfig(), ".txt", ".test.txt", DxpFieldEvalExportMode.None);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void RejectMatchesFixture(Sample sample)
    {
        Verify(sample, DxpPlainTextVisitorConfig.CreateRejectConfig(), ".reject.txt", ".reject.test.txt", DxpFieldEvalExportMode.None);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void CachedMatchesFixture(Sample sample)
    {
        VerifyCached(sample, DxpPlainTextVisitorConfig.CreateAcceptConfig(), ".cached.txt", ".cached.test.txt");
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void EvalMatchesFixture(Sample sample)
    {
        Verify(sample, DxpPlainTextVisitorConfig.CreateAcceptConfig(), ".eval.txt", ".eval.test.txt", DxpFieldEvalExportMode.Evaluate);
    }

    private void Verify(
        Sample sample,
        DxpPlainTextVisitorConfig config,
        string expectedExt,
        string actualSuffix,
        DxpFieldEvalExportMode evalMode)
    {
        string expectedPath = config.TrackedChangeMode == DxpPlainTextTrackedChangeMode.RejectChanges
            ? TestPaths.GetSampleOutputPath(sample.DocxPath, ".reject.txt")
            : TestPaths.GetSampleOutputPath(sample.DocxPath, expectedExt);
        string actualPath = TestPaths.GetSampleOutputPath(sample.DocxPath, actualSuffix);

        string actualText = TestCompare.Normalize(ToPlainText(sample.DocxPath, config, evalMode));
        File.WriteAllText(actualPath, actualText);

        if (!File.Exists(expectedPath))
            throw new XunitException($"Expected plain text file missing for {sample.FileName}. Add {expectedPath}. Actual output saved to {actualPath}.");

        string expectedText = TestCompare.Normalize(File.ReadAllText(expectedPath));
        if (!string.Equals(expectedText, actualText, StringComparison.Ordinal))
        {
            string diff = TestCompare.DescribeDifference(expectedText, actualText);
            throw new XunitException($"Mismatch for {sample.FileName}: {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
        }
    }

    private string ToPlainText(string docxPath, DxpPlainTextVisitorConfig config, DxpFieldEvalExportMode evalMode)
    {
        DxpFieldEval? fieldEval = null;
        if (evalMode == DxpFieldEvalExportMode.Evaluate)
            fieldEval = CreateEvalWithAsk();

        var visitor = new DxpPlainTextVisitor(config, Logger, fieldEval);
        var options = new DxpExportOptions { FieldEvalMode = evalMode };
        return DxpExport.ExportToString(docxPath, visitor, options, Logger);
    }

    private void VerifyCached(Sample sample, DxpPlainTextVisitorConfig config, string expectedExt, string actualSuffix)
    {
        string expectedPath = TestPaths.GetSampleOutputPath(sample.DocxPath, expectedExt);
        string actualPath = TestPaths.GetSampleOutputPath(sample.DocxPath, actualSuffix);

        string actualText = TestCompare.Normalize(ToPlainTextCached(sample.DocxPath, config));
        File.WriteAllText(actualPath, actualText);

        if (!File.Exists(expectedPath))
            throw new XunitException($"Expected plain text file missing for {sample.FileName} (CachedFields). Add {expectedPath}. Actual output saved to {actualPath}.");

        string expectedText = TestCompare.Normalize(File.ReadAllText(expectedPath));
        if (!string.Equals(expectedText, actualText, StringComparison.Ordinal))
        {
            string diff = TestCompare.DescribeDifference(expectedText, actualText);
            throw new XunitException($"Mismatch for {sample.FileName} (CachedFields): {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
        }
    }

    private string ToPlainTextCached(string docxPath, DxpPlainTextVisitorConfig config)
    {
        var visitor = new DxpPlainTextVisitor(config, Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        if (visitor is not Fields.DxpIFieldEvalProvider provider)
            throw new XunitException("DxpPlainTextVisitor should provide field evaluation context.");

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => new DxpFieldEvalMiddleware(next, provider.FieldEval, DxpEvalFieldMode.Cache, logger: Logger),
            next => new DxpContextMiddleware(next));

        new DxpWalker(Logger).Accept(docxPath, pipeline);
        return writer.ToString();
    }

    private static void ConfigureEvalContext(DxpFieldEval eval)
    {
        eval.Context.SetDocVariable("Var1", "two");
        eval.Context.SetMergeFieldAlias("GivenName", "FirstName");
        eval.Context.ValueResolver = new DxpChainedFieldValueResolver(
            new SampleFieldValueResolver(),
            new DxpContextFieldValueResolver());
    }

    private DxpFieldEval CreateEvalWithAsk()
    {
        var delegates = new DxpFieldEvalDelegates {
            AskAsync = (prompt, _) => Task.FromResult<DxpFieldValue?>(prompt switch {
                "Name?" => new DxpFieldValue("Bob"),
                "Hi Bob?" => new DxpFieldValue("Montreal"),
                _ => null
            })
        };

        var eval = new DxpFieldEval(delegates, logger: Logger);
        ConfigureEvalContext(eval);
        return eval;
    }

    private sealed class SampleFieldValueResolver : IDxpFieldValueResolver
    {
        public Task<DxpFieldValue?> ResolveAsync(string name, DxpFieldValueKindHint kind, DxpFieldEvalContext context)
        {
            _ = context;
            if (kind == DxpFieldValueKindHint.Any || kind == DxpFieldValueKindHint.MergeField)
            {
                if (string.Equals(name, "FirstName", StringComparison.OrdinalIgnoreCase))
                    return Task.FromResult<DxpFieldValue?>(new DxpFieldValue("Ana"));
                if (string.Equals(name, "EmptyField", StringComparison.OrdinalIgnoreCase))
                    return Task.FromResult<DxpFieldValue?>(new DxpFieldValue(string.Empty));
            }
            return Task.FromResult<DxpFieldValue?>(null);
        }
    }
}
