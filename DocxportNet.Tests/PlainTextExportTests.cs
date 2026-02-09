using DocxportNet.Tests.Utils;
using DocxportNet.Visitors.PlainText;
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
        var visitor = new DxpPlainTextVisitor(config, Logger);
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

        if (visitor is not DocxportNet.Fields.IDxpFieldEvalProvider provider)
            throw new XunitException("DxpPlainTextVisitor should provide field evaluation context.");

        var pipeline = DocxportNet.Walker.DxpVisitorMiddleware.Chain(
            visitor,
            next => new DocxportNet.Walker.DxpFieldEvalMiddleware(next, provider.FieldEval, DocxportNet.Walker.DxpFieldEvalMode.Cache, logger: Logger),
            next => new DocxportNet.Walker.DxpContextTracker(next));

        new DocxportNet.Walker.DxpWalker(Logger).Accept(docxPath, pipeline);
        return writer.ToString();
    }
}
