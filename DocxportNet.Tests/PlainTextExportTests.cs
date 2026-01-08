using DocxportNet.Tests.Utils;
using DocxportNet.Visitors.PlainText;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace DocxportNet.Tests;

public class PlainTextExportTests
{
	private readonly ILogger _logger;

	public PlainTextExportTests(ITestOutputHelper output)
	{
		_logger = new TestLogger(output, nameof(PlainTextExportTests));
	}

	private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
	private static readonly string SamplesDirectory = Path.Combine(ProjectRoot, "samples");

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
			.OrderBy(Path.GetFileName)
			.Select(path => new object[] { new Sample(path) });

	[Theory]
	[MemberData(nameof(SampleDocs))]
	public void AcceptMatchesFixture(Sample sample)
	{
		Verify(sample, DxpPlainTextVisitorConfig.CreateAcceptConfig(), ".test.txt");
	}

	[Theory]
	[MemberData(nameof(SampleDocs))]
	public void RejectMatchesFixture(Sample sample)
	{
		Verify(sample, DxpPlainTextVisitorConfig.CreateRejectConfig(), ".reject.test.txt");
	}

	private void Verify(Sample sample, DxpPlainTextVisitorConfig config, string actualSuffix)
	{
		string basePath = Path.Combine(Path.GetDirectoryName(sample.DocxPath)!, Path.GetFileNameWithoutExtension(sample.DocxPath));
		string expectedPath = config.TrackedChangeMode == DxpPlainTextTrackedChangeMode.RejectChanges
			? basePath + ".reject.txt"
			: basePath + ".txt";
		string actualPath = basePath + actualSuffix;

		string actualText = TestCompare.Normalize(ToPlainText(sample.DocxPath, config));
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

	private string ToPlainText(string docxPath, DxpPlainTextVisitorConfig config)
	{
		using var writer = new StringWriter();
		var visitor = new DxpPlainTextVisitor(writer, config, _logger);
		var walker = new DxpWalker();

		walker.Accept(docxPath, visitor);
		return writer.ToString();
	}
}
