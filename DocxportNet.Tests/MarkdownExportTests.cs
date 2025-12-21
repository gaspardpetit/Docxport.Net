using DocxportNet.Tests.Utils;
using DocxportNet.Visitors;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace DocxportNet.Tests;

public class MarkdownExportTests
{
	private readonly ILogger _logger;

	public MarkdownExportTests(ITestOutputHelper output)
	{
		_logger = new TestLogger(output, nameof(MarkdownExportTests));
	}

	public sealed record Sample : IXunitSerializable
	{
		public Sample()
		{
			DocxPath = string.Empty;
		}

		public Sample(string docxPath)
		{
			DocxPath = docxPath;
		}

		public string DocxPath { get; private set; }
		public string FileName => Path.GetFileName(DocxPath);
		public string ExpectedMarkdownPath => Path.ChangeExtension(DocxPath, ".md");
		public string TestOutputPath => Path.Combine(Path.GetDirectoryName(DocxPath)!, $"{Path.GetFileNameWithoutExtension(DocxPath)}.test.md");

		public void Serialize(IXunitSerializationInfo info) => info.AddValue(nameof(DocxPath), DocxPath);
		public void Deserialize(IXunitSerializationInfo info) => DocxPath = info.GetValue<string>(nameof(DocxPath));

		public override string ToString() => FileName; // keep theory display concise
	}

	private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
	private static readonly string SamplesDirectory = Path.Combine(ProjectRoot, "samples");

	public static IEnumerable<object[]> SampleDocs()
	{
		return Directory.EnumerateFiles(SamplesDirectory, "*.docx", SearchOption.TopDirectoryOnly)
			.OrderBy(Path.GetFileName) // deterministic ordering for discovery
			.Select(path => new object[] { new Sample(path) });
	}

	[Theory]
	[MemberData(nameof(SampleDocs))]
	public void TestDocxToMarkdown_Rich(Sample sample)
	{
		VerifyAgainstFixture(sample, DxpMarkdownVisitorConfig.RICH, ".md", ".test.md");
	}

	[Theory]
	[MemberData(nameof(SampleDocs))]
	public void TestDocxToMarkdown_Plain(Sample sample)
	{
		VerifyAgainstFixture(sample, DxpMarkdownVisitorConfig.PLAIN, ".plain.md", ".plain.test.md");
	}

	private void VerifyAgainstFixture(Sample sample, DxpMarkdownVisitorConfig config, string expectedExt, string actualSuffix)
	{
		string expectedPath = Path.ChangeExtension(sample.DocxPath, expectedExt);
		string actualPath = Path.Combine(Path.GetDirectoryName(sample.DocxPath)!, $"{Path.GetFileNameWithoutExtension(sample.DocxPath)}{actualSuffix}");

		string actualMarkdown = TestCompare.Normalize(ToMarkdown(sample.DocxPath, config));
		File.WriteAllText(actualPath, actualMarkdown);

		if (!File.Exists(expectedPath))
		{
			throw new XunitException($"Expected markdown file missing for {sample.FileName}. Add {expectedPath}. Actual output saved to {actualPath}.");
		}

		string expectedMarkdown = TestCompare.Normalize(File.ReadAllText(expectedPath));

		if (!string.Equals(expectedMarkdown, actualMarkdown, StringComparison.Ordinal))
		{
			string diff = TestCompare.DescribeDifference(expectedMarkdown, actualMarkdown);
			throw new XunitException($"Mismatch for {sample.FileName}: {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
		}
	}

	private string ToMarkdown(string docxPath, DxpMarkdownVisitorConfig config)
	{
		using var writer = new StringWriter();
		var visitor = new DxpMarkdownVisitor(writer, config, _logger);
		var walker = new DxpWalker();

		walker.Accept(docxPath, visitor);
		return writer.ToString();
	}
}
