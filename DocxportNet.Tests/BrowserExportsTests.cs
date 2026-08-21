using DocumentFormat.OpenXml.Packaging;
using DocxportNet.Wasm;

namespace DocxportNet.Tests;

public sealed class BrowserExportsTests
{
    private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
    private static readonly byte[] Sample = File.ReadAllBytes(Path.Combine(ProjectRoot, "samples", "TestLists.docx"));

    [Theory]
    [InlineData(BrowserExportFormat.Html, "<")]
    [InlineData(BrowserExportFormat.Markdown, "-")]
    [InlineData(BrowserExportFormat.Text, "1")]
    public void ExportsEveryTextFormat(BrowserExportFormat format, string expectedFragment)
    {
        string output = BrowserExports.ExportForTests(Sample, new BrowserExportRequest {
            Format = format,
            Preset = BrowserPreset.Rich
        });

        Assert.NotEmpty(output);
        Assert.Contains(expectedFragment, output);
    }

    [Fact]
    public void ExplicitOptionsOverrideThePreset()
    {
        string output = BrowserExports.ExportForTests(Sample, new BrowserExportRequest {
            Format = BrowserExportFormat.Html,
            Preset = BrowserPreset.Plain,
            Html = new BrowserHtmlOptions { RootCssClass = "browser-test-root", EmitImages = true }
        });

        Assert.Contains("browser-test-root", output);
    }

    [Fact]
    public void ResolvesDocVariableIntoAValidDocx()
    {
        byte[] source = File.ReadAllBytes(Path.Combine(ProjectRoot, "samples", "TestDocVariables.docx"));
        byte[] output = BrowserExports.ResolveDocxForTests(source, new BrowserResolveRequest {
            Fields = new BrowserFieldOptions {
                Mode = BrowserFieldMode.Evaluate,
                Variables = new Dictionary<string, string?> { ["Var1"] = "browser-value" }
            }
        });

        using var stream = new MemoryStream(output, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.NotNull(document.MainDocumentPart?.Document);
        Assert.Contains("browser-value", document.MainDocumentPart!.Document.InnerText);
    }

    [Fact]
    public void RejectsEmptyInput()
    {
        var error = Assert.Throws<ArgumentException>(() => BrowserExports.ExportForTests([], new BrowserExportRequest()));
        Assert.Contains("non-empty DOCX", error.Message);
    }

    [Fact]
    public void DetectsTrackedChanges()
    {
        byte[] tracked = File.ReadAllBytes(Path.Combine(ProjectRoot, "samples", "Tracked.docx"));
        byte[] unchanged = File.ReadAllBytes(Path.Combine(ProjectRoot, "samples", "AutoNumberCompatibility.docx"));

        Assert.True(BrowserExports.InspectForTests(tracked).HasTrackedChanges);
        Assert.False(BrowserExports.InspectForTests(unchanged).HasTrackedChanges);
    }
}
