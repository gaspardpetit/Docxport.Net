using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Wasm;
using M = DocumentFormat.OpenXml.Math;

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

    [Theory]
    [InlineData(BrowserExportFormat.Html, BrowserMathOutputFormat.Latex, @"\frac{a}{b}")]
    [InlineData(BrowserExportFormat.Markdown, BrowserMathOutputFormat.UnicodeMath, "(a)/(b)")]
    [InlineData(BrowserExportFormat.Text, BrowserMathOutputFormat.Text, "(a)/(b)")]
    public void BrowserOptionsSelectMathOutput(
        BrowserExportFormat exportFormat,
        BrowserMathOutputFormat mathFormat,
        string expected)
    {
        BrowserExportRequest request = new() { Format = exportFormat, Preset = BrowserPreset.Plain };
        request.Html = new BrowserHtmlOptions { MathOutputFormat = mathFormat };
        request.Markdown = new BrowserMarkdownOptions { MathOutputFormat = mathFormat, EmitMathDelimiters = false };
        request.Text = new BrowserTextOptions { MathOutputFormat = mathFormat };

        Assert.Contains(expected, BrowserExports.ExportForTests(CreateMathDocument(), request), StringComparison.Ordinal);
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

    private static byte[] CreateMathDocument()
    {
        using MemoryStream stream = new();
        using (WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new M.OfficeMath(new M.Fraction(
                    new M.Numerator(new M.Run(new M.Text("a"))),
                    new M.Denominator(new M.Run(new M.Text("b"))))))));
            main.Document.Save();
        }
        return stream.ToArray();
    }
}
