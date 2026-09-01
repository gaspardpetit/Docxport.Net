using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Omml;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;
using M = DocumentFormat.OpenXml.Math;

namespace DocxportNet.Tests.Omml;

public sealed class DxpOmmlPipelineIntegrationTests
{
    [Fact]
    public void ExporterPresetsDeclareSensibleMathDefaults()
    {
        Assert.Equal(DxpOmmlOutputFormat.MathMl, DxpHtmlVisitorConfig.CreateRichConfig().MathOutputFormat);
        Assert.Equal(DxpOmmlOutputFormat.MathMl, DxpHtmlVisitorConfig.CreatePlainConfig().MathOutputFormat);
        Assert.Equal(DxpOmmlOutputFormat.Latex, DxpMarkdownVisitorConfig.CreateRichConfig().MathOutputFormat);
        Assert.Equal(DxpOmmlOutputFormat.Latex, DxpMarkdownVisitorConfig.CreatePlainConfig().MathOutputFormat);
        Assert.Equal(DxpOmmlOutputFormat.Text, DxpPlainTextVisitorConfig.CreateAcceptConfig().MathOutputFormat);
        Assert.Equal(DxpOmmlOutputFormat.Text, DxpPlainTextVisitorConfig.CreateRejectConfig().MathOutputFormat);
    }

    [Fact]
    public void EveryTextExporterUsesTheSelectedStandaloneFormat()
    {
        M.OfficeMath equation = Fraction("a", "b");
        string unicodeMath = DxpOmmlConverter.Convert(equation, DxpOmmlOutputFormat.UnicodeMath).Output;
        string latex = DxpOmmlConverter.Convert(equation, DxpOmmlOutputFormat.Latex).Output;
        string readable = DxpOmmlConverter.Convert(equation, DxpOmmlOutputFormat.Text).Output;

        string html = Export(new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig() with
        {
            MathOutputFormat = DxpOmmlOutputFormat.UnicodeMath,
        }), equation);
        string markdown = Export(new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig() with
        {
            MathOutputFormat = DxpOmmlOutputFormat.UnicodeMath,
            EmitMathDelimiters = false,
        }), equation);
        string text = Export(new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig() with
        {
            MathOutputFormat = DxpOmmlOutputFormat.Latex,
        }), equation);
        string defaultText = Export(new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()), equation);

        Assert.Contains(unicodeMath, html, StringComparison.Ordinal);
        Assert.Contains(unicodeMath, markdown, StringComparison.Ordinal);
        Assert.Contains(latex, text, StringComparison.Ordinal);
        Assert.Contains(readable, defaultText, StringComparison.Ordinal);
    }

    [Fact]
    public void ConfiguredLatexUsesWalkerResolverForEmbeddedWordprocessingMl()
    {
        M.OfficeMath equation = new(new Hyperlink(new Run(new Text("A_B"))));
        DxpHtmlVisitorConfig htmlConfig = DxpHtmlVisitorConfig.CreatePlainConfig() with
        {
            MathOutputFormat = DxpOmmlOutputFormat.Latex,
        };
        DxpPlainTextVisitorConfig textConfig = DxpPlainTextVisitorConfig.CreateAcceptConfig() with
        {
            MathOutputFormat = DxpOmmlOutputFormat.Latex,
        };

        Assert.Contains(@"A\_B", Export(new DxpHtmlVisitor(htmlConfig), equation), StringComparison.Ordinal);
        Assert.Contains(@"A\_B", Export(new DxpPlainTextVisitor(textConfig), equation), StringComparison.Ordinal);
    }

    [Fact]
    public void NullMathFormatOmitsMathInEveryTextExporter()
    {
        M.OfficeMath equation = Fraction("visible", "content");

        string html = Export(new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig() with { MathOutputFormat = null }), equation);
        string markdown = Export(new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig() with { MathOutputFormat = null }), equation);
        string text = Export(new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig() with { MathOutputFormat = null }), equation);

        Assert.DoesNotContain("visible", html, StringComparison.Ordinal);
        Assert.DoesNotContain("visible", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("visible", text, StringComparison.Ordinal);
    }

    private static M.OfficeMath Fraction(string numerator, string denominator) => new(
        new M.Fraction(
            new M.Numerator(new M.Run(new M.Text(numerator))),
            new M.Denominator(new M.Run(new M.Text(denominator)))));

    private static string Export(DxpITextVisitor visitor, M.OfficeMath equation)
    {
        using MemoryStream stream = new();
        using (WordprocessingDocument document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(equation.CloneNode(true))));
            main.Document.Save();
        }

        stream.Position = 0;
        using WordprocessingDocument input = WordprocessingDocument.Open(stream, false);
        return DxpExport.ExportToString(input, visitor, new DxpExportOptions());
    }
}
