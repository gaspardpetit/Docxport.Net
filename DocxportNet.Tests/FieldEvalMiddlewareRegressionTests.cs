using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocxportNet.Tests.Utils;
using DocxportNet.Visitors.Html;
using Xunit.Abstractions;

namespace DocxportNet.Tests;

public sealed class FieldEvalMiddlewareRegressionTests : TestBase<FieldEvalMiddlewareRegressionTests>
{
    private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
    private static readonly string SamplesDirectory = Path.Combine(ProjectRoot, "samples");

    public FieldEvalMiddlewareRegressionTests(ITestOutputHelper output) : base(output)
    {
    }

    [Fact]
    public void Eval_TestFields_InlineIfPreservesFormatting()
    {
        string docxPath = Path.Combine(SamplesDirectory, "TestFields.docx");
        var config = DxpHtmlVisitorConfig.CreateRichConfig();
        var visitor = new DxpHtmlVisitor(config, Logger);
        var options = new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate };

        string html = TestCompare.Normalize(DxpExport.ExportToString(docxPath, visitor, options, Logger));

        Assert.Contains("Expect No Error: Not Empty", html, StringComparison.Ordinal);
        Assert.Contains("Expect <strong class=\"dxp-bold\">one</strong> (bold):", html, StringComparison.Ordinal);
        Assert.Contains("Expect <strong class=\"dxp-bold\">1</strong><span class=\"dxp-underline\">2</span><strong class=\"dxp-bold\">3: 1</strong><span class=\"dxp-underline\">2</span><strong class=\"dxp-bold\">3</strong>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Eval_DocVariableWithoutSeparate_EmitsStyledValue()
    {
        using var doc = CreateDocVariableDoc("TokenLabel", "VALUE");
        var config = DxpHtmlVisitorConfig.CreateRichConfig();
        var visitor = new DxpHtmlVisitor(config, Logger);
        visitor.FieldEval.Context.SetDocVariable("TokenLabel", "VALUE");
        var options = new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate };

        string html = TestCompare.Normalize(DxpExport.ExportToString(doc, visitor, options, Logger));

        Assert.Contains("<strong class=\"dxp-bold\">VALUE</strong>", html, StringComparison.Ordinal);
    }

    private static WordprocessingDocument CreateDocVariableDoc(string name, string value)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            var runProps = new RunProperties(new Bold());
            var paragraph = new Paragraph(
                new Run(runProps.CloneNode(true), new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(runProps.CloneNode(true), new FieldCode { Text = $" DOCVARIABLE {name} " }),
                new Run(runProps.CloneNode(true), new FieldChar { FieldCharType = FieldCharValues.End })
            );

            mainPart.Document.Body!.Append(paragraph);
            mainPart.Document.Save();
            document.Save();
        }

        stream.Position = 0;
        return WordprocessingDocument.Open(stream, false);
    }
}
