using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Fields;

namespace DocxportNet.Tests;

public sealed class DocxFieldPolicyTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SuppliedEvaluatorControlsAskRefAndNestedIfWithNullOrEmptyOptions(
        bool suppliedEvaluator, bool emptyOptions)
    {
        var calls = 0;
        var evaluator = suppliedEvaluator ? new DxpFieldEval(new DxpFieldEvalDelegates
        {
            AskAsync = (_, _) =>
            {
                calls++;
                return Task.FromResult<DxpFieldValue?>(new DxpFieldValue("12345"));
            }
        }) : null;
        var output = DxpDocxExport.Export(CreateTemplate(),
            emptyOptions ? new DxpExportOptions() : null, fieldEval: evaluator);
        using var stream = new MemoryStream(output);
        using var document = WordprocessingDocument.Open(stream, false);
        var paragraphs = document.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>().Select(p => string.Concat(p.Descendants<Text>()
                .Select(t => t.Text))).ToArray();
        Assert.Equal(3, paragraphs.Length);
        Assert.Equal(string.Empty, paragraphs[0]);
        Assert.Equal(suppliedEvaluator ? "12345" : "cached-ref", paragraphs[1]);
        Assert.Equal(suppliedEvaluator ? "12345" : "cached-if", paragraphs[2]);
        Assert.Equal(suppliedEvaluator ? 1 : 0, calls);
    }

    [Theory]
    [InlineData(DxpDocxFieldPolicy.Preserve, true, "cached-ref", "cached-if")]
    [InlineData(DxpDocxFieldPolicy.Evaluate, false, "12345", "12345")]
    public void ExplicitDocxPolicyOverridesEvaluatorPresence(DxpDocxFieldPolicy policy,
        bool suppliedEvaluator, string expectedRef, string expectedIf)
    {
        var evaluator = suppliedEvaluator ? new DxpFieldEval(new DxpFieldEvalDelegates
        {
            AskAsync = (_, _) => Task.FromResult<DxpFieldValue?>(new DxpFieldValue("12345"))
        }) : null;
        var output = DxpDocxExport.Export(CreateTemplate(useSet: !suppliedEvaluator),
            new DxpExportOptions { DocxFieldPolicy = policy }, fieldEval: evaluator);
        using var stream = new MemoryStream(output);
        using var document = WordprocessingDocument.Open(stream, false);
        var paragraphs = document.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>().Select(p => string.Concat(p.Descendants<Text>()
                .Select(t => t.Text))).ToArray();
        Assert.Equal(expectedRef, paragraphs[1]);
        Assert.Equal(expectedIf, paragraphs[2]);
    }

    [Fact]
    public void ExplicitEvaluateModeWithoutSuppliedEvaluatorStillEvaluatesFields()
    {
        var output = DxpDocxExport.Export(CreateTemplate(useSet: true),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.Evaluate });
        using var stream = new MemoryStream(output);
        using var document = WordprocessingDocument.Open(stream, false);
        var paragraphs = document.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>().Select(p => string.Concat(p.Descendants<Text>()
                .Select(t => t.Text))).ToArray();
        Assert.Equal("12345", paragraphs[1]);
        Assert.Equal("12345", paragraphs[2]);
    }

    private static byte[] CreateTemplate(bool useSet = false)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                Field(useSet ? " SET Answer \"12345\" " : " ASK Answer \"Prompt\" ", "cached-ask"),
                Field(" REF Answer ", "cached-ref"),
                Field(" IF { REF Answer } = \"12345\" \"12345\" \"wrong\" ", "cached-if")));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static Paragraph Field(string instruction, string cached) => new(
        new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
        new Run(new FieldCode(instruction) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
        new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
        new Run(new Text(cached)),
        new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
}
