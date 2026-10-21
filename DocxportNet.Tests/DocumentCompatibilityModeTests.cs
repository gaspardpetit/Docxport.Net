using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors;
using DocxportNet.Walker;

namespace DocxportNet.Tests;

public class DocumentCompatibilityModeTests
{
    [Theory]
    [InlineData(null, 12)]
    [InlineData("11", 11)]
    [InlineData("12", 12)]
    [InlineData("14", 14)]
    [InlineData("15", 15)]
    [InlineData("16", 16)]
    [InlineData("invalid", 12)]
    public void VisitorReceivesResolvedCompatibilityMode(string? settingValue, int expected)
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(new Paragraph(new Run(new Text("probe")))));
        if (settingValue != null)
        {
            var settings = main.AddNewPart<DocumentSettingsPart>();
            settings.Settings = new Settings(new Compatibility(new CompatibilitySetting
            {
                Name = CompatSettingNameValues.CompatibilityMode,
                Uri = "http://schemas.microsoft.com/office/word",
                Val = settingValue
            }));
        }

        var visitor = new RecordingVisitor();
        new DxpWalker().Accept(document, visitor);

        Assert.Equal(expected, visitor.CompatibilityMode);
    }

    private sealed class RecordingVisitor : DxpVisitor
    {
        public RecordingVisitor() : base(null) { }
        public int CompatibilityMode { get; private set; }

        public override IDisposable VisitDocumentBegin(WordprocessingDocument doc,
            DxpIDocumentContext documentContext)
        {
            CompatibilityMode = documentContext.CompatibilityMode;
            return base.VisitDocumentBegin(doc, documentContext);
        }
    }
}
