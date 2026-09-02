using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class ExportProgressTests
{
    [Fact]
    public void ReportsParagraphUnitsAndSuccessfulCompletion()
    {
        byte[] source = CreateDocument(
            new Paragraph(new Run(new Text("one"))),
            new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("two")))))));
        var reports = new List<DxpExportProgress>();
        var options = new DxpExportOptions { Progress = new InlineProgress(reports.Add) };

        string output = DxpExport.ExportToString(source, CreateVisitor(), options);

        Assert.Contains("one", output);
        Assert.Contains("two", output);
        Assert.Equal(
            [DxpExportPhase.Opening, DxpExportPhase.Preparing, DxpExportPhase.Converting,
             DxpExportPhase.Finalizing, DxpExportPhase.Completed],
            reports.Select(p => p.Phase).Distinct());
        var completed = Assert.Single(reports, p => p.Phase == DxpExportPhase.Completed);
        Assert.Equal(2, completed.CompletedUnits);
        Assert.Equal(2, completed.TotalUnits);
        Assert.Equal(100d, completed.Percentage);
        Assert.All(reports.Where(p => p.Percentage.HasValue), p => Assert.InRange(p.Percentage!.Value, 0, 100));
    }

    [Fact]
    public void PreparingHasNoPercentageAndZeroParagraphDocumentCompletes()
    {
        byte[] source = CreateDocument();
        var reports = new List<DxpExportProgress>();

        DxpExport.ExportToString(source, CreateVisitor(),
            new DxpExportOptions { Progress = new InlineProgress(reports.Add) });

        Assert.All(reports.Where(p => p.Phase is DxpExportPhase.Opening or DxpExportPhase.Preparing),
            p => Assert.Null(p.Percentage));
        var completed = reports.Last();
        Assert.Equal(DxpExportPhase.Completed, completed.Phase);
        Assert.Equal(0, completed.TotalUnits);
        Assert.Equal(100d, completed.Percentage);
    }

    [Fact]
    public void OutputIsUnchangedWhenProgressIsEnabled()
    {
        byte[] source = CreateDocument(new Paragraph(new Run(new Text("same output"))));
        string withoutProgress = DxpExport.ExportToString(source, CreateVisitor());
        string withProgress = DxpExport.ExportToString(source, CreateVisitor(),
            new DxpExportOptions { Progress = new InlineProgress(_ => { }) });

        Assert.Equal(withoutProgress, withProgress);
    }

    [Fact]
    public void FailedExportDoesNotReportCompletion()
    {
        byte[] source = CreateDocument(new Paragraph(new Run(new Text("fail"))));
        var reports = new List<DxpExportProgress>();

        using var stream = new MemoryStream(source, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Throws<InvalidOperationException>(() =>
            DxpExport.ExportToBytes(document, new ThrowingVisitor(),
                new DxpExportOptions { Progress = new InlineProgress(reports.Add) }));

        Assert.DoesNotContain(reports, p => p.Phase == DxpExportPhase.Completed);
    }

    private static byte[] CreateDocument(params OpenXmlElement[] blocks)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(blocks));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static DxpPlainTextVisitor CreateVisitor() =>
        new(new DxpPlainTextVisitorConfig());

    private sealed class InlineProgress(Action<DxpExportProgress> report) : IProgress<DxpExportProgress>
    {
        public void Report(DxpExportProgress value) => report(value);
    }

    private sealed class ThrowingVisitor : DxpVisitor
    {
        public ThrowingVisitor() : base(null) { }

        public override IDisposable VisitParagraphBegin(
            Paragraph p,
            DxpIDocumentContext d,
            DxpIParagraphContext paragraph) => throw new InvalidOperationException("Expected test failure.");
    }
}
