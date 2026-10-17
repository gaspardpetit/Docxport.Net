using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Visitors.Pdf;
using PdfSharp.Fonts;

namespace DocxportNet.Tests;

public sealed class PdfTextExportTests
{
    [Fact]
    public void ExportsParagraphsAndExplicitPageBreak()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("First paragraph"))),
                new Paragraph(new Run(new Break { Type = BreakValues.Page }, new Text("Second page")))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpPdfVisitor());

        Assert.StartsWith("%PDF", System.Text.Encoding.ASCII.GetString(result, 0, 4));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("First paragraph", pdf.GetPage(1).Text);
        Assert.Contains("Second page", pdf.GetPage(2).Text);
    }

    [Fact]
    public void WrapsLongParagraphAcrossPages()
    {
        GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text(
                string.Join(" ", Enumerable.Repeat("wrapped text", 3000)))))));
        }

        byte[] result = DxpExport.ExportToBytes(source.ToArray(), new DxpPdfVisitor());

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(result));
        Assert.True(pdf.NumberOfPages > 1);
        Assert.Contains("wrapped", pdf.GetPage(1).Text);
        Assert.Contains("wrapped", pdf.GetPage(pdf.NumberOfPages).Text);
    }
}
