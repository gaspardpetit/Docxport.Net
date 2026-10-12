using System.IO.Compression;
using System.Text;
using DocxportNet.Visitors.Doc;
using DocxportNet.Visitors.Docx;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;
using DocxportNet.Walker.Context;

namespace DocxportNet.Tests;

public class MeasurementExportTests
{
    [Theory]
    [InlineData("1440", 1440)]
    [InlineData("1in", 1440)]
    [InlineData("0.5in", 720)]
    [InlineData("6pt", 120)]
    [InlineData("2.54cm", 1440)]
    [InlineData("25.4mm", 1440)]
    [InlineData("1pc", 240)]
    [InlineData("-0.5in", -720)]
    public void WordMeasurementsNormalizeToTwips(string raw, int twips)
    {
        var measurement = DxpMeasurement.Parse(raw);
        Assert.Equal(twips, measurement.Twips);
        Assert.Equal(twips / 20.0, measurement.Points);
        Assert.Equal(twips / 1440.0, measurement.Inches);
    }

    [Theory]
    [InlineData("markdown")]
    [InlineData("html")]
    [InlineData("text")]
    [InlineData("docx")]
    [InlineData("doc")]
    public void WordToleratedUnitsExportThroughEveryVisitor(string format)
    {
        byte[] source = CreateSyntheticDocx();
        switch (format)
        {
            case "markdown":
            {
                string result = DxpExport.ExportToString(source,
                    new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig()));
                Assert.Contains("Measurement probe", result);
                Assert.Contains("width:8.5in", result);
                Assert.Contains("margin-left:36pt", result);
                Assert.Contains("margin-bottom:6pt", result);
                break;
            }
            case "html":
            {
                string result = DxpExport.ExportToString(source,
                    new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()));
                Assert.Contains("Measurement probe", result);
                Assert.Contains("width:8.5in", result);
                Assert.Contains("margin-left:36pt", result);
                Assert.Contains("margin-bottom:6pt", result);
                break;
            }
            case "text":
            {
                string result = DxpExport.ExportToString(source,
                    new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()));
                Assert.Contains("Measurement probe", result);
                break;
            }
            case "docx":
            {
                byte[] result = DxpExport.ExportToBytes(source, new DxpDocxVisitor());
                Assert.Contains("Measurement probe", DxpExport.ExportToString(result,
                    new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())));
                break;
            }
            case "doc":
            {
                byte[] result = DxpExport.ExportToBytes(source, new DxpDocVisitor());
                Assert.Contains("Measurement probe", DxpExport.ExportToString(result,
                    new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())));
                byte[] roundTrip = DxpExport.ExportToBytes(result, new DxpDocxVisitor());
                using var package = new ZipArchive(new MemoryStream(roundTrip), ZipArchiveMode.Read);
                using var reader = new StreamReader(package.GetEntry("word/document.xml")!.Open());
                string xml = reader.ReadToEnd();
                Assert.Contains("w:w=\"2880\"", xml);
                Assert.Equal(2, xml.Split("<w:gridCol w:w=\"1440\"", StringSplitOptions.None).Length - 1);
                break;
            }
        }
    }

    private static byte[] CreateSyntheticDocx()
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(zip, "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                "</Types>");
            Write(zip, "_rels/.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                "</Relationships>");
            Write(zip, "word/document.xml",
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>" +
                "<w:p><w:pPr><w:ind w:left=\"0.5in\"/><w:spacing w:after=\"6pt\"/>" +
                "<w:tabs><w:tab w:val=\"left\" w:pos=\"1in\"/></w:tabs></w:pPr>" +
                "<w:r><w:t>Measurement probe</w:t></w:r></w:p>" +
                "<w:tbl><w:tblPr><w:tblW w:type=\"dxa\" w:w=\"2in\"/></w:tblPr>" +
                "<w:tblGrid><w:gridCol w:w=\"1in\"/><w:gridCol w:w=\"1in\"/></w:tblGrid>" +
                "<w:tr><w:tc><w:p><w:r><w:t>Cell A</w:t></w:r></w:p></w:tc>" +
                "<w:tc><w:p><w:r><w:t>Cell B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                "<w:sectPr><w:pgSz w:w=\"8.5in\" w:h=\"11in\"/>" +
                "<w:pgMar w:top=\"1in\" w:right=\"1in\" w:bottom=\"1in\" w:left=\"1in\" " +
                "w:header=\"0.5in\" w:footer=\"0.5in\" w:gutter=\"0\"/></w:sectPr>" +
                "</w:body></w:document>");
        }
        return output.ToArray();
    }

    private static void Write(ZipArchive zip, string path, string content)
    {
        using var writer = new StreamWriter(zip.CreateEntry(path).Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}
