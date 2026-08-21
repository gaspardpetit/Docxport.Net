using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Resolution;
using DocxportNet.Middleware;
using DocxportNet.Tests.Utils;
using DocxportNet.Visitors;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Walker;
using System.Xml.Linq;
using Xunit.Abstractions;
using Xunit.Sdk;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxportNet.Tests;

public class HtmlExportTests : TestBase<HtmlExportTests>
{
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

        public void Serialize(IXunitSerializationInfo info) => info.AddValue(nameof(DocxPath), DocxPath);
        public void Deserialize(IXunitSerializationInfo info) => DocxPath = info.GetValue<string>(nameof(DocxPath));

        public override string ToString() => FileName;
    }

    private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
    private static readonly string SamplesDirectory = Path.Combine(ProjectRoot, "samples");

    public HtmlExportTests(ITestOutputHelper output) : base(output)
    {
    }

    public static IEnumerable<object[]> SampleDocs()
    {
        return Directory.EnumerateFiles(SamplesDirectory, "*.docx", SearchOption.TopDirectoryOnly)
            .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
            .OrderBy(Path.GetFileName)
            .Select(path => new object[] { new Sample(path) });
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToHtml_Accept(Sample sample)
    {
        VerifyAgainstFixture(sample, DxpHtmlVisitorConfig.CreateRichConfig(), ".html", ".test.html", DxpTrackedChangeMode.AcceptChanges, DxpFieldEvalExportMode.None);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToHtml_Reject(Sample sample)
    {
        VerifyAgainstFixture(sample, DxpHtmlVisitorConfig.CreateRichConfig(), ".reject.html", ".reject.test.html", DxpTrackedChangeMode.RejectChanges, DxpFieldEvalExportMode.None);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToHtml_Cached(Sample sample)
    {
        VerifyCachedAgainstFixture(sample, DxpHtmlVisitorConfig.CreateRichConfig(), ".cached.html", ".cached.test.html");
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToHtml_Eval(Sample sample)
    {
        VerifyAgainstFixture(sample, DxpHtmlVisitorConfig.CreateRichConfig(), ".eval.html", ".eval.test.html", DxpTrackedChangeMode.AcceptChanges, DxpFieldEvalExportMode.Evaluate);
    }

    [Fact]
    public void HtmlExport_NonBreakingSpaceRetainsParagraphWrapper()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Before</w:t></w:r></w:p>
  <w:p><w:r><w:t xml:space="preserve">&#160;</w:t></w:r></w:p>
  <w:p><w:r><w:t>After</w:t></w:r></w:p>
</w:body>
""";

        string html = ExportHtmlFromBodyXml(bodyXml, DxpHtmlVisitorConfig.CreateRichConfig());

        _ = XDocument.Parse(html);
        Assert.Contains("<p class=\"dxp-paragraph\">&#160;</p>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("</p>\n&#160;\n<p", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_RendersFootnotesInsideSectionBeforeFooter()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var footerPart = main.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(
                new Paragraph(
                    new Run(new Text("Footer text"))));

            var footnotesPart = main.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes = new Footnotes(
                new Footnote(
                    new Paragraph(
                        new Run(new FootnoteReferenceMark()),
                        new Run(new Text(" Footnote text"))))
                {
                    Id = 1
                });

            main.Document = new Document(
                new Body(
                    new Paragraph(
                        new Run(new Text("Body text")),
                        new Run(new FootnoteReference { Id = 1 })),
                    new SectionProperties(
                        new FooterReference { Id = main.GetIdOfPart(footerPart), Type = HeaderFooterValues.Default },
                        new PageSize { Width = 12240U, Height = 15840U },
                        new PageMargin { Top = 1440, Right = 1440U, Bottom = 1440, Left = 1440U, Header = 720U, Footer = 720U, Gutter = 0U })));

            main.Document.Save();
            footerPart.Footer.Save();
            footnotesPart.Footnotes.Save();
        }

        stream.Position = 0;

        using var readDoc = WordprocessingDocument.Open(stream, false);
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger);
        var html = TestCompare.Normalize(DxpExport.ExportToString(
            readDoc,
            visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger));

        Assert.Contains(".dxp-footnotes::before", html);

        int bodyIndex = html.IndexOf("Body text", StringComparison.Ordinal);
        int footnotesIndex = html.IndexOf("""<div class="dxp-footnotes">""", StringComparison.Ordinal);
        int footnoteTextIndex = html.IndexOf("Footnote text", StringComparison.Ordinal);
        int footerIndex = html.IndexOf("""<div class="dxp-footer""", StringComparison.Ordinal);

        Assert.True(bodyIndex >= 0, "Body text should be present.");
        Assert.True(footnotesIndex > bodyIndex, "Footnotes should render after the body content.");
        Assert.True(footnoteTextIndex > footnotesIndex, "Footnote text should render inside the footnotes block.");
        Assert.True(footerIndex > footnotesIndex, "Footnotes should render before the footer within the section canvas.");
    }

    [Fact]
    public void HtmlExport_PositiveMarginsUseNaturalFlowHeaderAndFooterRegions()
    {
        string html = ExportHeaderFooterLayout(
            top: 1440,
            bottom: 1440,
            headerDistance: 720,
            footerDistance: 720,
            header: new Header(
                new Paragraph(new Run(new Text("Header line one"))),
                new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("Header table"))))))),
            footer: new Footer(
                new Paragraph(new Run(new Text("Footer line one"))),
                new Paragraph(new Run(new Text("Footer line two")))));

        Assert.Contains("class=\"dxp-header\" style=\"min-height:1in;", html, StringComparison.Ordinal);
        Assert.Contains("margin-bottom:-1in;padding-top:0.5in;flex:0 0 auto;", html, StringComparison.Ordinal);
        Assert.Contains("class=\"dxp-footer\" style=\"display:flex;flex-direction:column;justify-content:flex-end;min-height:1in;", html, StringComparison.Ordinal);
        Assert.Contains("margin-top:-1in;padding-bottom:0.5in;flex:0 0 auto;", html, StringComparison.Ordinal);
        Assert.Contains("flex:0 0 auto", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-header\" style=\"height:", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-footer\" style=\"display:flex;flex-direction:column;justify-content:flex-end;height:", html, StringComparison.Ordinal);

        Assert.True(html.IndexOf("Header table", StringComparison.Ordinal) < html.IndexOf("Body text", StringComparison.Ordinal));
        Assert.True(html.IndexOf("Body text", StringComparison.Ordinal) < html.IndexOf("Footer line one", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlExport_MissingHeaderAndFooterPreserveBodyMargins()
    {
        string html = ExportHeaderFooterLayout(
            top: 1440,
            bottom: 2160,
            headerDistance: 720,
            footerDistance: 720);

        Assert.Contains("class=\"dxp-body\" style=\"flex:1 0 auto;padding-top:1in;padding-bottom:1.5in;\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-header\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-footer\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_NegativeMarginsKeepHeaderAndFooterOutOfBodyFlow()
    {
        string html = ExportHeaderFooterLayout(
            top: -1440,
            bottom: -2160,
            headerDistance: 720,
            footerDistance: 360,
            header: new Header(new Paragraph(new Run(new Text("Overlay header")))),
            footer: new Footer(new Paragraph(new Run(new Text("Overlay footer")))));

        Assert.Contains("class=\"dxp-header\" style=\"position:absolute;top:0.5in;", html, StringComparison.Ordinal);
        Assert.Contains("class=\"dxp-footer\" style=\"position:absolute;bottom:0.25in;", html, StringComparison.Ordinal);
        Assert.Contains("class=\"dxp-body\" style=\"flex:1 0 auto;padding-top:1in;padding-bottom:1.5in;\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-header\" style=\"min-height:", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-footer\" style=\"display:flex", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_TallInlineHeaderImageExpandsFlowAndPreservesExtent()
    {
        string html = ExportHeaderFooterLayout(
            top: 1440,
            bottom: 1440,
            headerDistance: 720,
            footerDistance: 720,
            header: CreateHeaderWithInlineImage(5943600L, 3346450L));

        Assert.Contains("class=\"dxp-header\" style=\"min-height:1in;", html, StringComparison.Ordinal);
        Assert.Contains("margin-bottom:-1in;padding-top:0.5in;flex:0 0 auto;", html, StringComparison.Ordinal);
        Assert.Contains("class=\"dxp-image\"", html, StringComparison.Ordinal);
        Assert.Contains("style=\"width:468pt;height:263.5pt;max-width:none;\"", html, StringComparison.Ordinal);
        Assert.True(html.IndexOf("dxp-image", StringComparison.Ordinal) < html.IndexOf("Body text", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlExport_EmptyHeaderAndFooterStillPreserveMinimumMargins()
    {
        string html = ExportHeaderFooterLayout(
            top: 1440,
            bottom: 2160,
            headerDistance: 720,
            footerDistance: 360,
            header: new Header(new Paragraph()),
            footer: new Footer(new Paragraph()));

        Assert.Contains("class=\"dxp-header\" style=\"min-height:1in;", html, StringComparison.Ordinal);
        Assert.Contains("margin-bottom:-1in;padding-top:0.5in;flex:0 0 auto;", html, StringComparison.Ordinal);
        Assert.Contains("class=\"dxp-footer\" style=\"display:flex;flex-direction:column;justify-content:flex-end;min-height:1.5in;", html, StringComparison.Ordinal);
        Assert.Contains("margin-top:-1.5in;padding-bottom:0.25in;flex:0 0 auto;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_HeaderAndFooterOffsetsCanExceedBodyMargins()
    {
        string html = ExportHeaderFooterLayout(
            top: 720,
            bottom: 720,
            headerDistance: 1440,
            footerDistance: 1440,
            header: new Header(new Paragraph(new Run(new Text("Offset header")))),
            footer: new Footer(new Paragraph(new Run(new Text("Offset footer")))));

        Assert.Contains("min-height:0.5in;margin-bottom:-0.5in;padding-top:1in;", html, StringComparison.Ordinal);
        Assert.Contains("min-height:0.5in;margin-top:-0.5in;padding-bottom:1in;", html, StringComparison.Ordinal);
        Assert.Contains("padding-top:0.5in;padding-bottom:0.5in;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_TallInlineFooterImageExpandsFlowAndPreservesExtent()
    {
        string html = ExportHeaderFooterLayout(
            top: 1440,
            bottom: 1440,
            headerDistance: 720,
            footerDistance: 720,
            footer: new Footer(new Paragraph(new Run(CreateInlineImageDrawing(5943600L, 3346450L)))));

        Assert.Contains("class=\"dxp-footer\" style=\"display:flex;flex-direction:column;justify-content:flex-end;min-height:1in;", html, StringComparison.Ordinal);
        Assert.Contains("margin-top:-1in;padding-bottom:0.5in;flex:0 0 auto;", html, StringComparison.Ordinal);
        Assert.Contains("style=\"width:468pt;height:263.5pt;max-width:none;\"", html, StringComparison.Ordinal);
        Assert.True(html.IndexOf("Body text", StringComparison.Ordinal) < html.LastIndexOf("dxp-image", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlExport_RendersCropRotationFlipAndAccessibilityMetadata()
    {
        string html = ExportHeaderFooterLayout(
            top: 1440,
            bottom: 1440,
            headerDistance: 720,
            footerDistance: 720,
            header: new Header(new Paragraph(new Run(CreateInlineImageDrawing(
                1270000L,
                635000L,
                new A.SourceRectangle { Left = 10000, Top = 20000, Right = 30000, Bottom = 10000 },
                rotation: 5400000,
                flipHorizontal: true,
                flipVertical: true,
                description: "Accessible description",
                title: "Image title")))));

        Assert.Contains("class=\"dxp-image-frame\"", html, StringComparison.Ordinal);
        Assert.Contains("width:100pt;height:50pt;overflow:hidden;transform:rotate(90deg) scaleX(-1) scaleY(-1);transform-origin:center;", html, StringComparison.Ordinal);
        Assert.Contains("alt=\"Accessible description\" title=\"Image title\"", html, StringComparison.Ordinal);
        Assert.Contains("width:166.667pt;height:71.429pt;left:-16.667pt;top:-14.286pt;", html, StringComparison.Ordinal);
    }

    [Fact(Skip = "Floating DrawingML requires a positioned overlay and wrap-aware layout model; this change covers inline flow content.")]
    public void HtmlExport_AnchoredHeaderAndFooterDrawingsRespectWordWrappingExtents()
    {
    }

    [Fact]
    public void HtmlExport_RendersHyperlinkFieldAsActiveAnchor()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(
                new Body(
                    new Paragraph(
                        new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                        new Run(new FieldCode(""" HYPERLINK "https://openparliament.ca/committees/industry/44-1/49/catherine-lovrics-2/" """) { Space = SpaceProcessingModeValues.Preserve }),
                        new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                        new Run(new Text("https://openparliament.ca/committees/industry/44-1/49/catherine-lovrics-2/")),
                        new Run(new FieldChar { FieldCharType = FieldCharValues.End }))));
            main.Document.Save();
        }

        stream.Position = 0;

        using var readDoc = WordprocessingDocument.Open(stream, false);
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger);
        var html = TestCompare.Normalize(DxpExport.ExportToString(
            readDoc,
            visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger));

        Assert.Contains("""<a class="dxp-link" href="https://openparliament.ca/committees/industry/44-1/49/catherine-lovrics-2/">""", html);
        Assert.Contains("https://openparliament.ca/committees/industry/44-1/49/catherine-lovrics-2/", html);
        Assert.DoesNotContain("data-field=\"HYPERLINK &quot;https://openparliament.ca/committees/industry/44-1/49/catherine-lovrics-2/&quot;\"", html);
    }

    [Fact]
    public void HtmlExport_CacheMode_DefaultFieldFallback_ReplaysCachedResults_AndSuppressesOnlyTruePageFields()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Unknown: </w:t></w:r>
    <w:fldSimple w:instr=" FOO ">
      <w:r><w:t>cached unknown</w:t></w:r>
    </w:fldSimple>
  </w:p>
  <w:p>
    <w:r><w:t xml:space="preserve">TOC: </w:t></w:r>
    <w:fldSimple w:instr=" TOC \o &quot;1-3&quot; ">
      <w:r><w:t>Heading 1</w:t></w:r>
      <w:r><w:tab/></w:r>
      <w:r><w:t>3</w:t></w:r>
    </w:fldSimple>
  </w:p>
  <w:p>
    <w:r><w:t xml:space="preserve">Page: </w:t></w:r>
    <w:fldSimple w:instr=" PAGE ">
      <w:r><w:t>4</w:t></w:r>
    </w:fldSimple>
  </w:p>
  <w:p>
    <w:r><w:t xml:space="preserve">PageRef: </w:t></w:r>
    <w:fldSimple w:instr=" PAGEREF Bookmark1 \h ">
      <w:r><w:t>7</w:t></w:r>
    </w:fldSimple>
  </w:p>
</w:body>
""";

        var html = TestCompare.Normalize(ExportHtmlCachedFromBodyXml(bodyXml, DxpHtmlVisitorConfig.CreateRichConfig()));

        Assert.Contains("Unknown:", html, StringComparison.Ordinal);
        Assert.Contains("cached unknown", html, StringComparison.Ordinal);
        Assert.Contains("TOC:", html, StringComparison.Ordinal);
        Assert.Contains("Heading 1", html, StringComparison.Ordinal);
        Assert.Contains("3", html, StringComparison.Ordinal);
        Assert.Contains("Page:", html, StringComparison.Ordinal);
        Assert.DoesNotContain(">4<", html, StringComparison.Ordinal);
        Assert.Contains("PageRef:", html, StringComparison.Ordinal);
        Assert.Contains("7", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphLayout_MapsFullLeaderSet()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:tabs><w:tab w:val="right" w:leader="none" w:pos="8640"/></w:tabs></w:pPr>
    <w:r><w:t>none</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r>
  </w:p>
  <w:p>
    <w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8640"/></w:tabs></w:pPr>
    <w:r><w:t>dot</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r>
  </w:p>
  <w:p>
    <w:pPr><w:tabs><w:tab w:val="right" w:leader="hyphen" w:pos="8640"/></w:tabs></w:pPr>
    <w:r><w:t>hyphen</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r>
  </w:p>
  <w:p>
    <w:pPr><w:tabs><w:tab w:val="right" w:leader="underscore" w:pos="8640"/></w:tabs></w:pPr>
    <w:r><w:t>underscore</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r>
  </w:p>
  <w:p>
    <w:pPr><w:tabs><w:tab w:val="right" w:leader="heavy" w:pos="8640"/></w:tabs></w:pPr>
    <w:r><w:t>heavy</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r>
  </w:p>
  <w:p>
    <w:pPr><w:tabs><w:tab w:val="right" w:leader="middleDot" w:pos="8640"/></w:tabs></w:pPr>
    <w:r><w:t>middleDot</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r>
  </w:p>
</w:body>
""";

        var captured = CaptureParagraphTabStops(bodyXml);

        Assert.Collection(
            captured,
            stop => Assert.Equal(DxpComputedTabLeaderKind.None, stop.Leader),
            stop => Assert.Equal(DxpComputedTabLeaderKind.Dot, stop.Leader),
            stop => Assert.Equal(DxpComputedTabLeaderKind.Hyphen, stop.Leader),
            stop => Assert.Equal(DxpComputedTabLeaderKind.Underscore, stop.Leader),
            stop => Assert.Equal(DxpComputedTabLeaderKind.Heavy, stop.Leader),
            stop => Assert.Equal(DxpComputedTabLeaderKind.MiddleDot, stop.Leader));
    }

    [Fact]
    public void HtmlExport_RendersLeaderSpecificTabSpans()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8640"/></w:tabs></w:pPr>
    <w:r><w:t>Dot leader</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>10</w:t></w:r>
  </w:p>
  <w:p>
    <w:pPr><w:tabs><w:tab w:val="right" w:leader="hyphen" w:pos="8640"/></w:tabs></w:pPr>
    <w:r><w:t>Hyphen leader</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>11</w:t></w:r>
  </w:p>
  <w:p>
    <w:pPr><w:tabs><w:tab w:val="right" w:leader="underscore" w:pos="8640"/></w:tabs></w:pPr>
    <w:r><w:t>Underscore leader</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>12</w:t></w:r>
  </w:p>
  <w:p>
    <w:pPr><w:tabs><w:tab w:val="right" w:leader="heavy" w:pos="8640"/></w:tabs></w:pPr>
    <w:r><w:t>Heavy leader</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>13</w:t></w:r>
  </w:p>
  <w:p>
    <w:pPr><w:tabs><w:tab w:val="right" w:leader="middleDot" w:pos="8640"/></w:tabs></w:pPr>
    <w:r><w:t>Middle dot leader</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>14</w:t></w:r>
  </w:p>
  <w:p>
    <w:pPr><w:tabs><w:tab w:val="right" w:pos="8640"/></w:tabs></w:pPr>
    <w:r><w:t>No leader</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>15</w:t></w:r>
  </w:p>
</w:body>
""";

        string html = ExportHtmlFromBodyXml(bodyXml, DxpHtmlVisitorConfig.CreateRichConfig());

        Assert.Contains("radial-gradient(circle, currentColor 0.8px, transparent 1px)", html, StringComparison.Ordinal);
        Assert.Contains("linear-gradient(to right, currentColor 0, currentColor 60%, transparent 60%, transparent 100%)", html, StringComparison.Ordinal);
        Assert.Contains("border-bottom:1px solid currentColor", html, StringComparison.Ordinal);
        Assert.Contains("border-bottom:2px solid currentColor", html, StringComparison.Ordinal);
        Assert.Contains("radial-gradient(circle, currentColor 1.2px, transparent 1.45px)", html, StringComparison.Ordinal);
        Assert.Contains("Dot leader", html, StringComparison.Ordinal);
        Assert.Contains(">10<", html, StringComparison.Ordinal);
        Assert.Contains("No leader", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_ParagraphIndent_EmitsMarginLeftAndTextIndent()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    <w:r><w:t>Indented paragraph</w:t></w:r>
  </w:p>
</w:body>
""";

        var html = TestCompare.Normalize(ExportHtmlFromBodyXml(bodyXml, DxpHtmlVisitorConfig.CreateRichConfig()));

        Assert.Contains("margin-left:36pt;", html, StringComparison.Ordinal);
        Assert.Contains("text-indent:-18pt;", html, StringComparison.Ordinal);
        Assert.Contains("Indented paragraph", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_CacheMode_ComplexFieldReplay_PreservesParagraphTabLeaders()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:ind w:left="240"/>
      <w:tabs>
        <w:tab w:val="right" w:leader="dot" w:pos="8640"/>
      </w:tabs>
    </w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> TOC \o "1-3" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>Heading 1</w:t></w:r>
    <w:r><w:tab/></w:r>
    <w:r><w:t>3</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:body>
""";

        var html = TestCompare.Normalize(ExportHtmlCachedFromBodyXml(bodyXml, DxpHtmlVisitorConfig.CreateRichConfig()));

        Assert.Contains("Heading 1", html, StringComparison.Ordinal);
        Assert.Contains("3", html, StringComparison.Ordinal);
        Assert.Contains("margin-left:12pt;", html, StringComparison.Ordinal);
        Assert.Contains("white-space:nowrap;", html, StringComparison.Ordinal);
        Assert.Contains("radial-gradient(circle, currentColor 0.8px, transparent 1px)", html, StringComparison.Ordinal);
        Assert.DoesNotContain("&#9;3", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_LeftTabParagraph_DoesNotForceNowrap()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:tabs>
        <w:tab w:val="left" w:pos="2880"/>
      </w:tabs>
    </w:pPr>
    <w:r><w:t>Left</w:t></w:r>
    <w:r><w:tab/></w:r>
    <w:r><w:t>Right</w:t></w:r>
  </w:p>
</w:body>
""";

        var html = TestCompare.Normalize(ExportHtmlFromBodyXml(bodyXml, DxpHtmlVisitorConfig.CreateRichConfig()));

        Assert.Contains("Left", html, StringComparison.Ordinal);
        Assert.Contains("Right", html, StringComparison.Ordinal);
        Assert.Contains("class=\"dxp-tab\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("white-space:nowrap;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_CacheMode_TocLikeHyperlinkReplay_UsesAlignedLeaderTab()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:ind w:left="240"/>
      <w:tabs>
        <w:tab w:val="right" w:leader="dot" w:pos="8640"/>
      </w:tabs>
    </w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:hyperlink w:anchor="Bookmark1" w:history="1">
      <w:r>
        <w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>
        <w:t>Heading 1</w:t>
      </w:r>
      <w:r>
        <w:rPr><w:webHidden/></w:rPr>
        <w:tab/>
      </w:r>
      <w:r>
        <w:rPr><w:webHidden/></w:rPr>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r>
        <w:rPr><w:webHidden/></w:rPr>
        <w:instrText xml:space="preserve"> PAGEREF Bookmark1 \h </w:instrText>
      </w:r>
      <w:r>
        <w:rPr><w:webHidden/></w:rPr>
        <w:fldChar w:fldCharType="separate"/>
      </w:r>
      <w:r>
        <w:rPr><w:webHidden/></w:rPr>
        <w:t>3</w:t>
      </w:r>
      <w:r>
        <w:rPr><w:webHidden/></w:rPr>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    </w:hyperlink>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
  <w:p>
    <w:bookmarkStart w:id="0" w:name="Bookmark1"/>
    <w:r><w:t>Target</w:t></w:r>
    <w:bookmarkEnd w:id="0"/>
  </w:p>
</w:body>
""";

        var html = TestCompare.Normalize(ExportHtmlCachedFromBodyXml(bodyXml, DxpHtmlVisitorConfig.CreateRichConfig()));
        Assert.Contains("href=\"#Bookmark1\"", html, StringComparison.Ordinal);
        Assert.Contains("Heading 1", html, StringComparison.Ordinal);
        Assert.Contains("3", html, StringComparison.Ordinal);
        Assert.Contains("dxp-tab dxp-tab-right", html, StringComparison.Ordinal);
        Assert.Contains("white-space:nowrap;", html, StringComparison.Ordinal);
        Assert.Contains("radial-gradient(circle, currentColor 0.8px, transparent 1px)", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Heading 1</span>&#9;3", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_CacheMode_MultiParagraphTocField_ReplaysParagraphBoundariesAndLeaders()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="TOC1"/>
      <w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8640"/></w:tabs>
    </w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:hyperlink w:anchor="Bookmark1" w:history="1">
      <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>A</w:t></w:r>
      <w:r><w:rPr><w:webHidden/></w:rPr><w:tab/></w:r>
      <w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:rPr><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF Bookmark1 \h </w:instrText></w:r>
      <w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:webHidden/></w:rPr><w:t>1</w:t></w:r>
      <w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
    </w:hyperlink>
  </w:p>
  <w:p>
    <w:pPr>
      <w:pStyle w:val="TOC2"/>
      <w:ind w:left="240"/>
      <w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8640"/></w:tabs>
    </w:pPr>
    <w:hyperlink w:anchor="Bookmark2" w:history="1">
      <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>A.1</w:t></w:r>
      <w:r><w:rPr><w:webHidden/></w:rPr><w:tab/></w:r>
      <w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:rPr><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF Bookmark2 \h </w:instrText></w:r>
      <w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:webHidden/></w:rPr><w:t>2</w:t></w:r>
      <w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
    </w:hyperlink>
  </w:p>
  <w:p>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
  <w:p><w:bookmarkStart w:id="0" w:name="Bookmark1"/><w:r><w:t>Target 1</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>
  <w:p><w:bookmarkStart w:id="1" w:name="Bookmark2"/><w:r><w:t>Target 2</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
</w:body>
""";

        var html = TestCompare.Normalize(ExportHtmlCachedFromBodyXml(bodyXml, DxpHtmlVisitorConfig.CreateRichConfig()));

        Assert.Contains("href=\"#Bookmark1\"", html, StringComparison.Ordinal);
        Assert.Contains("href=\"#Bookmark2\"", html, StringComparison.Ordinal);
        Assert.Contains("dxp-tab dxp-tab-right", html, StringComparison.Ordinal);
        Assert.Contains("white-space:nowrap;", html, StringComparison.Ordinal);
        Assert.Contains("radial-gradient(circle, currentColor 0.8px, transparent 1px)", html, StringComparison.Ordinal);
        Assert.Contains("margin-left:12pt;", html, StringComparison.Ordinal);
        Assert.DoesNotContain("A</span>&#9;1", html, StringComparison.Ordinal);
        Assert.DoesNotContain("A.1</span>&#9;2", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_EmitParagraphMetadata_EmitsParaIdAndDocumentPart()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var paragraph = new Paragraph(
                new Run(new Text("Body text")));
            paragraph.ParagraphId = "1234ABCD";

            main.Document = new Document(new Body(paragraph));
            main.Document.Save();
        }

        stream.Position = 0;

        using var readDoc = WordprocessingDocument.Open(stream, false);
        var config = DxpHtmlVisitorConfig.CreateRichConfig();
        config.EmitParagraphMetadata = true;
        var visitor = new DxpHtmlVisitor(config, Logger);
        var html = TestCompare.Normalize(DxpExport.ExportToString(
            readDoc,
            visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger));

        Assert.Contains("data-para-id=\"1234ABCD\"", html, StringComparison.Ordinal);
        Assert.Contains("data-docx-part=\"document\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_EmitParagraphMetadata_EmitsFootnotePart()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var footnotesPart = main.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes = new Footnotes(
                new Footnote(
                    new Paragraph(
                        new Run(new FootnoteReferenceMark()),
                        new Run(new Text(" Footnote text"))))
                {
                    Id = 1
                });

            main.Document = new Document(
                new Body(
                    new Paragraph(
                        new Run(new Text("Body text")),
                        new Run(new FootnoteReference { Id = 1 })),
                    new SectionProperties(
                        new PageSize { Width = 12240U, Height = 15840U },
                        new PageMargin { Top = 1440, Right = 1440U, Bottom = 1440, Left = 1440U, Header = 720U, Footer = 720U, Gutter = 0U })));

            main.Document.Save();
            footnotesPart.Footnotes.Save();
        }

        stream.Position = 0;

        using var readDoc = WordprocessingDocument.Open(stream, false);
        var config = DxpHtmlVisitorConfig.CreateRichConfig();
        config.EmitParagraphMetadata = true;
        var visitor = new DxpHtmlVisitor(config, Logger);
        var html = TestCompare.Normalize(DxpExport.ExportToString(
            readDoc,
            visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger));

        Assert.Contains("data-docx-part=\"document\"", html, StringComparison.Ordinal);
        Assert.Contains("data-docx-part=\"footnote\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_DoesNotEmitParagraphMetadata_ByDefault()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var paragraph = new Paragraph(
                new Run(new Text("Body text")));
            paragraph.ParagraphId = "1234ABCD";

            main.Document = new Document(new Body(paragraph));
            main.Document.Save();
        }

        stream.Position = 0;

        using var readDoc = WordprocessingDocument.Open(stream, false);
        var visitor = new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger);
        var html = TestCompare.Normalize(DxpExport.ExportToString(
            readDoc,
            visitor,
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger));

        Assert.DoesNotContain("data-para-id=", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data-docx-part=", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_MarkupClassifier_AcceptRejectInlineAndSplit()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r>
      <w:rPr><w:u w:val="single"/></w:rPr>
      <w:t>Inserted</w:t>
    </w:r>
    <w:r>
      <w:rPr><w:strike/></w:rPr>
      <w:t>Deleted</w:t>
    </w:r>
  </w:p>
</w:body>
""";

        var acceptConfig = DxpHtmlVisitorConfig.CreateRichConfig();
        acceptConfig.TrackedChangeMode = DxpTrackedChangeMode.AcceptChanges;
        acceptConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string accepted = ExportHtmlFromBodyXml(bodyXml, acceptConfig);

        Assert.Contains("Inserted", accepted, StringComparison.Ordinal);
        Assert.DoesNotContain("Deleted", accepted, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-underline\"", accepted, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-strike\"", accepted, StringComparison.Ordinal);

        var rejectConfig = DxpHtmlVisitorConfig.CreateRichConfig();
        rejectConfig.TrackedChangeMode = DxpTrackedChangeMode.RejectChanges;
        rejectConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string rejected = ExportHtmlFromBodyXml(bodyXml, rejectConfig);

        Assert.DoesNotContain("Inserted", rejected, StringComparison.Ordinal);
        Assert.Contains("Deleted", rejected, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-underline\"", rejected, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-strike\"", rejected, StringComparison.Ordinal);

        var inlineConfig = DxpHtmlVisitorConfig.CreateRichConfig();
        inlineConfig.TrackedChangeMode = DxpTrackedChangeMode.InlineChanges;
        inlineConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string inline = ExportHtmlFromBodyXml(bodyXml, inlineConfig);

        Assert.Contains("dxp-inserted", inline, StringComparison.Ordinal);
        Assert.Contains("dxp-deleted", inline, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-underline\"", inline, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-strike\"", inline, StringComparison.Ordinal);

        var splitConfig = DxpHtmlVisitorConfig.CreateRichConfig();
        splitConfig.TrackedChangeMode = DxpTrackedChangeMode.SplitChanges;
        splitConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string split = ExportHtmlFromBodyXml(bodyXml, splitConfig);

        Assert.Contains("dxp-split", split, StringComparison.Ordinal);
        Assert.Contains("Inserted", split, StringComparison.Ordinal);
        Assert.Contains("Deleted", split, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-underline\"", split, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-strike\"", split, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_MarkupClassifier_DoubleStrikeRejectsAsDeleted()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r>
      <w:rPr><w:dstrike/></w:rPr>
      <w:t>DoubleDeleted</w:t>
    </w:r>
  </w:p>
</w:body>
""";

        var acceptConfig = DxpHtmlVisitorConfig.CreateRichConfig();
        acceptConfig.TrackedChangeMode = DxpTrackedChangeMode.AcceptChanges;
        acceptConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string accepted = ExportHtmlFromBodyXml(bodyXml, acceptConfig);
        Assert.DoesNotContain("DoubleDeleted", accepted, StringComparison.Ordinal);

        var rejectConfig = DxpHtmlVisitorConfig.CreateRichConfig();
        rejectConfig.TrackedChangeMode = DxpTrackedChangeMode.RejectChanges;
        rejectConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string rejected = ExportHtmlFromBodyXml(bodyXml, rejectConfig);
        Assert.Contains("DoubleDeleted", rejected, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-strike\"", rejected, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_MarkupClassifier_RealTrackedChangesWin()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:ins w:id="1" w:author="Tester" w:date="2024-01-01T00:00:00Z">
      <w:r>
        <w:rPr><w:strike/></w:rPr>
        <w:t>TrackedInsert</w:t>
      </w:r>
    </w:ins>
  </w:p>
</w:body>
""";

        var rejectConfig = DxpHtmlVisitorConfig.CreateRichConfig();
        rejectConfig.TrackedChangeMode = DxpTrackedChangeMode.RejectChanges;
        rejectConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string rejected = ExportHtmlFromBodyXml(bodyXml, rejectConfig);

        Assert.DoesNotContain("TrackedInsert", rejected, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlExport_MarkupClassifier_PreservesOtherStyles()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r>
      <w:rPr><w:b/><w:u w:val="single"/></w:rPr>
      <w:t>BoldInsert</w:t>
    </w:r>
  </w:p>
</w:body>
""";

        var config = DxpHtmlVisitorConfig.CreateRichConfig();
        config.TrackedChangeMode = DxpTrackedChangeMode.AcceptChanges;
        config.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string html = ExportHtmlFromBodyXml(bodyXml, config);

        Assert.Contains("dxp-bold", html, StringComparison.Ordinal);
        Assert.Contains("BoldInsert", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"dxp-underline\"", html, StringComparison.Ordinal);
    }

    private void VerifyAgainstFixture(
        Sample sample,
        DxpHtmlVisitorConfig baseConfig,
        string expectedExt,
        string actualSuffix,
        DxpTrackedChangeMode mode,
        DxpFieldEvalExportMode evalMode)
    {
        string expectedPath = TestPaths.GetSampleOutputPath(sample.DocxPath, expectedExt);
        string actualPath = TestPaths.GetSampleOutputPath(sample.DocxPath, actualSuffix);

        var config = CloneConfig(baseConfig, mode);
        string html = TestCompare.Normalize(ToHtml(sample.DocxPath, config, evalMode));
        File.WriteAllText(actualPath, html);

        if (!File.Exists(expectedPath))
            throw new XunitException($"Expected HTML file missing for {sample.FileName} ({mode}). Add {expectedPath}. Actual output saved to {actualPath}.");

        string expectedHtml = TestCompare.Normalize(File.ReadAllText(expectedPath));

        if (!string.Equals(expectedHtml, html, StringComparison.Ordinal))
        {
            string diff = TestCompare.DescribeDifference(expectedHtml, html);
            throw new XunitException($"Mismatch for {sample.FileName} ({mode}): {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
        }
    }

    private string ToHtml(string docxPath, DxpHtmlVisitorConfig config, DxpFieldEvalExportMode evalMode)
    {
        DxpFieldEval? fieldEval = null;
        if (evalMode == DxpFieldEvalExportMode.Evaluate)
            fieldEval = CreateEvalWithAsk();

        var visitor = new DxpHtmlVisitor(config, Logger, fieldEval);
        var options = new DxpExportOptions { FieldEvalMode = evalMode };
        return DxpExport.ExportToString(docxPath, visitor, options, Logger);
    }

    private DxpHtmlVisitorConfig CloneConfig(DxpHtmlVisitorConfig source, DxpTrackedChangeMode mode)
    {
        return new DxpHtmlVisitorConfig {
            EmitImages = source.EmitImages,
            EmitParagraphMetadata = source.EmitParagraphMetadata,
            EmitStyleFont = source.EmitStyleFont,
            EmitRunColor = source.EmitRunColor,
            EmitRunBackground = source.EmitRunBackground,
            EmitTableBorders = source.EmitTableBorders,
            EmitDocumentColors = source.EmitDocumentColors,
            EmitParagraphAlignment = source.EmitParagraphAlignment,
            PreserveListSymbols = source.PreserveListSymbols,
            RichTables = source.RichTables,
            EmitSectionHeadersFooters = source.EmitSectionHeadersFooters,
            EmitUnreferencedBookmarks = source.EmitUnreferencedBookmarks,
            EmitPageNumbers = source.EmitPageNumbers,
            UsePlainComments = source.UsePlainComments,
            EmitCustomProperties = source.EmitCustomProperties,
            EmitTimeline = source.EmitTimeline,
            StylesheetHref = source.StylesheetHref,
            EmbedDefaultStylesheet = source.EmbedDefaultStylesheet,
            RootCssClass = source.RootCssClass,
            TrackedChangeMode = mode,
            MarkupChangeClassifier = source.MarkupChangeClassifier
        };
    }

    private static void ConfigureEvalContext(DxpFieldEval eval)
    {
        eval.Context.SetDocVariable("Var1", "two");
        eval.Context.SetMergeFieldAlias("GivenName", "FirstName");
        eval.Context.ValueResolver = new DxpChainedFieldValueResolver(
            new SampleFieldValueResolver(),
            new DxpContextFieldValueResolver());
    }

    private DxpFieldEval CreateEvalWithAsk()
    {
        var delegates = new DxpFieldEvalDelegates {
            AskAsync = (request, _) => Task.FromResult<DxpFieldValue?>(request.PromptText switch {
                "Name?" => new DxpFieldValue("Bob"),
                "Hi Bob?" => new DxpFieldValue("Montreal"),
                _ => null
            })
        };

        var eval = new DxpFieldEval(delegates, logger: Logger);
        ConfigureEvalContext(eval);
        return eval;
    }

    private sealed class SampleFieldValueResolver : IDxpFieldValueResolver
    {
        public Task<DxpFieldValue?> ResolveAsync(string name, DxpFieldValueKindHint kind, DxpFieldEvalContext context)
        {
            _ = context;
            if (kind == DxpFieldValueKindHint.Any || kind == DxpFieldValueKindHint.MergeField)
            {
                if (string.Equals(name, "FirstName", StringComparison.OrdinalIgnoreCase))
                    return Task.FromResult<DxpFieldValue?>(new DxpFieldValue("Ana"));
                if (string.Equals(name, "EmptyField", StringComparison.OrdinalIgnoreCase))
                    return Task.FromResult<DxpFieldValue?>(new DxpFieldValue(string.Empty));
            }
            return Task.FromResult<DxpFieldValue?>(null);
        }
    }

    private void VerifyCachedAgainstFixture(Sample sample, DxpHtmlVisitorConfig baseConfig, string expectedExt, string actualSuffix)
    {
        string expectedPath = TestPaths.GetSampleOutputPath(sample.DocxPath, expectedExt);
        string actualPath = TestPaths.GetSampleOutputPath(sample.DocxPath, actualSuffix);

        var config = CloneConfig(baseConfig, DxpTrackedChangeMode.AcceptChanges);
        string html = TestCompare.Normalize(ToHtmlCached(sample.DocxPath, config));
        File.WriteAllText(actualPath, html);

        if (!File.Exists(expectedPath))
            throw new XunitException($"Expected HTML file missing for {sample.FileName} (CachedFields). Add {expectedPath}. Actual output saved to {actualPath}.");

        string expectedHtml = TestCompare.Normalize(File.ReadAllText(expectedPath));
        if (!string.Equals(expectedHtml, html, StringComparison.Ordinal))
        {
            string diff = TestCompare.DescribeDifference(expectedHtml, html);
            throw new XunitException($"Mismatch for {sample.FileName} (CachedFields): {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
        }
    }

    private string ToHtmlCached(string docxPath, DxpHtmlVisitorConfig config)
    {
        var visitor = new DxpHtmlVisitor(config, Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        if (visitor is not Fields.DxpIFieldEvalProvider provider)
            throw new XunitException("DxpHtmlVisitor should provide field evaluation context.");

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => DxpFieldEvalMiddleware.CreateCachedFieldMiddleware(next, provider.FieldEval, logger: Logger),
            next => new DxpContextMiddleware(next));

        new DxpWalker(Logger).Accept(docxPath, pipeline);
        return writer.ToString();
    }

    private IReadOnlyList<DxpComputedTabStop> CaptureParagraphTabStops(string bodyXml)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var xml = XDocument.Parse(bodyXml);
            var body = new Body();
            body.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            body.InnerXml = string.Concat(xml.Root!.Nodes());
            main.Document = new Document(body);
            main.Document.Save();
        }

        stream.Position = 0;

        using var readDoc = WordprocessingDocument.Open(stream, false);
        var visitor = new TabStopCaptureVisitor();
        new DxpWalker(Logger).Accept(readDoc, new DxpContextMiddleware(visitor));
        return visitor.Captured;
    }

    private string ExportHtmlFromBodyXml(
        string bodyXml,
        DxpHtmlVisitorConfig config,
        Action<WordprocessingDocument>? configureDocument = null)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var xml = XDocument.Parse(bodyXml);
            var body = new Body();
            body.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            body.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            body.InnerXml = string.Concat(xml.Root!.Nodes());
            main.Document = new Document(body);
            main.Document.Save();
            configureDocument?.Invoke(doc);
        }

        stream.Position = 0;

        using var readDoc = WordprocessingDocument.Open(stream, false);
        return TestCompare.Normalize(DxpExport.ExportToString(
            readDoc,
            new DxpHtmlVisitor(config, Logger),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger));
    }

    private string ExportHeaderFooterLayout(
        int top,
        int bottom,
        uint headerDistance,
        uint footerDistance,
        Header? header = null,
        Footer? footer = null)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var section = new SectionProperties();

            if (header != null)
            {
                var part = main.AddNewPart<HeaderPart>("rIdHeaderLayout");
                part.Header = header;
                if (header.Descendants<Drawing>().Any())
                {
                    var imagePart = part.AddNewPart<ImagePart>("image/png", "rIdImage1");
                    using var image = new MemoryStream(Convert.FromBase64String(
                        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
                    imagePart.FeedData(image);
                }
                section.Append(new HeaderReference { Id = main.GetIdOfPart(part), Type = HeaderFooterValues.Default });
                part.Header.Save();
            }

            if (footer != null)
            {
                var part = main.AddNewPart<FooterPart>("rIdFooterLayout");
                part.Footer = footer;
                if (footer.Descendants<Drawing>().Any())
                {
                    var imagePart = part.AddNewPart<ImagePart>("image/png", "rIdImage1");
                    using var image = new MemoryStream(Convert.FromBase64String(
                        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
                    imagePart.FeedData(image);
                }
                section.Append(new FooterReference { Id = main.GetIdOfPart(part), Type = HeaderFooterValues.Default });
                part.Footer.Save();
            }

            section.Append(
                new PageSize { Width = 12240U, Height = 15840U },
                new PageMargin {
                    Top = top,
                    Right = 1440U,
                    Bottom = bottom,
                    Left = 1440U,
                    Header = headerDistance,
                    Footer = footerDistance,
                    Gutter = 0U
                });

            main.Document = new Document(
                new Body(
                    new Paragraph(new Run(new Text("Body text"))),
                    section));
            main.Document.Save();
        }

        stream.Position = 0;
        using var readDoc = WordprocessingDocument.Open(stream, false);
        return TestCompare.Normalize(DxpExport.ExportToString(
            readDoc,
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig(), Logger),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger));
    }

    private static Header CreateHeaderWithInlineImage(long widthEmu, long heightEmu)
        => new(new Paragraph(new Run(CreateInlineImageDrawing(widthEmu, heightEmu))));

    private static Drawing CreateInlineImageDrawing(
        long widthEmu,
        long heightEmu,
        A.SourceRectangle? sourceRectangle = null,
        int? rotation = null,
        bool flipHorizontal = false,
        bool flipVertical = false,
        string? description = null,
        string? title = null)
    {
        var transform = new A.Transform2D(
            new A.Offset { X = 0L, Y = 0L },
            new A.Extents { Cx = widthEmu, Cy = heightEmu }) {
            Rotation = rotation,
            HorizontalFlip = flipHorizontal,
            VerticalFlip = flipVertical
        };
        var picture = new PIC.Picture(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties { Id = 1U, Name = "Header image", Description = description, Title = title },
                new PIC.NonVisualPictureDrawingProperties()),
            new PIC.BlipFill(
                new A.Blip { Embed = "rIdImage1" },
                sourceRectangle ?? new A.SourceRectangle(),
                new A.Stretch(new A.FillRectangle())),
            new PIC.ShapeProperties(
                transform,
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        var inline = new WP.Inline(
            new WP.Extent { Cx = widthEmu, Cy = heightEmu },
            new WP.DocProperties { Id = 1U, Name = "Header image", Description = description, Title = title },
            new WP.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(new A.GraphicData(picture) {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
            })) {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U
        };

        return new Drawing(inline);
    }

    private string ExportHtmlCachedFromBodyXml(
        string bodyXml,
        DxpHtmlVisitorConfig config,
        Action<WordprocessingDocument>? configureDocument = null)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var xml = XDocument.Parse(bodyXml);
            var body = new Body();
            body.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            body.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            body.InnerXml = string.Concat(xml.Root!.Nodes());
            main.Document = new Document(body);
            main.Document.Save();
            configureDocument?.Invoke(doc);
        }

        stream.Position = 0;

        var visitor = new DxpHtmlVisitor(config, Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        if (visitor is not Fields.DxpIFieldEvalProvider provider)
            throw new XunitException("DxpHtmlVisitor should provide field evaluation context.");

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => DxpFieldEvalMiddleware.CreateCachedFieldMiddleware(next, provider.FieldEval, logger: Logger),
            next => new DxpContextMiddleware(next));

        using (var readDoc = WordprocessingDocument.Open(stream, false))
            new DxpWalker(Logger).Accept(readDoc, pipeline);

        return writer.ToString();
    }

    private sealed class TabStopCaptureVisitor : DxpVisitor
    {
        public List<DxpComputedTabStop> Captured { get; } = new();

        public TabStopCaptureVisitor() : base(null)
        {
        }

        public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
        {
            if (paragraph.Layout?.TabStops.Count > 0)
                Captured.Add(paragraph.Layout.TabStops[0]);
            return DxpDisposable.Empty;
        }
    }
}
