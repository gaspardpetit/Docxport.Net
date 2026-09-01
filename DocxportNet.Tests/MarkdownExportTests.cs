using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Resolution;
using DocxportNet.Middleware;
using DocxportNet.Tests.Utils;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Omml;
using System.Xml.Linq;
using Xunit.Abstractions;
using Xunit.Sdk;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxportNet.Tests;

public class MarkdownExportTests : TestBase<MarkdownExportTests>
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
        public string ExpectedMarkdownPath => Path.ChangeExtension(DocxPath, ".md");
        public string TestOutputPath => TestPaths.GetSampleOutputPath(DocxPath, ".test.md");

        public void Serialize(IXunitSerializationInfo info) => info.AddValue(nameof(DocxPath), DocxPath);
        public void Deserialize(IXunitSerializationInfo info) => DocxPath = info.GetValue<string>(nameof(DocxPath));

        public override string ToString() => FileName; // keep theory display concise
    }

    private static readonly string ProjectRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
    private static readonly string SamplesDirectory = Path.Combine(ProjectRoot, "samples");

    public MarkdownExportTests(ITestOutputHelper output) : base(output)
    {
    }

    public static IEnumerable<object[]> SampleDocs()
    {
        return Directory.EnumerateFiles(SamplesDirectory, "*.docx", SearchOption.TopDirectoryOnly)
            .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
            .OrderBy(Path.GetFileName) // deterministic ordering for discovery
            .Select(path => new object[] { new Sample(path) });
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Rich(Sample sample)
    {
        VerifyAgainstFixture(sample, DxpMarkdownVisitorConfig.CreateRichConfig(), ".md", ".test.md", DxpFieldEvalExportMode.None);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Plain(Sample sample)
    {
        VerifyAgainstFixture(sample, DxpMarkdownVisitorConfig.CreatePlainConfig(), ".plain.md", ".plain.test.md", DxpFieldEvalExportMode.None);
    }

    [Fact]
    public void MarkdownExport_RendersSharedCropRotationFlipAndDimensions()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p /></w:body>
""";
        string markdown = ExportMarkdownFromBodyXml(
            bodyXml,
            DxpMarkdownVisitorConfig.CreateRichConfig(),
            doc => {
                var main = doc.MainDocumentPart!;
                var imagePart = main.AddImagePart("image/png", "rIdImage1");
                using var image = new MemoryStream(Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
                imagePart.FeedData(image);
                main.Document!.Body = new Body(new Paragraph(new Run(CreatePresentedImageDrawing())));
                main.Document.Save();
            });

        Assert.Contains("class=\"dxp-image-frame\"", markdown, StringComparison.Ordinal);
        Assert.Contains("width:100pt;height:50pt;overflow:hidden;transform:rotate(90deg) scaleX(-1);", markdown, StringComparison.Ordinal);
        Assert.Contains("alt=\"Markdown image\" title=\"Markdown title\"", markdown, StringComparison.Ordinal);
        Assert.Contains("width:166.667pt;height:71.429pt;left:-16.667pt;top:-14.286pt;", markdown, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Rich_Reject(Sample sample)
    {
        VerifyVariant(sample, DxpMarkdownVisitorConfig.CreateRichConfig(), ".reject.test.md", DxpTrackedChangeMode.RejectChanges, DxpFieldEvalExportMode.None);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Rich_Inline(Sample sample)
    {
        VerifyVariant(sample, DxpMarkdownVisitorConfig.CreateRichConfig(), ".inline.test.md", DxpTrackedChangeMode.InlineChanges, DxpFieldEvalExportMode.None);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Rich_Split(Sample sample)
    {
        VerifyVariant(sample, DxpMarkdownVisitorConfig.CreateRichConfig(), ".split.test.md", DxpTrackedChangeMode.SplitChanges, DxpFieldEvalExportMode.None);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Plain_Reject(Sample sample)
    {
        VerifyVariant(sample, DxpMarkdownVisitorConfig.CreatePlainConfig(), ".plain.reject.test.md", DxpTrackedChangeMode.RejectChanges, DxpFieldEvalExportMode.None);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Plain_Split(Sample sample)
    {
        VerifyVariant(sample, DxpMarkdownVisitorConfig.CreatePlainConfig(), ".plain.split.test.md", DxpTrackedChangeMode.SplitChanges, DxpFieldEvalExportMode.None);
    }


    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Plain_Inline(Sample sample)
    {
        VerifyVariant(sample, DxpMarkdownVisitorConfig.CreatePlainConfig(), ".plain.inline.test.md", DxpTrackedChangeMode.InlineChanges, DxpFieldEvalExportMode.None);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Plain_Cached(Sample sample)
    {
        VerifyCachedAgainstFixture(sample, DxpMarkdownVisitorConfig.CreatePlainConfig(), ".plain.cached.md", ".plain.cached.test.md");
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Rich_Eval(Sample sample)
    {
        VerifyAgainstFixture(sample, DxpMarkdownVisitorConfig.CreateRichConfig(), ".eval.md", ".eval.test.md", DxpFieldEvalExportMode.Evaluate);
    }

    [Theory]
    [MemberData(nameof(SampleDocs))]
    public void TestDocxToMarkdown_Plain_Eval(Sample sample)
    {
        VerifyAgainstFixture(sample, DxpMarkdownVisitorConfig.CreatePlainConfig(), ".plain.eval.md", ".plain.eval.test.md", DxpFieldEvalExportMode.Evaluate);
    }

    [Fact]
    public void MarkdownExport_NonBreakingSpaceRetainsParagraphSeparation()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Before</w:t></w:r></w:p>
  <w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">&#160;</w:t></w:r></w:p>
  <w:p><w:r><w:t>After</w:t></w:r></w:p>
</w:body>
""";

        string markdown = ExportMarkdownFromBodyXml(bodyXml, DxpMarkdownVisitorConfig.CreateRichConfig());

        Assert.Matches("<span[^>]*>\u00A0</span>\\r?\\n\\r?\\n", markdown);
    }

    [Fact]
    public void MarkdownExport_RendersInlineAndDisplayMathAsUnicodeMath()
    {
        const string bodyXml = """
            <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
                <w:p>
                  <w:r><w:t xml:space="preserve">Before </w:t></w:r>
                  <m:oMath><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>2</m:t></m:r></m:sub></m:sSub></m:oMath>
                  <w:r><w:t xml:space="preserve"> after</w:t></w:r>
                </w:p>
                <w:p><m:oMathPara><m:oMath><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara></w:p>
            </w:body>
            """;

        DxpMarkdownVisitorConfig config = DxpMarkdownVisitorConfig.CreatePlainConfig() with
        {
            MathOutputFormat = DxpOmmlOutputFormat.UnicodeMath,
        };
        string markdown = ExportMarkdownFromBodyXml(bodyXml, config);

        Assert.Contains("Before $x_(2)$ after", markdown, StringComparison.Ordinal);
        Assert.Contains("$$\n(1)/(2)\n$$", markdown, StringComparison.Ordinal);

        string defaultMarkdown = ExportMarkdownFromBodyXml(bodyXml, DxpMarkdownVisitorConfig.CreatePlainConfig());
        Assert.Contains("Before $x_{2}$ after", defaultMarkdown, StringComparison.Ordinal);
        Assert.Contains("$$\n\\frac{1}{2}\n$$", defaultMarkdown, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_CanEmitRawMathMlOrDisableMath()
    {
        const string bodyXml = """
            <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
              <w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>
            </w:body>
            """;

        DxpMarkdownVisitorConfig mathMl = DxpMarkdownVisitorConfig.CreatePlainConfig() with
        {
            MathOutputFormat = DxpOmmlOutputFormat.MathMl,
        };
        DxpMarkdownVisitorConfig omitted = mathMl with { MathOutputFormat = null };

        Assert.Contains("<math", ExportMarkdownFromBodyXml(bodyXml, mathMl), StringComparison.Ordinal);
        Assert.DoesNotContain("x", ExportMarkdownFromBodyXml(bodyXml, omitted), StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_UsesWalkerResolverForEmbeddedWordprocessingMl()
    {
        const string bodyXml = """
            <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
              <w:p><m:oMath><w:hyperlink><w:r><w:t>A_B</w:t></w:r></w:hyperlink></m:oMath></w:p>
            </w:body>
            """;

        string markdown = ExportMarkdownFromBodyXml(bodyXml, DxpMarkdownVisitorConfig.CreatePlainConfig());

        Assert.Contains("$A\\_B$", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_AppliesTrackedChangePolicyInsideMath()
    {
        const string bodyXml = """
            <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
              <w:p><m:oMath><w:customXml>
                <w:ins><w:r><w:t>new</w:t></w:r></w:ins>
                <w:del><w:r><w:delText>old</w:delText></w:r></w:del>
              </w:customXml></m:oMath></w:p>
            </w:body>
            """;
        DxpMarkdownVisitorConfig accepted = DxpMarkdownVisitorConfig.CreatePlainConfig() with
            { TrackedChangeMode = DxpTrackedChangeMode.AcceptChanges };
        DxpMarkdownVisitorConfig rejected = accepted with { TrackedChangeMode = DxpTrackedChangeMode.RejectChanges };

        Assert.Contains("$new$", ExportMarkdownFromBodyXml(bodyXml, accepted), StringComparison.Ordinal);
        Assert.Contains("$old$", ExportMarkdownFromBodyXml(bodyXml, rejected), StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_EmitRichLayoutHtml_RichModeEmitsParagraphWrapper()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:ind w:left="720"/>
      <w:spacing w:line="230" w:lineRule="auto"/>
    </w:pPr>
    <w:r><w:t>Styled paragraph</w:t></w:r>
  </w:p>
</w:body>
""";

        var config = DxpMarkdownVisitorConfig.CreateRichConfig();
        string markdown = ExportMarkdownFromBodyXml(bodyXml, config);

        Assert.Contains("<p style=", markdown, StringComparison.Ordinal);
        Assert.Contains("Styled paragraph", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_EmitRichLayoutHtml_PlainModeSuppressesParagraphWrapper()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:ind w:left="720"/>
      <w:spacing w:line="230" w:lineRule="auto"/>
    </w:pPr>
    <w:r><w:t>Styled paragraph</w:t></w:r>
  </w:p>
</w:body>
""";

        var config = DxpMarkdownVisitorConfig.CreatePlainConfig();
        string markdown = ExportMarkdownFromBodyXml(bodyXml, config);

        Assert.DoesNotContain("<p style=", markdown, StringComparison.Ordinal);
        Assert.Contains("Styled paragraph", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_EmitRichLayoutHtml_RichModeEmitsTabSpan()
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

        var config = DxpMarkdownVisitorConfig.CreateRichConfig();
        string markdown = ExportMarkdownFromBodyXml(bodyXml, config);

        Assert.Contains("<span class=\"dxp-tab", markdown, StringComparison.Ordinal);
        Assert.Contains("Left", markdown, StringComparison.Ordinal);
        Assert.Contains("Right", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_EmitRichLayoutHtml_PlainModeSuppressesTabSpan()
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

        var config = DxpMarkdownVisitorConfig.CreatePlainConfig();
        string markdown = ExportMarkdownFromBodyXml(bodyXml, config);

        Assert.DoesNotContain("<span class=\"dxp-tab", markdown, StringComparison.Ordinal);
        Assert.Contains("Left\tRight", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_EmitRichLayoutHtml_DoesNotSuppressHeaderHtmlGuardedElsewhere()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t>Body text</w:t></w:r>
  </w:p>
  <w:sectPr>
    <w:headerReference w:type="default" r:id="rIdHeader1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
  </w:sectPr>
</w:body>
""";

        var config = DxpMarkdownVisitorConfig.CreatePlainConfig();
        string markdown = ExportMarkdownFromBodyXml(
            bodyXml,
            config,
            document => {
                var main = document.MainDocumentPart ?? throw new InvalidOperationException("Main document part should exist.");
                var headerPart = main.AddNewPart<HeaderPart>("rIdHeader1");
                headerPart.Header = new Header(
                    new Paragraph(
                        new Run(new Text("Header text"))));
                headerPart.Header.Save();
            });

        Assert.Contains("<div class=\"header\"", markdown, StringComparison.Ordinal);
        Assert.Contains("Header text", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_CacheMode_DefaultFieldFallback_ReplaysCachedResults_AndSuppressesOnlyTruePageFields()
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

        var markdown = TestCompare.Normalize(ExportMarkdownCachedFromBodyXml(bodyXml, DxpMarkdownVisitorConfig.CreateRichConfig()));

        Assert.Contains("Unknown: cached unknown", markdown, StringComparison.Ordinal);
        Assert.Contains("TOC: Heading 1", markdown, StringComparison.Ordinal);
        Assert.Contains("3", markdown, StringComparison.Ordinal);
        Assert.Contains("Page:", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("Page: 4", markdown, StringComparison.Ordinal);
        Assert.Contains("PageRef: 7", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_MarkupClassifier_AcceptRejectInlineAndSplit()
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

        var acceptConfig = DxpMarkdownVisitorConfig.CreateRichConfig();
        acceptConfig.TrackedChangeMode = DxpTrackedChangeMode.AcceptChanges;
        acceptConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string accepted = ExportMarkdownFromBodyXml(bodyXml, acceptConfig);

        Assert.Contains("Inserted", accepted, StringComparison.Ordinal);
        Assert.DoesNotContain("Deleted", accepted, StringComparison.Ordinal);
        Assert.DoesNotContain("<u>", accepted, StringComparison.Ordinal);
        Assert.DoesNotContain("<del>", accepted, StringComparison.Ordinal);

        var rejectConfig = DxpMarkdownVisitorConfig.CreateRichConfig();
        rejectConfig.TrackedChangeMode = DxpTrackedChangeMode.RejectChanges;
        rejectConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string rejected = ExportMarkdownFromBodyXml(bodyXml, rejectConfig);

        Assert.DoesNotContain("Inserted", rejected, StringComparison.Ordinal);
        Assert.Contains("Deleted", rejected, StringComparison.Ordinal);
        Assert.DoesNotContain("<u>", rejected, StringComparison.Ordinal);
        Assert.DoesNotContain("<del>", rejected, StringComparison.Ordinal);

        var inlineConfig = DxpMarkdownVisitorConfig.CreateRichConfig();
        inlineConfig.TrackedChangeMode = DxpTrackedChangeMode.InlineChanges;
        inlineConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string inline = ExportMarkdownFromBodyXml(bodyXml, inlineConfig);

        Assert.Contains("Inserted", inline, StringComparison.Ordinal);
        Assert.Contains("Deleted", inline, StringComparison.Ordinal);
        Assert.Contains("color:blue", inline, StringComparison.Ordinal);
        Assert.Contains("color:red", inline, StringComparison.Ordinal);
        Assert.DoesNotContain("<u>Inserted</u>", inline, StringComparison.Ordinal);
        Assert.DoesNotContain("<del>Deleted</del>", inline, StringComparison.Ordinal);

        var splitConfig = DxpMarkdownVisitorConfig.CreateRichConfig();
        splitConfig.TrackedChangeMode = DxpTrackedChangeMode.SplitChanges;
        splitConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string split = ExportMarkdownFromBodyXml(bodyXml, splitConfig);

        Assert.Contains("<table", split, StringComparison.Ordinal);
        Assert.Contains("Inserted", split, StringComparison.Ordinal);
        Assert.Contains("Deleted", split, StringComparison.Ordinal);
        Assert.DoesNotContain("<u>", split, StringComparison.Ordinal);
        Assert.DoesNotContain("<del>", split, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_MarkupClassifier_DoubleStrikeRejectsAsDeleted()
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

        var acceptConfig = DxpMarkdownVisitorConfig.CreateRichConfig();
        acceptConfig.TrackedChangeMode = DxpTrackedChangeMode.AcceptChanges;
        acceptConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string accepted = ExportMarkdownFromBodyXml(bodyXml, acceptConfig);
        Assert.DoesNotContain("DoubleDeleted", accepted, StringComparison.Ordinal);

        var rejectConfig = DxpMarkdownVisitorConfig.CreateRichConfig();
        rejectConfig.TrackedChangeMode = DxpTrackedChangeMode.RejectChanges;
        rejectConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string rejected = ExportMarkdownFromBodyXml(bodyXml, rejectConfig);
        Assert.Contains("DoubleDeleted", rejected, StringComparison.Ordinal);
        Assert.DoesNotContain("<del>", rejected, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_MarkupClassifier_RealTrackedChangesWin()
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

        var rejectConfig = DxpMarkdownVisitorConfig.CreateRichConfig();
        rejectConfig.TrackedChangeMode = DxpTrackedChangeMode.RejectChanges;
        rejectConfig.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string rejected = ExportMarkdownFromBodyXml(bodyXml, rejectConfig);

        Assert.DoesNotContain("TrackedInsert", rejected, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_MarkupClassifier_PreservesOtherStyles()
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

        var config = DxpMarkdownVisitorConfig.CreateRichConfig();
        config.TrackedChangeMode = DxpTrackedChangeMode.AcceptChanges;
        config.MarkupChangeClassifier = DxpMarkupChangeClassifiers.UnderlineInsertedStrikeDeleted();
        string markdown = ExportMarkdownFromBodyXml(bodyXml, config);

        Assert.Contains("<b>BoldInsert</b>", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("<u>", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownExport_EmitRichLayoutHtml_DoesNotSuppressDeletedMarkup()
    {
        const string bodyXml = """
<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:del w:id="1" w:author="test">
      <w:r>
        <w:delText>Deleted text</w:delText>
      </w:r>
    </w:del>
  </w:p>
</w:body>
""";

        var config = DxpMarkdownVisitorConfig.CreatePlainConfig();
        config.TrackedChangeMode = DxpTrackedChangeMode.InlineChanges;
        string markdown = ExportMarkdownFromBodyXml(bodyXml, config);

        Assert.Contains("<del>", markdown, StringComparison.Ordinal);
        Assert.Contains("Deleted text", markdown, StringComparison.Ordinal);
    }

    private void VerifyAgainstFixture(
        Sample sample,
        DxpMarkdownVisitorConfig config,
        string expectedExt,
        string actualSuffix,
        DxpFieldEvalExportMode evalMode)
    {
        string expectedPath = TestPaths.GetSampleOutputPath(sample.DocxPath, expectedExt);
        string actualPath = TestPaths.GetSampleOutputPath(sample.DocxPath, actualSuffix);

        string actualMarkdown = TestCompare.Normalize(ToMarkdown(sample.DocxPath, CloneConfig(config, DxpTrackedChangeMode.AcceptChanges), evalMode));
        File.WriteAllText(actualPath, actualMarkdown);

        if (!File.Exists(expectedPath))
        {
            throw new XunitException($"Expected markdown file missing for {sample.FileName}. Add {expectedPath}. Actual output saved to {actualPath}.");
        }

        string expectedMarkdown = TestCompare.Normalize(File.ReadAllText(expectedPath));

        if (!string.Equals(expectedMarkdown, actualMarkdown, StringComparison.Ordinal))
        {
            string diff = TestCompare.DescribeDifference(expectedMarkdown, actualMarkdown);
            throw new XunitException($"Mismatch for {sample.FileName}: {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
        }

        // Emit additional tracked-change variants for inspection.
        WriteVariant(sample, config, DxpTrackedChangeMode.RejectChanges, actualSuffix.Replace(".test", ".reject.test"), evalMode);
        WriteVariant(sample, config, DxpTrackedChangeMode.InlineChanges, actualSuffix.Replace(".test", ".inline.test"), evalMode);
    }

    private void VerifyVariant(Sample sample, DxpMarkdownVisitorConfig config, string suffix, DxpTrackedChangeMode mode, DxpFieldEvalExportMode evalMode)
    {
        string expectedPath = TestPaths.GetSampleOutputPath(sample.DocxPath, suffix.Replace(".test", string.Empty));
        string actualPath = TestPaths.GetSampleOutputPath(sample.DocxPath, suffix);

        var cfg = CloneConfig(config, mode);
        string markdown = TestCompare.Normalize(ToMarkdown(sample.DocxPath, cfg, evalMode));
        File.WriteAllText(actualPath, markdown);

        if (!File.Exists(expectedPath))
            throw new XunitException($"Expected markdown file missing for {sample.FileName} ({mode}). Add {expectedPath}. Actual output saved to {actualPath}.");

        string expectedMarkdown = TestCompare.Normalize(File.ReadAllText(expectedPath));
        if (!string.Equals(expectedMarkdown, markdown, StringComparison.Ordinal))
        {
            string diff = TestCompare.DescribeDifference(expectedMarkdown, markdown);
            throw new XunitException($"Mismatch for {sample.FileName} ({mode}): {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
        }
    }

    private void WriteVariant(Sample sample, DxpMarkdownVisitorConfig baseConfig, DxpTrackedChangeMode mode, string suffix, DxpFieldEvalExportMode evalMode)
    {
        var cfg = CloneConfig(baseConfig, mode);
        string path = TestPaths.GetSampleOutputPath(sample.DocxPath, suffix);
        string markdown = TestCompare.Normalize(ToMarkdown(sample.DocxPath, cfg, evalMode));
        File.WriteAllText(path, markdown);
    }

    private string ToMarkdown(string docxPath, DxpMarkdownVisitorConfig config, DxpFieldEvalExportMode evalMode)
    {
        DxpFieldEval? fieldEval = null;
        if (evalMode == DxpFieldEvalExportMode.Evaluate)
            fieldEval = CreateEvalWithAsk();

        var visitor = new DxpMarkdownVisitor(config, Logger, fieldEval);
        var options = new DxpExportOptions { FieldEvalMode = evalMode };
        return DxpExport.ExportToString(docxPath, visitor, options, Logger);
    }

    private void VerifyCachedAgainstFixture(Sample sample, DxpMarkdownVisitorConfig config, string expectedExt, string actualSuffix)
    {
        string expectedPath = TestPaths.GetSampleOutputPath(sample.DocxPath, expectedExt);
        string actualPath = TestPaths.GetSampleOutputPath(sample.DocxPath, actualSuffix);

        string actualMarkdown = TestCompare.Normalize(ToMarkdownCached(sample.DocxPath, CloneConfig(config, DxpTrackedChangeMode.AcceptChanges)));
        File.WriteAllText(actualPath, actualMarkdown);

        if (!File.Exists(expectedPath))
            throw new XunitException($"Expected markdown file missing for {sample.FileName} (CachedFields). Add {expectedPath}. Actual output saved to {actualPath}.");

        string expectedMarkdown = TestCompare.Normalize(File.ReadAllText(expectedPath));
        if (!string.Equals(expectedMarkdown, actualMarkdown, StringComparison.Ordinal))
        {
            string diff = TestCompare.DescribeDifference(expectedMarkdown, actualMarkdown);
            throw new XunitException($"Mismatch for {sample.FileName} (CachedFields): {diff}. Expected: {expectedPath}. Actual: {actualPath}.");
        }
    }

    private string ToMarkdownCached(string docxPath, DxpMarkdownVisitorConfig config)
    {
        var visitor = new DxpMarkdownVisitor(config, Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        if (visitor is not Fields.DxpIFieldEvalProvider provider)
            throw new XunitException("DxpMarkdownVisitor should provide field evaluation context.");

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => Walker.DxpFieldEvalMiddleware.CreateCachedFieldMiddleware(next, provider.FieldEval, logger: Logger),
            next => new DxpContextMiddleware(next));

        new Walker.DxpWalker(Logger).Accept(docxPath, pipeline);
        return writer.ToString();
    }

    private string ExportMarkdownFromBodyXml(
        string bodyXml,
        DxpMarkdownVisitorConfig config,
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
            new DxpMarkdownVisitor(config, Logger),
            new DxpExportOptions { FieldEvalMode = DxpFieldEvalExportMode.None },
            Logger));
    }

    private static Drawing CreatePresentedImageDrawing()
    {
        var picture = new PIC.Picture(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties {
                    Id = 1U,
                    Name = "Markdown image",
                    Description = "Markdown image",
                    Title = "Markdown title"
                },
                new PIC.NonVisualPictureDrawingProperties()),
            new PIC.BlipFill(
                new A.Blip { Embed = "rIdImage1" },
                new A.SourceRectangle { Left = 10000, Top = 20000, Right = 30000, Bottom = 10000 },
                new A.Stretch(new A.FillRectangle())),
            new PIC.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 1270000L, Cy = 635000L }) {
                    Rotation = 5400000,
                    HorizontalFlip = true
                },
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        return new Drawing(new WP.Inline(
            new WP.Extent { Cx = 1270000L, Cy = 635000L },
            new WP.DocProperties {
                Id = 1U,
                Name = "Markdown image",
                Description = "Markdown image",
                Title = "Markdown title"
            },
            new A.Graphic(new A.GraphicData(picture) {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
            })));
    }

    private string ExportMarkdownCachedFromBodyXml(
        string bodyXml,
        DxpMarkdownVisitorConfig config,
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

        var visitor = new DxpMarkdownVisitor(config, Logger);
        using var writer = new StringWriter();
        visitor.SetOutput(writer);

        if (visitor is not Fields.DxpIFieldEvalProvider provider)
            throw new XunitException("DxpMarkdownVisitor should provide field evaluation context.");

        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => Walker.DxpFieldEvalMiddleware.CreateCachedFieldMiddleware(next, provider.FieldEval, logger: Logger),
            next => new DxpContextMiddleware(next));

        using (var readDoc = WordprocessingDocument.Open(stream, false))
            new Walker.DxpWalker(Logger).Accept(readDoc, pipeline);

        return writer.ToString();
    }

    private static DxpMarkdownVisitorConfig CloneConfig(DxpMarkdownVisitorConfig source, DxpTrackedChangeMode mode)
    {
        return new DxpMarkdownVisitorConfig {
            EmitImages = source.EmitImages,
            EmitStyleFont = source.EmitStyleFont,
            EmitRunColor = source.EmitRunColor,
            EmitRunBackground = source.EmitRunBackground,
            EmitTableBorders = source.EmitTableBorders,
            EmitDocumentColors = source.EmitDocumentColors,
            EmitParagraphAlignment = source.EmitParagraphAlignment,
            EmitRichLayoutHtml = source.EmitRichLayoutHtml,
            PreserveListSymbols = source.PreserveListSymbols,
            RichTables = source.RichTables,
            UsePlainCodeBlocks = source.UsePlainCodeBlocks,
            UseMarkdownInlineStyles = source.UseMarkdownInlineStyles,
            EmitSectionHeadersFooters = source.EmitSectionHeadersFooters,
            EmitUnreferencedBookmarks = source.EmitUnreferencedBookmarks,
            EmitPageNumbers = source.EmitPageNumbers,
            UsePlainComments = source.UsePlainComments,
            EmitCustomProperties = source.EmitCustomProperties,
            EmitTimeline = source.EmitTimeline,
            MathOutputFormat = source.MathOutputFormat,
            EmitMathDelimiters = source.EmitMathDelimiters,
            MathEmbeddedContentResolver = source.MathEmbeddedContentResolver,
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
}
