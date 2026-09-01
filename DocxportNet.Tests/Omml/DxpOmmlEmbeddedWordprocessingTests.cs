using System.Xml.Linq;
using DocxportNet.Omml;

namespace DocxportNet.Tests.Omml;

public sealed class DxpOmmlEmbeddedWordprocessingTests
{
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Theory]
    [InlineData("002.omml", @"\frac{1}{2}")]
    [InlineData("005.omml", "1^{2}")]
    public void PinnedPlurimathControlPropertyFixturesRemainOracleCompatible(string fixture, string expectedLatex)
    {
        string omml = File.ReadAllText(Path.Combine(OmmlTestData.UpstreamRoot, fixture));
        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Latex);

        Assert.Equal(expectedLatex, result.Output);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "OMML001");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.ElementName == "m:ctrlPr");
    }

    [Fact]
    public void MathRunPreservesWordContentAndPresentation()
    {
        string omml = $"""
            <m:oMath xmlns:m="{M}" xmlns:w="{W}"><m:r>
              <w:rPr><w:b/><w:i/><w:color w:val="12abEF"/><w:sz w:val="24"/>
                <w:rFonts w:ascii="Aptos"/><w:vertAlign w:val="superscript"/><w:lang w:val="fr-CA"/></w:rPr>
              <w:t>A</w:t><w:tab/><w:noBreakHyphen/><w:softHyphen/><w:sym w:font="Symbol" w:char="F061"/>
            </m:r></m:oMath>
            """;

        DxpOmmlConversionResult mathml = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.MathMl);
        XNamespace math = "http://www.w3.org/1998/Math/MathML";
        XElement style = XElement.Parse(mathml.Output).Descendants(math + "mstyle").Single();
        Assert.Equal("bold-italic", style.Attribute("mathvariant")?.Value);
        Assert.Equal("#12ABEF", style.Attribute("mathcolor")?.Value);
        Assert.Equal("12pt", style.Attribute("mathsize")?.Value);
        Assert.Equal("fr-CA", style.Attribute(XNamespace.Xml + "lang")?.Value);
        Assert.Contains("font-family:'Aptos'", style.Attribute("style")?.Value, StringComparison.Ordinal);
        Assert.Contains("vertical-align:super", style.Attribute("style")?.Value, StringComparison.Ordinal);
        Assert.Contains("A-\u00ADα", style.Value, StringComparison.Ordinal);
        Assert.Single(style.Descendants(math + "mspace"),
            space => space.Attribute("data-omml-tab")?.Value == "true");

        DxpOmmlConversionResult latex = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Latex);
        Assert.Contains(@"\textcolor[HTML]{12ABEF}", latex.Output, StringComparison.Ordinal);
        Assert.Contains(@"\fontsize{12pt}", latex.Output, StringComparison.Ordinal);
        Assert.Contains(@"{}^{\scriptstyle", latex.Output, StringComparison.Ordinal);
        Assert.Contains(@"\boldsymbol{", latex.Output, StringComparison.Ordinal);
        Assert.Contains(latex.Diagnostics, diagnostic => diagnostic.ElementName == "w:rFonts");
    }

    [Fact]
    public void ControlPropertiesAreParsedAndAppliedWithExplicitApproximation()
    {
        string omml = $"""
            <m:oMath xmlns:m="{M}" xmlns:w="{W}"><m:f><m:fPr><m:ctrlPr><w:rPr>
              <w:b/><w:color w:val="FF0000"/><w:sz w:val="20"/><w:rFonts w:ascii="Cambria Math"/>
            </w:rPr></m:ctrlPr></m:fPr>
              <m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den>
            </m:f></m:oMath>
            """;

        DxpOmmlConversionResult mathml = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.MathMl);
        XNamespace math = "http://www.w3.org/1998/Math/MathML";
        XElement style = XElement.Parse(mathml.Output).Descendants(math + "mstyle")
            .Single(element => element.Attribute("data-omml-control-properties")?.Value == "true");
        Assert.Equal("true", style.Attribute("data-omml-control-bold")?.Value);
        Assert.Equal("FF0000", style.Attribute("data-omml-control-color")?.Value);
        Assert.Equal("10", style.Attribute("data-omml-control-size-pt")?.Value);
        Assert.Contains(mathml.Diagnostics, diagnostic => diagnostic.ElementName == "m:ctrlPr");

        DxpOmmlConversionResult latex = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Latex);
        Assert.Equal(@"\frac{1}{2}", latex.Output);
        Assert.Contains(latex.Diagnostics, diagnostic => diagnostic.ElementName == "m:ctrlPr");
    }

    [Fact]
    public void ArgumentLevelControlPropertiesAreConsumedRatherThanTreatedAsUnknownMath()
    {
        string omml = $"""
            <m:oMath xmlns:m="{M}" xmlns:w="{W}"><m:f>
              <m:num><m:ctrlPr><w:rPr><w:color w:val="00AA00"/></w:rPr></m:ctrlPr><m:r><m:t>1</m:t></m:r></m:num>
              <m:den><m:r><m:t>2</m:t></m:r></m:den>
            </m:f></m:oMath>
            """;

        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.MathMl);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "OMML001");
        Assert.Contains("data-omml-control-color=\"00AA00\"", result.Output, StringComparison.Ordinal);
    }

    [Fact]
    public void TransparentContainersAndHyperlinksPreserveVisibleContent()
    {
        string omml = $"""
            <m:oMath xmlns:m="{M}" xmlns:w="{W}" xmlns:r="{R}">
              <w:customXml><w:smartTag><w:sdt><w:sdtContent>
                <w:hyperlink r:id="rId7" w:anchor="target"><w:r><w:t>linked</w:t></w:r></w:hyperlink>
              </w:sdtContent></w:sdt></w:smartTag></w:customXml>
            </m:oMath>
            """;
        DxpOmmlConversionOptions options = new()
        {
            IncludeHyperlinkTargets = true,
            HyperlinkTargetResolver = (id, anchor) => id == "rId7" ? $"https://example.test/#{anchor}" : null,
        };

        Assert.Equal(@"linked (https://example.test/\#target)", DxpOmmlConverter.ToLatex(omml, options));
    }

    [Fact]
    public void SimpleAndComplexFieldsUseCachedResultsOrCanBeOmitted()
    {
        string omml = $"""
            <m:oMath xmlns:m="{M}" xmlns:w="{W}"><w:customXml>
              <w:fldSimple w:instr=" DATE "><w:r><w:t>simple-result</w:t></w:r></w:fldSimple>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText> MERGEFIELD X </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r>
              <w:r><w:t>complex-result</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:customXml></m:oMath>
            """;

        Assert.Equal("simple-resultcomplex-result", DxpOmmlConverter.ToText(omml));
        Assert.Equal(string.Empty, DxpOmmlConverter.ToText(omml,
            new DxpOmmlConversionOptions { FieldMode = DxpOmmlFieldMode.Omit }));
    }

    [Fact]
    public void AdjacentTopLevelRunsRetainComplexFieldState()
    {
        string omml = $"""
            <m:oMath xmlns:m="{M}" xmlns:w="{W}">
              <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>cached</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </m:oMath>
            """;

        Assert.Equal("cached", DxpOmmlConverter.ToText(omml));
    }

    [Theory]
    [InlineData(DxpOmmlRevisionMode.Accept, false)]
    [InlineData(DxpOmmlRevisionMode.Reject, true)]
    [InlineData(DxpOmmlRevisionMode.Preserve, true)]
    public void ControlCharacterRevisionsFollowPolicy(DxpOmmlRevisionMode mode, bool styled)
    {
        string omml = $"""
            <m:oMath xmlns:m="{M}" xmlns:w="{W}"><m:f><m:fPr><m:ctrlPr><w:del><w:rPr><w:b/></w:rPr></w:del></m:ctrlPr></m:fPr>
              <m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den>
            </m:f></m:oMath>
            """;
        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.MathMl,
            new DxpOmmlConversionOptions { RevisionMode = mode });

        Assert.Equal(styled, result.Output.Contains("data-omml-control-properties", StringComparison.Ordinal));
        if (mode == DxpOmmlRevisionMode.Preserve)
            Assert.Contains("data-omml-revision=\"deleted\"", result.Output, StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.ElementName == "m:ctrlPr");
    }

    [Theory]
    [InlineData(DxpOmmlRevisionMode.Accept, "newto")]
    [InlineData(DxpOmmlRevisionMode.Reject, "oldfrom")]
    [InlineData(DxpOmmlRevisionMode.Preserve, "[inserted:new][deleted:old][inserted:to][deleted:from]")]
    public void RevisionPolicyCoversInsertDeleteAndMoves(DxpOmmlRevisionMode mode, string expected)
    {
        string omml = $"""
            <m:oMath xmlns:m="{M}" xmlns:w="{W}"><w:customXml>
              <w:ins><w:r><w:t>new</w:t></w:r></w:ins><w:del><w:r><w:delText>old</w:delText></w:r></w:del>
              <w:moveTo><w:r><w:t>to</w:t></w:r></w:moveTo><w:moveFrom><w:r><w:t>from</w:t></w:r></w:moveFrom>
            </w:customXml></m:oMath>
            """;

        Assert.Equal(expected, DxpOmmlConverter.ToText(omml,
            new DxpOmmlConversionOptions { RevisionMode = mode }));
    }

    [Fact]
    public void RangeMarkersAreIgnoredWithoutDiagnostics()
    {
        string omml = $"""
            <m:oMath xmlns:m="{M}" xmlns:w="{W}"><w:customXml>
              <w:bookmarkStart w:id="1" w:name="b"/><w:commentRangeStart w:id="2"/>
              <w:proofErr w:type="spellStart"/><w:r><w:t>x</w:t></w:r>
              <w:proofErr w:type="spellEnd"/><w:commentRangeEnd w:id="2"/><w:bookmarkEnd w:id="1"/>
            </w:customXml></m:oMath>
            """;

        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Text);
        Assert.Equal("x", result.Output);
        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void UnexpectedVisibleContentUsesFallbackAndDiagnostic()
    {
        string omml = $"""
            <m:oMath xmlns:m="{M}" xmlns:w="{W}"><w:drawing><w:t>alt</w:t></w:drawing></m:oMath>
            """;

        DxpOmmlConversionResult extracted = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Text);
        Assert.Equal("alt", extracted.Output);
        Assert.Contains(extracted.Diagnostics, diagnostic => diagnostic.Code == "OMML011");
        Assert.Equal("?", DxpOmmlConverter.ToText(omml,
            new DxpOmmlConversionOptions { FallbackPolicy = DxpOmmlFallbackPolicy.Placeholder, Placeholder = "?" }));
        Assert.Throws<DxpOmmlUnsupportedException>(() => DxpOmmlConverter.ToText(omml,
            new DxpOmmlConversionOptions { FallbackPolicy = DxpOmmlFallbackPolicy.Throw }));
    }
}
