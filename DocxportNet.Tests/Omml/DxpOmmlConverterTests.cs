using System.Globalization;
using System.Xml.Linq;
using DocxportNet.Omml;
using M = DocumentFormat.OpenXml.Math;

namespace DocxportNet.Tests.Omml;

public sealed class DxpOmmlConverterTests
{
    private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    [Fact]
    public void ConvertsInlineFallbackToAllFourOutputFormats()
    {
        string omml = Inline("x&amp;y_");

        DxpOmmlConversionResult mathml = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.MathMl);
        DxpOmmlConversionResult latex = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Latex);
        DxpOmmlConversionResult unicodeMath = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.UnicodeMath);
        DxpOmmlConversionResult text = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Text);

        XNamespace math = "http://www.w3.org/1998/Math/MathML";
        XElement mathElement = XElement.Parse(mathml.Output);
        Assert.Equal(math + "math", mathElement.Name);
        Assert.Equal("inline", mathElement.Attribute("display")?.Value);
        Assert.Equal("x&y_", mathElement.Descendants(math + "mtext").Single().Value);
        Assert.Equal(@"x\&y\_", latex.Output);
        Assert.Equal("x&y_", unicodeMath.Output);
        Assert.Equal("x&y_", text.Output);
        Assert.All(new[] { mathml, latex, unicodeMath, text }, result =>
        {
            Assert.False(result.IsDisplay);
            Assert.True(result.IsLossy);
            DxpOmmlDiagnostic diagnostic = Assert.Single(result.Diagnostics);
            Assert.Equal("OMML001", diagnostic.Code);
            Assert.Equal("m:r", diagnostic.ElementName);
            Assert.Equal("/m:oMath[1]/m:r[1]", diagnostic.Path);
        });
    }

    [Fact]
    public void DisplayParagraphPreservesAdjacentExpressionOrderAndCanBeOverridden()
    {
        string omml = $"""
            <m:oMathPara xmlns:m="{MathNamespace}">
              <m:oMath><m:r><m:t>a</m:t></m:r></m:oMath>
              <m:oMath><m:r><m:t>b</m:t></m:r></m:oMath>
            </m:oMathPara>
            """;

        DxpOmmlConversionResult inferred = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Text);
        string overridden = DxpOmmlConverter.ToMathMl(omml, new DxpOmmlConversionOptions { Display = false });

        Assert.True(inferred.IsDisplay);
        Assert.Equal("ab", inferred.Output);
        Assert.Equal(2, inferred.Diagnostics.Count);
        Assert.Equal("inline", XElement.Parse(overridden).Attribute("display")?.Value);
    }

    [Theory]
    [InlineData(DxpOmmlFallbackPolicy.ExtractText, "x")]
    [InlineData(DxpOmmlFallbackPolicy.Placeholder, "?")]
    [InlineData(DxpOmmlFallbackPolicy.Omit, "")]
    public void AppliesConfiguredFallbackPolicy(DxpOmmlFallbackPolicy policy, string expected)
    {
        DxpOmmlConversionOptions options = new() { FallbackPolicy = policy, Placeholder = "?" };

        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(
            Inline("x"),
            DxpOmmlOutputFormat.Text,
            options);

        Assert.Equal(expected, result.Output);
        Assert.Single(result.Diagnostics);
    }

    [Fact]
    public void ThrowFallbackIncludesElementAndPathDiagnostic()
    {
        DxpOmmlConversionOptions options = new() { FallbackPolicy = DxpOmmlFallbackPolicy.Throw };

        DxpOmmlUnsupportedException exception = Assert.Throws<DxpOmmlUnsupportedException>(() =>
            DxpOmmlConverter.ToText(Inline("x"), options));

        Assert.Equal("m:r", exception.Diagnostic.ElementName);
        Assert.Equal("/m:oMath[1]/m:r[1]", exception.Diagnostic.Path);
    }

    [Fact]
    public void TryConvertReportsUnsupportedValidOmmlSeparatelyFromMalformedXml()
    {
        bool converted = DxpOmmlConverter.TryConvert(
            Inline("x"),
            DxpOmmlOutputFormat.Text,
            out DxpOmmlConversionResult? result,
            out DxpOmmlException? error,
            new DxpOmmlConversionOptions { FallbackPolicy = DxpOmmlFallbackPolicy.Throw });

        Assert.False(converted);
        Assert.Null(result);
        Assert.IsType<DxpOmmlUnsupportedException>(error);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("<root />")]
    [InlineData("<m:oMathPara xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' />")]
    [InlineData("<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'>")]
    [InlineData("<!DOCTYPE x [<!ENTITY e 'x'>]><x>&e;</x>")]
    public void TryConvertReturnsTypedParseFailure(string? input)
    {
        bool converted = DxpOmmlConverter.TryConvert(
            input,
            DxpOmmlOutputFormat.Text,
            out DxpOmmlConversionResult? result,
            out DxpOmmlException? error);

        Assert.False(converted);
        Assert.Null(result);
        Assert.IsType<DxpOmmlParseException>(error);
    }

    [Fact]
    public void RejectsInputBeyondConfiguredLimit()
    {
        DxpOmmlConversionOptions options = new() { MaxInputCharacters = 10 };

        DxpOmmlParseException exception = Assert.Throws<DxpOmmlParseException>(() =>
            DxpOmmlConverter.ToText(Inline("x"), options));

        Assert.Contains("character limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AcceptsAlternateAndDefaultOmmlNamespacePrefixes()
    {
        string alternate = $"<q:oMath xmlns:q=\"{MathNamespace}\"><q:r><q:t>a</q:t></q:r></q:oMath>";
        string defaultNamespace = $"<oMath xmlns=\"{MathNamespace}\"><r><t>b</t></r></oMath>";

        Assert.Equal("a", DxpOmmlConverter.ToText(alternate));
        Assert.Equal("b", DxpOmmlConverter.ToText(defaultNamespace));
    }

    [Fact]
    public void UnexpectedVisibleRootTextIsDiagnosedInsteadOfDiscarded()
    {
        string omml = $"<m:oMath xmlns:m=\"{MathNamespace}\">visible</m:oMath>";

        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Text);

        Assert.Equal("visible", result.Output);
        Assert.Equal("#text", Assert.Single(result.Diagnostics).ElementName);
    }

    [Fact]
    public void FallbackLineEndingsAreDeterministic()
    {
        string omml = $"""
            <m:oMath xmlns:m="{MathNamespace}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <m:r><m:t>a</m:t><w:br/><m:t>b</m:t></m:r>
            </m:oMath>
            """;

        Assert.Equal("a\nb", DxpOmmlConverter.ToText(omml));
        Assert.Equal("a\nb", DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Fact]
    public void ConvertsOpenXmlSdkMathRootsWithoutDocumentContext()
    {
        M.OfficeMath inline = new(new M.Run(new M.Text("x")));
        M.Paragraph display = new(new M.OfficeMath(new M.Run(new M.Text("y"))));

        Assert.Equal("x", DxpOmmlConverter.Convert(inline, DxpOmmlOutputFormat.Text).Output);
        Assert.True(DxpOmmlConverter.Convert(display, DxpOmmlOutputFormat.Text).IsDisplay);
    }

    [Fact]
    public void OutputIsCultureInvariantAndHtmlReadyOutputIsMathMl()
    {
        CultureInfo originalCulture = CultureInfo.CurrentCulture;
        CultureInfo originalUiCulture = CultureInfo.CurrentUICulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ar-SA");
            CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo("ar-SA");
            string first = DxpOmmlConverter.ToMathMl(Inline("123"));
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-CA");
            CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo("fr-CA");

            Assert.Equal(first, DxpOmmlConverter.ToMathMl(Inline("123")));
            Assert.Equal(first, DxpOmmlConverter.ToHtml(Inline("123")));
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
            CultureInfo.CurrentUICulture = originalUiCulture;
        }
    }

    [Fact]
    public async Task ConverterHasNoSharedStateAcrossConcurrentCalls()
    {
        Task<string>[] conversions = Enumerable.Range(0, 100)
            .Select(index => Task.Run(() => DxpOmmlConverter.ToUnicodeMath(Inline(index.ToString(CultureInfo.InvariantCulture)))))
            .ToArray();

        string[] outputs = await Task.WhenAll(conversions);

        Assert.Equal(Enumerable.Range(0, 100).Select(index => index.ToString(CultureInfo.InvariantCulture)), outputs);
    }

    [Fact]
    public void StandaloneImplementationDoesNotReferencePipelineTypesOrReparseSdkElements()
    {
        string sourceRoot = Path.Combine(OmmlTestData.FixtureRoot, "..", "..", "..", "DocxportNet", "Omml");
        string source = string.Join("\n", Directory.EnumerateFiles(sourceRoot, "*.cs")
            .OrderBy(path => path, StringComparer.Ordinal)
            .Select(File.ReadAllText));

        Assert.DoesNotContain("DxpWalker", source, StringComparison.Ordinal);
        Assert.DoesNotContain("DxpIDocumentContext", source, StringComparison.Ordinal);
        Assert.DoesNotContain("DxpVisitor", source, StringComparison.Ordinal);
        Assert.DoesNotContain(".OuterXml", source, StringComparison.Ordinal);
    }

    [Fact]
    public void FoundationParserAcceptsEveryWellFormedPinnedCorpusFixture()
    {
        List<string> rejected = new();
        foreach (string fixture in OmmlTestData.UpstreamFixtures())
        {
            try
            {
                DxpOmmlConverter.ToText(File.ReadAllText(fixture));
            }
            catch (DxpOmmlParseException)
            {
                rejected.Add(Path.GetFileName(fixture));
            }
        }

        Assert.Equal(["187.omml", "issue-158.omml"], rejected.OrderBy(name => name, StringComparer.Ordinal));
    }

    private static string Inline(string text) =>
        $"<m:oMath xmlns:m=\"{MathNamespace}\"><m:r><m:t>{text}</m:t></m:r></m:oMath>";
}
