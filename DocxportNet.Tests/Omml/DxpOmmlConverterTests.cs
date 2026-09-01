using System.Globalization;
using System.Xml.Linq;
using DocxportNet.Omml;
using M = DocumentFormat.OpenXml.Math;

namespace DocxportNet.Tests.Omml;

public sealed class DxpOmmlConverterTests
{
    private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    [Fact]
    public void ConvertsInlineRunToAllFourOutputFormats()
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
        Assert.Equal("x&y_", string.Concat(mathElement.Descendants().Where(e => e.Name == math + "mi" || e.Name == math + "mo").Select(e => e.Value)));
        Assert.Equal(@"x\&y\_", latex.Output);
        Assert.Equal("x&y_", unicodeMath.Output);
        Assert.Equal("x&y_", text.Output);
        Assert.All(new[] { mathml, latex, unicodeMath, text }, result =>
        {
            Assert.False(result.IsDisplay);
            Assert.False(result.IsLossy);
            Assert.Empty(result.Diagnostics);
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
        Assert.Empty(inferred.Diagnostics);
        Assert.Equal("inline", XElement.Parse(overridden).Attribute("display")?.Value);
    }

    [Theory]
    [InlineData(DxpOmmlFallbackPolicy.ExtractText, "x")]
    [InlineData(DxpOmmlFallbackPolicy.Placeholder, "?")]
    [InlineData(DxpOmmlFallbackPolicy.Omit, "")]
    public void AppliesConfiguredFallbackPolicyToUnsupportedElements(DxpOmmlFallbackPolicy policy, string expected)
    {
        DxpOmmlConversionOptions options = new() { FallbackPolicy = policy, Placeholder = "?" };

        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(
            $"<m:oMath xmlns:m=\"{MathNamespace}\"><m:unknown><m:t>x</m:t></m:unknown></m:oMath>",
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
            DxpOmmlConverter.ToText($"<m:oMath xmlns:m=\"{MathNamespace}\"><m:unknown><m:t>x</m:t></m:unknown></m:oMath>", options));

        Assert.Equal("m:unknown", exception.Diagnostic.ElementName);
        Assert.Equal("/m:oMath[1]/m:unknown[1]", exception.Diagnostic.Path);
    }

    [Fact]
    public void TryConvertReportsUnsupportedValidOmmlSeparatelyFromMalformedXml()
    {
        bool converted = DxpOmmlConverter.TryConvert(
            $"<m:oMath xmlns:m=\"{MathNamespace}\"><m:unknown><m:t>x</m:t></m:unknown></m:oMath>",
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
    public void RejectsXmlBeyondConfiguredDepthBeforeSemanticRecursion()
    {
        string nested = Inline($"<m:box><m:e><m:box><m:e><m:r><m:t>x</m:t></m:r></m:e></m:box></m:e></m:box>");
        DxpOmmlConversionOptions options = new() { MaxNestingDepth = 5 };

        DxpOmmlResourceLimitException exception = Assert.Throws<DxpOmmlResourceLimitException>(() =>
            DxpOmmlConverter.ToText(nested, options));

        Assert.Contains("nesting-depth", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectsOpenXmlSdkTreeBeyondConfiguredElementCount()
    {
        M.OfficeMath inline = new(new M.Run(new M.Text("x")));
        DxpOmmlConversionOptions options = new() { MaxElementCount = 2 };

        DxpOmmlResourceLimitException exception = Assert.Throws<DxpOmmlResourceLimitException>(() =>
            DxpOmmlConverter.Convert(inline, DxpOmmlOutputFormat.Text, options));

        Assert.Contains("element-count", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectsOutputBeyondConfiguredLimitAndTryConvertReportsIt()
    {
        DxpOmmlConversionOptions options = new() { MaxOutputCharacters = 3 };

        bool converted = DxpOmmlConverter.TryConvert(
            Inline("abcd"), DxpOmmlOutputFormat.Text, out DxpOmmlConversionResult? result,
            out DxpOmmlException? error, options);

        Assert.False(converted);
        Assert.Null(result);
        Assert.IsType<DxpOmmlResourceLimitException>(error);
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
    public void ClassifiesMixedTokensAndPreservesSupplementaryScalars()
    {
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(Inline("x+12𝑦")));
        XNamespace m = "http://www.w3.org/1998/Math/MathML";
        Assert.Equal(new[] { "mi:x", "mo:+", "mn:12", "mi:𝑦" },
            math.Descendants().Where(e => e.Name != m + "mrow").Select(e => $"{e.Name.LocalName}:{e.Value}"));
    }

    [Fact]
    public void AppliesOmmlScriptAndStyleWithoutChangingTextOutputs()
    {
        string omml = $"<m:oMath xmlns:m=\"{MathNamespace}\"><m:r><m:rPr><m:scr m:val=\"sans-serif\"/><m:sty m:val=\"bi\"/></m:rPr><m:t>x</m:t></m:r></m:oMath>";
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml));
        XNamespace m = "http://www.w3.org/1998/Math/MathML";
        Assert.Equal("sans-serif-bold-italic", math.Descendants(m + "mstyle").Single().Attribute("mathvariant")?.Value);
        Assert.Equal(@"\boldsymbol{\mathsf{x}}", DxpOmmlConverter.ToLatex(omml));
        Assert.Equal("\\mbfitsans\"x\"", DxpOmmlConverter.ToUnicodeMath(omml));
        Assert.Equal("x", DxpOmmlConverter.ToText(omml));
    }

    [Fact]
    public void NormalAndLiteralRunsBecomeTextTokens()
    {
        string omml = $"<m:oMath xmlns:m=\"{MathNamespace}\"><m:r><m:rPr><m:nor/></m:rPr><m:t>sin x</m:t></m:r><m:r><m:rPr><m:lit/></m:rPr><m:t>+1</m:t></m:r></m:oMath>";
        XNamespace m = "http://www.w3.org/1998/Math/MathML";
        Assert.Equal(new[] { "sin x", "+1" }, XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(m + "mtext").Select(e => e.Value));
    }

    [Fact]
    public void ResolvesWordSymbolAndHandlesInvisibleCharacterDeliberately()
    {
        string omml = $"<m:oMath xmlns:m=\"{MathNamespace}\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><m:r><w:sym w:font=\"Symbol\" w:char=\"F06D\"/><m:t>​</m:t></m:r></m:oMath>";
        Assert.StartsWith("µ", DxpOmmlConverter.ToText(omml), StringComparison.Ordinal);
        Assert.Equal("µ", DxpOmmlConverter.ToText(omml));
        Assert.Equal("µ​", DxpOmmlConverter.ToUnicodeMath(omml));
        Assert.Equal("µ{}", DxpOmmlConverter.ToLatex(omml));
        XNamespace m = "http://www.w3.org/1998/Math/MathML";
        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(m + "mspace"));
    }

    [Fact]
    public void PreservesWhitespaceEmptyTextAndDecimalTokens()
    {
        string omml = $"<m:oMath xmlns:m=\"{MathNamespace}\"><m:r><m:t></m:t><m:t xml:space=\"preserve\"> 12.5 </m:t></m:r></m:oMath>";
        XNamespace m = "http://www.w3.org/1998/Math/MathML";
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml));
        Assert.Equal("12.5", math.Descendants(m + "mn").Single().Value);
        Assert.Equal(" 12.5 ", DxpOmmlConverter.ToText(omml));
    }

    [Fact]
    public void PinnedStyleFixtureCoversEverySupportedMathVariant()
    {
        string fixture = Path.Combine(OmmlTestData.UpstreamRoot, "184.omml");
        XNamespace m = "http://www.w3.org/1998/Math/MathML";
        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(File.ReadAllText(fixture), DxpOmmlOutputFormat.MathMl);
        string[] variants = XElement.Parse(result.Output).Descendants(m + "mstyle")
            .Select(e => (string?)e.Attribute("mathvariant")).Where(v => v != null).Cast<string>().ToArray();
        Assert.Equal(new[] { "normal", "bold", "italic", "bold-italic", "double-struck", "bold-fraktur", "script", "bold-script", "fraktur", "sans-serif", "bold-sans-serif", "sans-serif-italic", "sans-serif-bold-italic", "monospace" }, variants);
        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void SymbolFontTextIsTranslatedWhenRunFontIsKnown()
    {
        string omml = $"<m:oMath xmlns:m=\"{MathNamespace}\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><m:r><w:rPr><w:rFonts w:ascii=\"Symbol\"/></w:rPr><m:t>m</m:t></m:r></m:oMath>";
        Assert.Equal("µ", DxpOmmlConverter.ToText(omml));
    }

    [Fact]
    public void PreservesApplicableWordFormattingLanguageDirectionAndAlignment()
    {
        string omml = $"<m:oMath xmlns:m=\"{MathNamespace}\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><m:r><m:rPr><m:aln/></m:rPr><w:rPr><w:b/><w:i/><w:rtl/><w:lang w:val=\"ar\"/></w:rPr><m:t>x</m:t></m:r></m:oMath>";
        XNamespace m = "http://www.w3.org/1998/Math/MathML";
        XElement style = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(m + "mstyle").Single();
        Assert.Equal("bold-italic", (string?)style.Attribute("mathvariant"));
        Assert.Equal("rtl", (string?)style.Attribute("dir"));
        Assert.Equal("ar", (string?)style.Attribute(XNamespace.Xml + "lang"));
        Assert.Single(style.Elements(m + "malignmark"));
        Assert.Equal(@"&\boldsymbol{x}", DxpOmmlConverter.ToLatex(omml));
        Assert.StartsWith("&", DxpOmmlConverter.ToUnicodeMath(omml), StringComparison.Ordinal);
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
    public void DirectMathTextInAnArgumentIsRecoveredWithoutLosingVisibleContent()
    {
        string omml = $"""
            <m:oMath xmlns:m="{MathNamespace}">
              <m:nary><m:sub><m:t>symmetric</m:t></m:sub><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary>
            </m:oMath>
            """;

        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Text);

        Assert.Contains("symmetric", result.Output, StringComparison.Ordinal);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "OMML001");
    }

    [Fact]
    public void LatexControlSequencesFromOmmlTextAreAlwaysEscapedAsText()
    {
        string fixture = Path.Combine(OmmlTestData.NormativeRoot, "latex-injection-is-text.omml");

        Assert.Equal(@"\textbackslash{}end\{array\}\$\%\#\&\_\^{}\{\}",
            DxpOmmlConverter.ToLatex(File.ReadAllText(fixture)));
    }

    [Fact]
    public void NestedMathAndFutureExtensionsRecoverVisibleTextWithDiagnostics()
    {
        string nested = $"<m:oMath xmlns:m=\"{MathNamespace}\"><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMath>";
        string extended = $"<m:oMath xmlns:m=\"{MathNamespace}\"><m:r future=\"ignored\"><m:t>x</m:t></m:r><m:future><m:t>y</m:t></m:future></m:oMath>";

        DxpOmmlConversionResult nestedResult = DxpOmmlConverter.Convert(nested, DxpOmmlOutputFormat.Text);
        DxpOmmlConversionResult extendedResult = DxpOmmlConverter.Convert(extended, DxpOmmlOutputFormat.Text);

        Assert.Equal("x", nestedResult.Output);
        Assert.Contains(nestedResult.Diagnostics, diagnostic => diagnostic.ElementName == "m:oMath");
        Assert.Equal("xy", extendedResult.Output);
        Assert.Contains(extendedResult.Diagnostics, diagnostic => diagnostic.ElementName == "m:future");
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
