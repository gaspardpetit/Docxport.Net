using System.Xml.Linq;
using DocxportNet.Omml;
using Mx = DocumentFormat.OpenXml.Math;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Tests.Omml;

public sealed class DxpOmmlEquationLayoutTests
{
    private const string M = OmmlTestData.MathNamespace;
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";

    [Fact]
    public void RunBreakPreservesBoundaryAndAlignmentIndex()
    {
        string omml = Inline(Run("a") + BreakRun("+b", "3"));

        DxpOmmlConversionResult mathMl = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.MathMl);
        XElement boundary = XElement.Parse(mathMl.Output).Descendants(MathMl + "mspace").Single();
        Assert.Equal("newline", (string?)boundary.Attribute("linebreak"));
        Assert.Equal("3", (string?)boundary.Attribute("data-omml-align-at"));
        Assert.Empty(mathMl.Diagnostics);

        DxpOmmlConversionResult latex = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Latex);
        Assert.Equal(@"\begin{aligned}a\\+b\end{aligned}", latex.Output);
        AssertAlignmentApproximation(latex);
        Assert.Equal("a\n+b", DxpOmmlConverter.ToUnicodeMath(omml));
        Assert.Equal("a\n+b", DxpOmmlConverter.ToText(omml));
    }

    [Fact]
    public void MultipleRunBreaksUseTheSchemaDefaultAlignmentIndex()
    {
        string omml = Inline(Run("a") + BreakRun("b") + BreakRun("c"));
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml));

        Assert.Equal(["0", "0"], math.Descendants(MathMl + "mspace")
            .Select(element => (string?)element.Attribute("data-omml-align-at")));
        Assert.Equal("a\nb\nc", DxpOmmlConverter.ToText(omml));
        Assert.Equal(@"\begin{aligned}a\\b\\c\end{aligned}", DxpOmmlConverter.ToLatex(omml));
    }

    [Theory]
    [MemberData(nameof(NestedRunBreaks))]
    public void RunBreakWorksInsideEverySupportedParentStructure(string structure)
    {
        string omml = Inline(structure);
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml));

        Assert.Single(math.Descendants(MathMl + "mspace"),
            element => (string?)element.Attribute("linebreak") == "newline");
        Assert.Contains('\n', DxpOmmlConverter.ToUnicodeMath(omml));
        Assert.Contains('\n', DxpOmmlConverter.ToText(omml));
        Assert.Contains(@"\begin{aligned}", DxpOmmlConverter.ToLatex(omml));
    }

    [Theory]
    [InlineData("left", "left")]
    [InlineData("right", "right")]
    [InlineData("center", "center")]
    [InlineData("centerGroup", "centerGroup")]
    public void AppliesEveryMathParagraphJustification(string value, string expected)
    {
        string omml = Paragraph($"<m:oMathParaPr><m:jc m:val=\"{value}\"/></m:oMathParaPr>",
            $"<m:oMath>{Run("a")}</m:oMath>");
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml));

        Assert.Equal("block", (string?)math.Attribute("display"));
        Assert.Equal(expected, (string?)math.Attribute("data-omml-justification"));
        Assert.Contains(value == "right" ? "right" : value == "left" ? "left" : "center",
            (string?)math.Attribute("style"));
    }

    [Fact]
    public void LocalParagraphJustificationOverridesDocumentDefault()
    {
        string omml = Paragraph("<m:oMathParaPr><m:jc m:val=\"right\"/></m:oMathParaPr>",
            $"<m:oMath>{Run("x")}</m:oMath>");
        DxpOmmlConversionOptions options = new() { DefaultJustification = DxpOmmlJustification.Left };

        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml, options));

        Assert.Equal("right", (string?)math.Attribute("data-omml-justification"));
    }

    [Fact]
    public void ParagraphBreakPreservesMultipleEquationsWithoutFallback()
    {
        string omml = Paragraph(string.Empty,
            $"<m:oMath>{Run("a")}</m:oMath><m:r><w:br/></m:r><m:oMath>{Run("b")}</m:oMath>");

        DxpOmmlConversionResult mathMl = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.MathMl);
        Assert.DoesNotContain(mathMl.Diagnostics, diagnostic => diagnostic.ElementName == "m:r");
        Assert.Equal(2, XElement.Parse(mathMl.Output).Descendants(MathMl + "mtr").Count());
        Assert.Equal("a\nb", DxpOmmlConverter.ToText(omml));
        Assert.Equal("a\nb", DxpOmmlConverter.ToUnicodeMath(omml));
        Assert.Equal(@"\begin{aligned}a\\b\end{aligned}", DxpOmmlConverter.ToLatex(omml));
    }

    [Fact]
    public void WordBreakInsideAnEquationRunIsAlsoSemantic()
    {
        string omml = Inline("<m:r><m:t>a</m:t><w:br/><m:t>b</m:t></m:r>")
            .Replace("<m:oMath ", "<m:oMath xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ");

        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mspace"));
        Assert.Equal("a\nb", DxpOmmlConverter.ToText(omml));
        Assert.Equal(@"\begin{aligned}a\\b\end{aligned}", DxpOmmlConverter.ToLatex(omml));
    }

    [Fact]
    public void ParagraphBreakRunDoesNotDiscardAdjacentVisibleContent()
    {
        string omml = Paragraph(string.Empty,
            $"<m:oMath>{Run("a")}</m:oMath><m:r><m:t>left</m:t><w:br/><m:t>right</m:t></m:r>");

        Assert.Equal("aleft\nright", DxpOmmlConverter.ToText(omml));
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml));
        Assert.Equal(2, math.Elements(MathMl + "mtable").Elements(MathMl + "mtr").Count());
    }

    [Fact]
    public void PreservesRelativeAlignmentMarkersAcrossParagraphEquations()
    {
        string alignedRun = "<m:r><m:rPr><m:aln/></m:rPr><m:t>=</m:t></m:r>";
        string omml = Paragraph("<m:oMathParaPr><m:jc m:val=\"centerGroup\"/></m:oMathParaPr>",
            $"<m:oMath>{Run("a")}{alignedRun}</m:oMath><m:r><w:br/></m:r><m:oMath>{Run("b")}{alignedRun}</m:oMath>");

        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml));
        Assert.Equal(2, math.Descendants(MathMl + "malignmark").Count());
        Assert.Equal("center", (string?)math.Descendants(MathMl + "mtable").Single().Attribute("columnalign"));
        Assert.Equal(@"\begin{aligned}a&=\\b&=\end{aligned}", DxpOmmlConverter.ToLatex(omml));
    }

    [Theory]
    [InlineData(DxpOmmlBreakBinary.Before, "before")]
    [InlineData(DxpOmmlBreakBinary.After, "after")]
    [InlineData(DxpOmmlBreakBinary.Repeat, "repeat")]
    public void MapsEveryBinaryBreakSetting(DxpOmmlBreakBinary value, string expected)
    {
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(Inline(Run("x")),
            new DxpOmmlConversionOptions { BreakBinary = value }));
        Assert.Equal(expected, (string?)math.Attribute("data-omml-break-binary"));
    }

    [Theory]
    [InlineData(DxpOmmlBreakBinarySubtraction.MinusMinus, "--")]
    [InlineData(DxpOmmlBreakBinarySubtraction.MinusPlus, "-+")]
    [InlineData(DxpOmmlBreakBinarySubtraction.PlusMinus, "+-")]
    public void MapsEverySubtractionBreakSetting(DxpOmmlBreakBinarySubtraction value, string expected)
    {
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(Inline(Run("x")),
            new DxpOmmlConversionOptions { BreakBinarySubtraction = value }));
        Assert.Equal(expected, (string?)math.Attribute("data-omml-break-binary-subtraction"));
    }

    [Theory]
    [InlineData(DxpOmmlBreakBinary.Before, "a\n+b", @"\begin{aligned}a\\+b\end{aligned}")]
    [InlineData(DxpOmmlBreakBinary.After, "a+\nb", @"\begin{aligned}a+\\b\end{aligned}")]
    [InlineData(DxpOmmlBreakBinary.Repeat, "a+\n+b", @"\begin{aligned}a+\\+b\end{aligned}")]
    public void AppliesBinaryBreakPlacement(DxpOmmlBreakBinary value, string text, string latex)
    {
        string omml = Inline(Run("a") + BreakRun("+b"));
        DxpOmmlConversionOptions options = new() { BreakBinary = value };

        Assert.Equal(text, DxpOmmlConverter.ToText(omml, options));
        Assert.Equal(latex, DxpOmmlConverter.ToLatex(omml, options));
    }

    [Theory]
    [InlineData(DxpOmmlBreakBinarySubtraction.MinusMinus, "a-\n-b")]
    [InlineData(DxpOmmlBreakBinarySubtraction.MinusPlus, "a-\n+b")]
    [InlineData(DxpOmmlBreakBinarySubtraction.PlusMinus, "a+\n-b")]
    public void AppliesRepeatedSubtractionPolicy(DxpOmmlBreakBinarySubtraction value, string expected)
    {
        string omml = Inline(Run("a") + BreakRun("-b"));
        DxpOmmlConversionOptions options = new()
        {
            BreakBinary = DxpOmmlBreakBinary.Repeat,
            BreakBinarySubtraction = value,
        };

        Assert.Equal(expected, DxpOmmlConverter.ToText(omml, options));
    }

    [Fact]
    public void AppliesBinaryBreakPlacementToOperatorEmulatorBoxes()
    {
        string box = $"<m:box><m:boxPr><m:opEmu/><m:brk/></m:boxPr><m:e>{Run("=")}</m:e></m:box>";
        string omml = Inline(Run("a") + box + Run("b"));
        DxpOmmlConversionOptions options = new() { BreakBinary = DxpOmmlBreakBinary.Repeat };

        Assert.Equal("a=\n=b", DxpOmmlConverter.ToText(omml, options));
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml, options));
        Assert.Equal(2, math.Descendants(MathMl + "mo").Count(element => element.Value == "="));
        Assert.Single(math.Descendants(MathMl + "mspace"),
            element => (string?)element.Attribute("linebreak") == "newline");
    }

    [Fact]
    public void ExposesEveryDocumentMathSettingThroughOptions()
    {
        DxpOmmlConversionOptions options = new()
        {
            DisplayDefaults = true,
            MathFont = "Cambria Math",
            BreakBinary = DxpOmmlBreakBinary.Repeat,
            BreakBinarySubtraction = DxpOmmlBreakBinarySubtraction.MinusPlus,
            DefaultJustification = DxpOmmlJustification.Left,
            LeftMarginTwips = 10,
            RightMarginTwips = 20,
            PreSpacingTwips = 30,
            PostSpacingTwips = 40,
            InterSpacingTwips = 50,
            IntraSpacingTwips = 60,
            WrapIndentTwips = 70,
            WrapRight = true,
        };

        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(Inline(Run("x")),
            DxpOmmlOutputFormat.MathMl, options);
        XElement math = XElement.Parse(result.Output);

        Assert.True(result.IsDisplay);
        Assert.Equal("block", (string?)math.Attribute("display"));
        Assert.Equal("Cambria Math", (string?)math.Attribute("data-omml-math-font"));
        Assert.Equal("repeat", (string?)math.Attribute("data-omml-break-binary"));
        Assert.Equal("-+", (string?)math.Attribute("data-omml-break-binary-subtraction"));
        Assert.Equal("left", (string?)math.Attribute("data-omml-justification"));
        Assert.Equal("10", (string?)math.Attribute("data-omml-left-margin-twips"));
        Assert.Equal("20", (string?)math.Attribute("data-omml-right-margin-twips"));
        Assert.Equal("30", (string?)math.Attribute("data-omml-pre-spacing-twips"));
        Assert.Equal("40", (string?)math.Attribute("data-omml-post-spacing-twips"));
        Assert.Equal("50", (string?)math.Attribute("data-omml-inter-spacing-twips"));
        Assert.Equal("60", (string?)math.Attribute("data-omml-intra-spacing-twips"));
        Assert.Null(math.Attribute("data-omml-wrap-indent-twips"));
        Assert.Equal("true", (string?)math.Attribute("data-omml-wrap-right"));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OMML002" &&
            diagnostic.ElementName == "m:mathPr");
    }

    [Fact]
    public void WrapRightTakesPrecedenceOverTheMutuallyExclusiveWrapIndent()
    {
        DxpOmmlConversionOptions indent = new() { WrapIndentTwips = 1440 };
        DxpOmmlConversionOptions right = new() { WrapIndentTwips = 1440, WrapRight = true };

        Assert.Equal("1440", (string?)XElement.Parse(DxpOmmlConverter.ToMathMl(Inline(Run("x")), indent))
            .Attribute("data-omml-wrap-indent-twips"));
        XElement rightMath = XElement.Parse(DxpOmmlConverter.ToMathMl(Inline(Run("x")), right));
        Assert.Null(rightMath.Attribute("data-omml-wrap-indent-twips"));
        Assert.Equal("true", (string?)rightMath.Attribute("data-omml-wrap-right"));
    }

    [Fact]
    public void ParsesTypedSdkBreakAndParagraphLayout()
    {
        Mx.Run broken = new(new Mx.RunProperties(new Mx.Break { AlignAt = 2 }), new Mx.Text("b"));
        Mx.OfficeMath inline = new(new Mx.Run(new Mx.Text("a")), broken);
        Assert.Equal("a\nb", DxpOmmlConverter.Convert(inline, DxpOmmlOutputFormat.Text).Output);

        Mx.Paragraph paragraph = new(
            new Mx.ParagraphProperties(new Mx.Justification { Val = Mx.JustificationValues.Right }),
            new Mx.OfficeMath(new Mx.Run(new Mx.Text("x"))),
            new Mx.Run(new W.Break()),
            new Mx.OfficeMath(new Mx.Run(new Mx.Text("y"))));
        XElement math = XElement.Parse(DxpOmmlConverter.Convert(paragraph, DxpOmmlOutputFormat.MathMl).Output);
        Assert.Equal("right", (string?)math.Attribute("data-omml-justification"));
        Assert.Equal(2, math.Descendants(MathMl + "mtr").Count());
    }

    [Fact]
    public void CoversEveryPinnedUpstreamLineBreakFixture()
    {
        string directory = Path.Combine(OmmlTestData.UpstreamRoot, "line_break");
        string[] fixtures = Directory.EnumerateFiles(directory, "*.omml").Order().ToArray();
        Assert.Equal(90, fixtures.Length);

        foreach (string fixture in fixtures)
        {
            string omml = File.ReadAllText(fixture);
            int expectedBreaks = XElement.Parse(omml).Descendants()
                .Count(element => element.Name.NamespaceName == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" &&
                                  element.Name.LocalName is "br" or "cr");
            DxpOmmlConversionResult[] results = Enum.GetValues<DxpOmmlOutputFormat>()
                .Select(format => DxpOmmlConverter.Convert(omml, format)).ToArray();
            Assert.All(results, result => Assert.DoesNotContain(result.Diagnostics,
                diagnostic => diagnostic.ElementName == "m:r"));
            DxpOmmlConversionResult mathMl = results.Single(result => result.Format == DxpOmmlOutputFormat.MathMl);
            XElement math = XElement.Parse(mathMl.Output);
            Assert.Equal(expectedBreaks == 0 ? 0 : expectedBreaks + 1,
                math.Elements(MathMl + "mtable").Elements(MathMl + "mtr").Count());
            Assert.Equal(expectedBreaks, results.Single(result => result.Format == DxpOmmlOutputFormat.Text)
                .Output.Count(character => character == '\n'));
            Assert.Equal(expectedBreaks, results.Single(result => result.Format == DxpOmmlOutputFormat.UnicodeMath)
                .Output.Count(character => character == '\n'));
            if (expectedBreaks != 0)
                Assert.StartsWith(@"\begin{aligned}", results.Single(result => result.Format == DxpOmmlOutputFormat.Latex).Output);
        }
    }

    public static TheoryData<string> NestedRunBreaks => new()
    {
        $"<m:f><m:num>{BreakRun("a")}</m:num><m:den>{Run("b")}</m:den></m:f>",
        $"<m:rad><m:deg>{Run("2")}</m:deg><m:e>{BreakRun("x")}</m:e></m:rad>",
        $"<m:sSub><m:e>{BreakRun("x")}</m:e><m:sub>{Run("1")}</m:sub></m:sSub>",
        $"<m:d><m:e>{BreakRun("x")}</m:e></m:d>",
        $"<m:bar><m:e>{BreakRun("x")}</m:e></m:bar>",
        $"<m:func><m:fName>{Run("sin")}</m:fName><m:e>{BreakRun("x")}</m:e></m:func>",
        $"<m:limLow><m:e>{BreakRun("x")}</m:e><m:lim>{Run("0")}</m:lim></m:limLow>",
        $"<m:nary><m:sub>{Run("0")}</m:sub><m:sup>{Run("1")}</m:sup><m:e>{BreakRun("x")}</m:e></m:nary>",
        $"<m:m><m:mr><m:e>{BreakRun("x")}</m:e></m:mr></m:m>",
        $"<m:eqArr><m:e>{BreakRun("x")}</m:e></m:eqArr>",
        $"<m:box><m:e>{BreakRun("x")}</m:e></m:box>",
        $"<m:borderBox><m:e>{BreakRun("x")}</m:e></m:borderBox>",
        $"<m:phant><m:e>{BreakRun("x")}</m:e></m:phant>",
    };

    private static void AssertAlignmentApproximation(DxpOmmlConversionResult result) =>
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OMML002" &&
            diagnostic.ElementName == "m:brk");

    private static string Run(string text) => $"<m:r><m:t>{text}</m:t></m:r>";
    private static string BreakRun(string text, string? alignmentAt = null) =>
        $"<m:r><m:rPr><m:brk{(alignmentAt == null ? string.Empty : $" m:alnAt=\"{alignmentAt}\"")}/></m:rPr><m:t>{text}</m:t></m:r>";
    private static string Inline(string content) => $"<m:oMath xmlns:m=\"{M}\">{content}</m:oMath>";
    private static string Paragraph(string properties, string content) =>
        $"<m:oMathPara xmlns:m=\"{M}\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">{properties}{content}</m:oMathPara>";
}
