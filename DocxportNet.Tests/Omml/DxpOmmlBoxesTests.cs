using System.Xml.Linq;
using DocxportNet.Omml;
using Mx = DocumentFormat.OpenXml.Math;

namespace DocxportNet.Tests.Omml;

public sealed class DxpOmmlBoxesTests
{
    private const string M = OmmlTestData.MathNamespace;
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";

    [Fact]
    public void PlainBoxGroupsItsArgumentWithoutChangingItsMeaning()
    {
        string omml = Box("", Run("x"));

        foreach (DxpOmmlOutputFormat format in Enum.GetValues<DxpOmmlOutputFormat>())
        {
            DxpOmmlConversionResult result = DxpOmmlConverter.Convert(omml, format);
            Assert.Empty(result.Diagnostics);
            if (format != DxpOmmlOutputFormat.MathMl) Assert.Equal("x", result.Output);
        }
        XElement group = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mrow")
            .Single(row => row.Attribute("data-omml-operator-emulator") != null);
        Assert.Equal("false", (string?)group.Attribute("data-omml-operator-emulator"));
        Assert.Equal("false", (string?)group.Attribute("data-omml-no-break"));
        Assert.Equal("false", (string?)group.Attribute("data-omml-differential"));
    }

    [Fact]
    public void AppliesEveryBoxPropertyAndManualBreakAlignmentIndex()
    {
        string properties = "<m:opEmu/><m:noBreak/><m:diff/><m:brk m:alnAt=\"3\"/><m:aln/><m:ctrlPr/>";
        string omml = Box(properties, Run("="));
        DxpOmmlConversionResult mathMlResult = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.MathMl);
        XElement math = XElement.Parse(mathMlResult.Output);

        XElement lineBreak = math.Descendants(MathMl + "mspace").Single(e => (string?)e.Attribute("linebreak") == "newline");
        Assert.Equal("3", (string?)lineBreak.Attribute("data-omml-align-at"));
        Assert.Single(math.Descendants(MathMl + "malignmark"));
        Assert.Single(math.Descendants(MathMl + "mspace"), e => (string?)e.Attribute("width") == "0.1667em");
        XElement op = math.Descendants(MathMl + "mo").Single(e => e.Value == "=");
        Assert.Equal("true", (string?)op.Attribute("data-omml-operator-emulator"));
        Assert.Equal("true", (string?)op.Attribute("data-omml-no-break"));
        Assert.Equal("true", (string?)op.Attribute("data-omml-differential"));
        AssertApproximation(mathMlResult, "m:box");

        DxpOmmlConversionResult latex = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Latex);
        Assert.Equal("\\begin{aligned}\\\\&\\,\\nobreak{\\mathop{=}}\\nobreak\\end{aligned}", latex.Output);
        AssertApproximation(latex, "m:box");
        Assert.Equal("\n& =", DxpOmmlConverter.ToUnicodeMath(omml));
        Assert.Equal("\n=", DxpOmmlConverter.ToText(omml));
    }

    [Fact]
    public void BoxPropertyDefaultsAndDuplicatesAreDeterministic()
    {
        string properties = "<m:opEmu m:val=\"0\"/><m:opEmu/><m:noBreak m:val=\"off\"/><m:diff m:val=\"false\"/><m:brk/><m:aln m:val=\"0\"/>";
        string omml = Box(properties, Run("+"));
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml));

        Assert.Equal("0", (string?)math.Descendants(MathMl + "mspace").Single(e => (string?)e.Attribute("linebreak") == "newline")
            .Attribute("data-omml-align-at"));
        Assert.Empty(math.Descendants(MathMl + "malignmark"));
        XElement group = math.Descendants(MathMl + "mrow").Single(e => e.Attribute("data-omml-operator-emulator") != null);
        Assert.Equal("false", (string?)group.Attribute("data-omml-operator-emulator"));
    }

    [Fact]
    public void BorderBoxDefaultsToFourSides()
    {
        string omml = BorderBox("", Run("x"));
        XElement enclosure = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "menclose").Single();

        Assert.Equal("top bottom left right", (string?)enclosure.Attribute("notation"));
        Assert.Equal(@"\boxed{x}", DxpOmmlConverter.ToLatex(omml));
        Assert.Equal("▭(x)", DxpOmmlConverter.ToUnicodeMath(omml));
        DxpOmmlConversionResult text = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Text);
        Assert.Equal("enclose[top bottom left right](x)", text.Output);
        AssertApproximation(text, "m:borderBox");
    }

    [Theory]
    [InlineData("hideTop", "bottom left right", 14)]
    [InlineData("hideBot", "top left right", 13)]
    [InlineData("hideLeft", "top bottom right", 11)]
    [InlineData("hideRight", "top bottom left", 7)]
    public void HidesEachBorderIndependently(string property, string notation, int mask)
    {
        string omml = BorderBox($"<m:{property}/>", Run("x"));

        Assert.Equal(notation, (string?)XElement.Parse(DxpOmmlConverter.ToMathMl(omml))
            .Descendants(MathMl + "menclose").Single().Attribute("notation"));
        Assert.Equal($"▭({mask}&x)", DxpOmmlConverter.ToUnicodeMath(omml));
        DxpOmmlConversionResult latex = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Latex);
        Assert.Equal($"\\enclose{{{notation}}}{{x}}", latex.Output);
        AssertApproximation(latex, "m:borderBox");
    }

    [Theory]
    [InlineData("strikeH", "horizontalstrike", 31)]
    [InlineData("strikeV", "verticalstrike", 47)]
    [InlineData("strikeBLTR", "updiagonalstrike", 143)]
    [InlineData("strikeTLBR", "downdiagonalstrike", 79)]
    public void AppliesEachStrikeIndependently(string property, string notation, int mask)
    {
        string omml = BorderBox($"<m:{property}/>", Run("x"));
        string mathMlNotation = (string?)XElement.Parse(DxpOmmlConverter.ToMathMl(omml))
            .Descendants(MathMl + "menclose").Single().Attribute("notation") ?? string.Empty;

        Assert.Contains(notation, mathMlNotation.Split(' '));
        Assert.Equal($"▭({mask}&x)", DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Fact]
    public void SupportsAllHiddenBordersAndEveryStrikeTogether()
    {
        string hiddenProperties = "<m:hideTop/><m:hideBot/><m:hideLeft/><m:hideRight/>";
        string hidden = BorderBox(hiddenProperties, Run("x"));
        XElement group = XElement.Parse(DxpOmmlConverter.ToMathMl(hidden)).Descendants(MathMl + "mrow")
            .Single(e => (string?)e.Attribute("data-omml-border-box") == "true");
        Assert.Equal("none", (string?)group.Attribute("data-omml-notation"));
        Assert.Equal("▭(0&x)", DxpOmmlConverter.ToUnicodeMath(hidden));
        Assert.Equal(@"\enclose{none}{x}", DxpOmmlConverter.ToLatex(hidden));

        string strikes = BorderBox(hiddenProperties + "<m:strikeH/><m:strikeV/><m:strikeBLTR/><m:strikeTLBR/><m:ctrlPr/>", Run("x"));
        Assert.Equal("horizontalstrike verticalstrike updiagonalstrike downdiagonalstrike",
            (string?)XElement.Parse(DxpOmmlConverter.ToMathMl(strikes)).Descendants(MathMl + "menclose").Single().Attribute("notation"));
        Assert.Equal("▭(240&x)", DxpOmmlConverter.ToUnicodeMath(strikes));
    }

    [Fact]
    public void BorderBoxesPreserveNestedExpressions()
    {
        string fraction = $"<m:f><m:num>{Run("a")}</m:num><m:den>{Run("b")}</m:den></m:f>";
        string omml = BorderBox("", fraction);

        Assert.Equal(@"\boxed{\frac{a}{b}}", DxpOmmlConverter.ToLatex(omml));
        XElement enclosure = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "menclose").Single();
        Assert.Single(enclosure.Descendants(MathMl + "mfrac"));
    }

    [Fact]
    public void PhantomDefaultsToVisibleContent()
    {
        string omml = Phantom("", Run("x"));

        Assert.Equal("x", DxpOmmlConverter.ToLatex(omml));
        Assert.Equal("x", DxpOmmlConverter.ToUnicodeMath(omml));
        Assert.Equal("x", DxpOmmlConverter.ToText(omml));
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml));
        Assert.Empty(math.Descendants(MathMl + "mphantom"));
        Assert.Equal("true", (string?)math.Descendants(MathMl + "mrow").Single(e => e.Attribute("data-omml-show") != null)
            .Attribute("data-omml-show"));
    }

    [Fact]
    public void HiddenPhantomNeverExposesContentInReadableText()
    {
        string omml = Phantom("<m:show m:val=\"off\"/>", Run("secret"));

        Assert.Equal(@"\phantom{secret}", DxpOmmlConverter.ToLatex(omml));
        Assert.Equal("⟡(secret)", DxpOmmlConverter.ToUnicodeMath(omml));
        DxpOmmlConversionResult text = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Text);
        Assert.Equal(string.Empty, text.Output);
        AssertApproximation(text, "m:phant");
        XElement phantom = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mphantom").Single();
        Assert.Equal("secret", phantom.Value);
    }

    [Theory]
    [InlineData("zeroWid", "width", "0", @"\mathrlap{x}", "⬌(x)")]
    [InlineData("zeroAsc", "height", "0", @"\smash[t]{x}", "⬆(x)")]
    [InlineData("zeroDesc", "depth", "0", @"\smash[b]{x}", "⬇(x)")]
    public void AppliesEachVisiblePhantomDimensionIndependently(string property,
        string attribute, string attributeValue, string latex, string unicodeMath)
    {
        string omml = Phantom($"<m:{property}/>", Run("x"));
        XElement padded = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mpadded").Single();

        Assert.Equal(attributeValue, (string?)padded.Attribute(attribute));
        Assert.Equal(latex, DxpOmmlConverter.ToLatex(omml));
        Assert.Equal(unicodeMath, DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Fact]
    public void CombinesHiddenAndZeroDimensionPhantomSemantics()
    {
        string verticalOnly = Phantom("<m:show m:val=\"0\"/><m:zeroWid/>", Run("x"));
        Assert.Equal(@"\vphantom{x}", DxpOmmlConverter.ToLatex(verticalOnly));
        Assert.Equal("⇳(x)", DxpOmmlConverter.ToUnicodeMath(verticalOnly));

        string widthOnly = Phantom("<m:show m:val=\"0\"/><m:zeroAsc/><m:zeroDesc/>", Run("x"));
        Assert.Equal(@"\hphantom{x}", DxpOmmlConverter.ToLatex(widthOnly));
        Assert.Equal("⬄(x)", DxpOmmlConverter.ToUnicodeMath(widthOnly));

        string zeroSize = Phantom("<m:show m:val=\"0\"/><m:zeroWid/><m:zeroAsc/><m:zeroDesc/>", Run("x"));
        XElement padded = XElement.Parse(DxpOmmlConverter.ToMathMl(zeroSize)).Descendants(MathMl + "mpadded").Single();
        Assert.Equal(new[] { "0", "0", "0" }, new[] { (string?)padded.Attribute("width"), (string?)padded.Attribute("height"), (string?)padded.Attribute("depth") });
        Assert.Equal("⬌(⬍(⟡(x)))", DxpOmmlConverter.ToUnicodeMath(zeroSize));
    }

    [Fact]
    public void RetainsPhantomTransparencyAndDiagnosesUnrepresentableSpacingSemantics()
    {
        string omml = Phantom("<m:transp/><m:ctrlPr/>", Run("x"));

        foreach (DxpOmmlOutputFormat format in Enum.GetValues<DxpOmmlOutputFormat>())
        {
            DxpOmmlConversionResult result = DxpOmmlConverter.Convert(omml, format);
            AssertApproximation(result, "m:phant");
            if (format == DxpOmmlOutputFormat.MathMl)
                Assert.Equal("true", (string?)XElement.Parse(result.Output).Descendants()
                    .Single(e => e.Attribute("data-omml-transparent") != null).Attribute("data-omml-transparent"));
        }
    }

    [Theory]
    [InlineData("120.omml", "m:borderBox")]
    [InlineData("121.omml", "m:borderBox")]
    [InlineData("line_break/line-break-057.omml", "m:phant")]
    public void ConvertsApplicableUpstreamFixturesWithoutTargetFallback(string fixture, string elementName)
    {
        string omml = File.ReadAllText(Path.Combine(OmmlTestData.UpstreamRoot,
            fixture.Replace('/', Path.DirectorySeparatorChar)));
        foreach (DxpOmmlOutputFormat format in Enum.GetValues<DxpOmmlOutputFormat>())
            Assert.DoesNotContain(DxpOmmlConverter.Convert(omml, format).Diagnostics,
                diagnostic => diagnostic.Code == "OMML001" && diagnostic.ElementName == elementName);
    }

    [Fact]
    public void ConvertsSdkBoxBorderBoxAndPhantomWithoutSerialization()
    {
        Mx.OfficeMath box = new(new Mx.Box(new Mx.Base(new Mx.Run(new Mx.Text("a")))));
        Mx.OfficeMath border = new(new Mx.BorderBox(new Mx.Base(new Mx.Run(new Mx.Text("b")))));
        Mx.OfficeMath phantom = new(new Mx.Phantom(new Mx.Base(new Mx.Run(new Mx.Text("c")))));

        Assert.Equal("a", DxpOmmlConverter.Convert(box, DxpOmmlOutputFormat.UnicodeMath).Output);
        Assert.Equal("▭(b)", DxpOmmlConverter.Convert(border, DxpOmmlOutputFormat.UnicodeMath).Output);
        Assert.Equal("c", DxpOmmlConverter.Convert(phantom, DxpOmmlOutputFormat.UnicodeMath).Output);
    }

    private static void AssertApproximation(DxpOmmlConversionResult result, string elementName)
    {
        DxpOmmlDiagnostic diagnostic = Assert.Single(result.Diagnostics,
            diagnostic => diagnostic.Code == "OMML002" && diagnostic.ElementName == elementName);
        Assert.Equal(DxpOmmlDiagnosticSeverity.Warning, diagnostic.Severity);
    }

    private static string Box(string properties, string argument) => Structure("box", "boxPr", properties, argument);
    private static string BorderBox(string properties, string argument) => Structure("borderBox", "borderBoxPr", properties, argument);
    private static string Phantom(string properties, string argument) => Structure("phant", "phantPr", properties, argument);
    private static string Structure(string name, string propertyName, string properties, string argument) =>
        Inline($"<m:{name}>{(properties.Length == 0 ? string.Empty : $"<m:{propertyName}>{properties}</m:{propertyName}>")}<m:e>{argument}</m:e></m:{name}>");
    private static string Run(string text) => $"<m:r><m:t>{text}</m:t></m:r>";
    private static string Inline(string content) => $"<m:oMath xmlns:m=\"{M}\">{content}</m:oMath>";
}
