using System.Xml.Linq;
using DocxportNet.Omml;
using Mx = DocumentFormat.OpenXml.Math;

namespace DocxportNet.Tests.Omml;

public sealed class DxpOmmlStructuresTests
{
    private const string M = OmmlTestData.MathNamespace;
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";

    [Theory]
    [InlineData(null, "mfrac", null, null, @"\frac{a}{b}")]
    [InlineData("bar", "mfrac", null, null, @"\frac{a}{b}")]
    [InlineData("skw", "mfrac", "bevelled", "true", "{a}/{b}")]
    [InlineData("lin", "mrow", null, null, "{a}/{b}")]
    [InlineData("noBar", "mfrac", "linethickness", "0", @"\genfrac{}{}{0pt}{}{a}{b}")]
    public void SupportsEveryFractionType(string? type, string elementName,
        string? attributeName, string? attributeValue, string expectedLatex)
    {
        string property = type == null ? string.Empty : $"<m:fPr><m:type m:val=\"{type}\"/></m:fPr>";
        string omml = Inline($"<m:f>{property}<m:num>{Run("a")}</m:num><m:den>{Run("b")}</m:den></m:f>");

        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml));
        XElement rendered = math.Descendants(MathMl + elementName).First();
        if (attributeName != null) Assert.Equal(attributeValue, (string?)rendered.Attribute(attributeName));
        Assert.Equal(expectedLatex, DxpOmmlConverter.ToLatex(omml));
        Assert.Equal("(a)/(b)", DxpOmmlConverter.ToUnicodeMath(omml));
        Assert.Equal("(a)/(b)", DxpOmmlConverter.ToText(omml));
    }

    [Fact]
    public void FractionsPreserveEmptyAndNestedArgumentsAndControlProperties()
    {
        string inner = $"<m:f><m:num>{Run("1")}</m:num><m:den>{Run("2")}</m:den></m:f>";
        string omml = Inline($"<m:f><m:fPr><m:ctrlPr/></m:fPr><m:num>{inner}</m:num><m:den/></m:f>");
        DxpOmmlConversionOptions throwing = new() { FallbackPolicy = DxpOmmlFallbackPolicy.Throw };

        Assert.Equal(@"\frac{\frac{1}{2}}{}", DxpOmmlConverter.ToLatex(omml, throwing));
        XElement outer = XElement.Parse(DxpOmmlConverter.ToMathMl(omml, throwing)).Descendants(MathMl + "mfrac").First();
        Assert.Equal(2, outer.Elements(MathMl + "mrow").Count());
    }

    [Fact]
    public void SmallFractionsApplyOnlyToInlineMath()
    {
        string fraction = $"<m:f><m:num>{Run("1")}</m:num><m:den>{Run("2")}</m:den></m:f>";
        DxpOmmlConversionOptions options = new() { SmallFractions = true };
        XElement inline = XElement.Parse(DxpOmmlConverter.ToMathMl(Inline(fraction), options));
        XElement display = XElement.Parse(DxpOmmlConverter.ToMathMl($"<m:oMathPara xmlns:m=\"{M}\"><m:oMath>{fraction}</m:oMath></m:oMathPara>", options));

        Assert.Single(inline.Descendants(MathMl + "mstyle"), e => (string?)e.Attribute("scriptlevel") == "1");
        Assert.DoesNotContain(display.Descendants(MathMl + "mstyle"), e => (string?)e.Attribute("scriptlevel") == "1");
    }

    [Theory]
    [InlineData("", false, "msqrt", @"\sqrt{x}", "√(x)")]
    [InlineData("<m:deg><m:r><m:t>3</m:t></m:r></m:deg>", false, "mroot", @"\sqrt[3]{x}", "√(3&x)")]
    [InlineData("<m:deg/>", false, "mroot", @"\sqrt[]{x}", "√(&x)")]
    [InlineData("<m:deg><m:r><m:t>3</m:t></m:r></m:deg>", true, "msqrt", @"\sqrt{x}", "√(x)")]
    public void SupportsMissingVisibleEmptyAndHiddenRadicalDegrees(string degree,
        bool hidden, string mathMlName, string latex, string unicodeMath)
    {
        string properties = hidden ? "<m:radPr><m:degHide/></m:radPr>" : string.Empty;
        string omml = Inline($"<m:rad>{properties}{degree}<m:e>{Run("x")}</m:e></m:rad>");

        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + mathMlName));
        Assert.Equal(latex, DxpOmmlConverter.ToLatex(omml));
        Assert.Equal(unicodeMath, DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Fact]
    public void ExplicitFalseDegreeHideKeepsTheIndex()
    {
        string omml = Inline($"<m:rad><m:radPr><m:degHide m:val=\"off\"/><m:ctrlPr/></m:radPr><m:deg>{Run("3")}</m:deg><m:e>{Run("x")}</m:e></m:rad>");
        DxpOmmlConversionOptions throwing = new() { FallbackPolicy = DxpOmmlFallbackPolicy.Throw };
        Assert.Equal(@"\sqrt[3]{x}", DxpOmmlConverter.ToLatex(omml, throwing));
    }

    [Theory]
    [InlineData("sSub", "msub", @"x_{i}", "x_(i)")]
    [InlineData("sSup", "msup", @"x^{2}", "x^(2)")]
    [InlineData("sSubSup", "msubsup", @"x_{i}^{2}", "x_(i)^(2)")]
    [InlineData("sPre", "mmultiscripts", @"{}_{i}^{2}x", "_(i)^(2) x")]
    public void SupportsEveryOrdinaryScriptForm(string kind, string mathMlName,
        string latex, string unicodeMath)
    {
        string omml = Inline($"<m:{kind}><m:e>{Run("x")}</m:e><m:sub>{Run("i")}</m:sub><m:sup>{Run("2")}</m:sup></m:{kind}>");

        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + mathMlName));
        Assert.Equal(latex, DxpOmmlConverter.ToLatex(omml));
        Assert.Equal(unicodeMath, DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Fact]
    public void ScriptsPreserveEmptyArgumentsStructuredBasesAndAlignmentIntent()
    {
        string fraction = $"<m:f><m:num>{Run("a")}</m:num><m:den>{Run("b")}</m:den></m:f>";
        string omml = Inline($"<m:sSubSup><m:sSubSupPr><m:alnScr/><m:ctrlPr/></m:sSubSupPr><m:e>{fraction}</m:e><m:sub/><m:sup>{Run("2")}</m:sup></m:sSubSup>");
        DxpOmmlConversionOptions throwing = new() { FallbackPolicy = DxpOmmlFallbackPolicy.Throw };

        XElement script = XElement.Parse(DxpOmmlConverter.ToMathMl(omml, throwing)).Descendants(MathMl + "msubsup").Single();
        Assert.Equal("true", (string?)script.Attribute("data-omml-align-scripts"));
        Assert.Single(script.Descendants(MathMl + "mfrac"));
        Assert.Equal(@"\frac{a}{b}_{}^{2}", DxpOmmlConverter.ToLatex(omml, throwing));
    }

    [Fact]
    public void PreservesEmptyScriptBaseAndDeeplyNestedRadicalDegree()
    {
        string fraction = $"<m:f><m:num>{Run("1")}</m:num><m:den>{Run("2")}</m:den></m:f>";
        string radical = $"<m:rad><m:deg>{fraction}</m:deg><m:e>{Run("x")}</m:e></m:rad>";
        string omml = Inline($"<m:sSubSup><m:e/><m:sub>{radical}</m:sub><m:sup/></m:sSubSup>");

        Assert.Equal(@"{}_{\sqrt[\frac{1}{2}]{x}}^{}", DxpOmmlConverter.ToLatex(omml));
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml));
        Assert.Single(math.Descendants(MathMl + "mroot"));
        Assert.Single(math.Descendants(MathMl + "mfrac"));
    }

    [Theory]
    [InlineData("002.omml", @"\frac{1}{2}", "(1)/(2)")]
    [InlineData("005.omml", @"1^{2}", "1^(2)")]
    [InlineData("006.omml", @"1_{2}", "1_(2)")]
    [InlineData("007.omml", @"1_{3}^{2}", "1_(3)^(2)")]
    [InlineData("008.omml", @"{}_{3}^{1}2", "_(3)^(1) 2")]
    [InlineData("010.omml", @"\sqrt[2]{1}", "√(2&1)")]
    public void MatchesPinnedOracleForCoreGoalFourFixtures(string fixture,
        string latex, string unicodeMath)
    {
        string omml = File.ReadAllText(Path.Combine(OmmlTestData.UpstreamRoot, fixture));
        Assert.Equal(latex, DxpOmmlConverter.ToLatex(omml));
        Assert.Equal(unicodeMath, DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Fact]
    public void ConvertsStructuredOpenXmlSdkElementsWithoutSerialization()
    {
        Mx.Fraction fraction = new(
            new Mx.Numerator(new Mx.Run(new Mx.Text("1"))),
            new Mx.Denominator(new Mx.Run(new Mx.Text("2"))));
        Mx.OfficeMath math = new(fraction);

        Assert.Equal(@"\frac{1}{2}", DxpOmmlConverter.Convert(math, DxpOmmlOutputFormat.Latex).Output);
    }

    private static string Inline(string content) => $"<m:oMath xmlns:m=\"{M}\">{content}</m:oMath>";
    private static string Run(string text) => $"<m:r><m:t>{text}</m:t></m:r>";
}
