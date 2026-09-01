using System.Xml.Linq;
using DocxportNet.Omml;
using Mx = DocumentFormat.OpenXml.Math;

namespace DocxportNet.Tests.Omml;

public sealed class DxpOmmlDecorationsTests
{
    private const string M = OmmlTestData.MathNamespace;
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";

    [Fact]
    public void DelimiterDefaultsToParenthesesAndDefaultSeparator()
    {
        string omml = Inline($"<m:d><m:e>{Run("a")}</m:e><m:e>{Run("b")}</m:e></m:d>");
        XElement row = DelimiterRow(omml);

        Assert.Equal(new[] { "(", "|", ")" }, row.Elements(MathMl + "mo").Select(e => e.Value));
        Assert.Equal(@"\left(a|b\right)", DxpOmmlConverter.ToLatex(omml));
        Assert.Equal("(a|b)", DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Fact]
    public void ExplicitEmptyBoundariesAndSeparatorRemainEmpty()
    {
        string omml = Inline($"<m:d><m:dPr><m:begChr m:val=\"\"/><m:sepChr m:val=\"\"/><m:endChr m:val=\"\"/></m:dPr><m:e>{Run("a")}</m:e><m:e>{Run("b")}</m:e></m:d>");
        XElement row = DelimiterRow(omml);

        Assert.Empty(row.Elements(MathMl + "mo"));
        Assert.Equal(@"\left.ab\right.", DxpOmmlConverter.ToLatex(omml));
        Assert.Equal("ab", DxpOmmlConverter.ToText(omml));
    }

    [Fact]
    public void ExplicitSeparatorIsEscapedForLatex()
    {
        string omml = Delimiter("(", ")", "_", $"{Run("a")}</m:e><m:e>{Run("b")}", grow: false);
        Assert.Equal(@"(a\_b)", DxpOmmlConverter.ToLatex(omml));
    }

    [Theory]
    [InlineData("[", "]", "[", "]")]
    [InlineData("{", "}", @"\{", @"\}")]
    [InlineData("|", "|", "|", "|")]
    [InlineData("‖", "‖", @"\Vert", @"\Vert")]
    [InlineData("⟨", "⟩", @"\langle", @"\rangle")]
    [InlineData("⌊", "⌋", @"\lfloor", @"\rfloor")]
    [InlineData("⌈", "⌉", @"\lceil", @"\rceil")]
    [InlineData("⟦", "⟧", @"\lbbrack", @"\rbbrack")]
    [InlineData("⦅", "⦆", "⦅", "⦆")]
    public void MapsCommonAndArbitraryUnicodeDelimiters(string begin, string end,
        string latexBegin, string latexEnd)
    {
        string omml = Delimiter(begin, end, string.Empty, Run("x"), grow: false);
        Assert.Equal($"{latexBegin}x{latexEnd}", DxpOmmlConverter.ToLatex(omml));
        Assert.Equal($"{begin}x{end}", DxpOmmlConverter.ToText(omml));
    }

    [Fact]
    public void AppliesGrowShapeAndControlProperties()
    {
        string omml = Inline($"<m:d><m:dPr><m:grow m:val=\"off\"/><m:shp m:val=\"match\"/><m:ctrlPr/></m:dPr><m:e>{Run("x")}</m:e></m:d>");
        DxpOmmlConversionOptions throwing = new() { FallbackPolicy = DxpOmmlFallbackPolicy.Throw };
        XElement row = DelimiterRow(omml, throwing);

        Assert.Equal("match", (string?)row.Attribute("data-omml-shape"));
        Assert.All(row.Elements(MathMl + "mo").Where(e => (string?)e.Attribute("fence") == "true"),
            fence => Assert.Equal("false", (string?)fence.Attribute("stretchy")));
        Assert.Equal("(x)", DxpOmmlConverter.ToLatex(omml, throwing));
    }

    [Fact]
    public void GrowAndShapeElementsWithoutValuesUseTheirSchemaDefaults()
    {
        string omml = Inline($"<m:d><m:dPr><m:grow/><m:shp/></m:dPr><m:e>{Run("x")}</m:e></m:d>");
        XElement row = DelimiterRow(omml);

        Assert.Equal("centered", (string?)row.Attribute("data-omml-shape"));
        Assert.All(row.Elements(MathMl + "mo").Where(e => (string?)e.Attribute("fence") == "true"),
            fence => Assert.Equal("true", (string?)fence.Attribute("stretchy")));
        Assert.Equal(@"\left(x\right)", DxpOmmlConverter.ToLatex(omml));
    }

    [Fact]
    public void DelimitersRemainAroundPendingMatrixAndEquationArrayStructures()
    {
        string matrix = $"<m:m><m:mr><m:e>{Run("1")}</m:e><m:e>{Run("2")}</m:e></m:mr></m:m>";
        string equationArray = $"<m:eqArr><m:e>{Run("a")}</m:e><m:e>{Run("b")}</m:e></m:eqArr>";
        string omml = Inline($"<m:d><m:e>{matrix}</m:e><m:e>{equationArray}</m:e></m:d>");

        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(omml, DxpOmmlOutputFormat.Text);
        Assert.Equal("(12|ab)", result.Output);
        Assert.Equal(new[] { "m:m", "m:eqArr" }, result.Diagnostics.Select(d => d.ElementName));
    }

    [Theory]
    [InlineData("́", "acute")]
    [InlineData("̀", "grave")]
    [InlineData("̂", "hat")]
    [InlineData("̌", "check")]
    [InlineData("̃", "tilde")]
    [InlineData("̄", "bar")]
    [InlineData("̆", "breve")]
    [InlineData("̇", "dot")]
    [InlineData("̈", "ddot")]
    [InlineData("⃗", "vec")]
    public void MapsCommonAccents(string character, string command)
    {
        string omml = Accent(character, Run("x"));
        Assert.Equal($"\\{command}{{x}}", DxpOmmlConverter.ToLatex(omml));
        XElement mover = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mover").Single();
        Assert.Equal("true", (string?)mover.Attribute("accent"));
        Assert.NotEmpty(Assert.IsType<string>(mover.Element(MathMl + "mo")?.Value));
    }

    [Fact]
    public void DefaultsToHatAndSupportsArbitraryAccentAroundStructuredContent()
    {
        string fraction = $"<m:f><m:num>{Run("1")}</m:num><m:den>{Run("2")}</m:den></m:f>";
        Assert.Equal(@"\hat{x}", DxpOmmlConverter.ToLatex(Inline($"<m:acc><m:e>{Run("x")}</m:e></m:acc>")));
        string arbitrary = Accent("☀", fraction);
        Assert.Equal(@"\overset{\text{☀}}{\frac{1}{2}}", DxpOmmlConverter.ToLatex(arbitrary));
        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(arbitrary)).Descendants(MathMl + "mfrac"));
    }

    [Fact]
    public void PreservesNestedDecorationsAndDecorationAroundPendingMatrices()
    {
        string nested = Inline($"<m:bar><m:e><m:acc><m:e>{Run("x")}</m:e></m:acc></m:e></m:bar>");
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(nested));
        Assert.Single(math.Descendants(MathMl + "munder"), e => (string?)e.Attribute("accentunder") == "false");
        Assert.Single(math.Descendants(MathMl + "mover"), e => (string?)e.Attribute("accent") == "true");

        string matrix = Inline($"<m:acc><m:e><m:m><m:mr><m:e>{Run("1")}</m:e><m:e>{Run("2")}</m:e></m:mr></m:m></m:e></m:acc>");
        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(matrix, DxpOmmlOutputFormat.MathMl);
        Assert.Single(XElement.Parse(result.Output).Descendants(MathMl + "mover"));
        Assert.Equal("m:m", Assert.Single(result.Diagnostics).ElementName);
    }

    [Theory]
    [InlineData("top", "mover", "accent", "false", @"\overline{x}", "(x)̅")]
    [InlineData("bot", "munder", "accentunder", "false", @"\underline{x}", "(x)̲")]
    public void SupportsBarsAboveAndBelow(string position, string elementName,
        string accentAttribute, string accentValue, string latex, string unicodeMath)
    {
        string omml = Inline($"<m:bar><m:barPr><m:pos m:val=\"{position}\"/><m:ctrlPr/></m:barPr><m:e>{Run("x")}</m:e></m:bar>");
        XElement bar = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + elementName).Single();
        Assert.Equal(accentValue, (string?)bar.Attribute(accentAttribute));
        Assert.Equal(latex, DxpOmmlConverter.ToLatex(omml));
        Assert.Equal(unicodeMath, DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Fact]
    public void BarDefaultsBelowWhenPositionIsAbsentOrHasNoValue()
    {
        string absent = Inline($"<m:bar><m:e>{Run("x")}</m:e></m:bar>");
        string missingValue = Inline($"<m:bar><m:barPr><m:pos/></m:barPr><m:e>{Run("x")}</m:e></m:bar>");

        Assert.Equal(@"\underline{x}", DxpOmmlConverter.ToLatex(absent));
        Assert.Equal(@"\underline{x}", DxpOmmlConverter.ToLatex(missingValue));
    }

    [Theory]
    [InlineData("top", "top", "⏞", "mover", @"\overbrace{x}")]
    [InlineData("bot", "bot", "⏟", "munder", @"\underbrace{x}")]
    [InlineData("top", "bot", "⏜", "mover", @"\overparen{x}")]
    [InlineData("bot", "top", "⏝", "munder", @"\underparen{x}")]
    public void SupportsGroupCharacterPositionAndVerticalJustification(string position,
        string verticalJustification, string character, string elementName, string latex)
    {
        string omml = GroupCharacter(character, position, verticalJustification, Run("x"));
        XElement group = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + elementName).Single();
        Assert.Equal(verticalJustification, (string?)group.Attribute("data-omml-vertical-justification"));
        Assert.Equal(latex, DxpOmmlConverter.ToLatex(omml));
    }

    [Fact]
    public void GroupCharacterDefaultsAndMissingValueDefaultsAreDistinct()
    {
        string absent = Inline($"<m:groupChr><m:e>{Run("x")}</m:e></m:groupChr>");
        string presentWithoutValue = Inline($"<m:groupChr><m:groupChrPr><m:vertJc/></m:groupChrPr><m:e>{Run("x")}</m:e></m:groupChr>");
        XElement absentGroup = XElement.Parse(DxpOmmlConverter.ToMathMl(absent)).Descendants(MathMl + "munder").Single();
        XElement presentGroup = XElement.Parse(DxpOmmlConverter.ToMathMl(presentWithoutValue)).Descendants(MathMl + "munder").Single();

        Assert.Equal("top", (string?)absentGroup.Attribute("data-omml-vertical-justification"));
        Assert.Equal("bot", (string?)presentGroup.Attribute("data-omml-vertical-justification"));
        Assert.Equal("⏟", absentGroup.Element(MathMl + "mo")?.Value);
        Assert.Equal(@"\underbrace{x}", DxpOmmlConverter.ToLatex(absent));
    }

    [Theory]
    [InlineData("050.omml", "(1)")]
    [InlineData("051.omml", "[2]")]
    [InlineData("052.omml", "{3}")]
    [InlineData("054.omml", "⌊5⌋")]
    [InlineData("055.omml", "⌈6⌉")]
    [InlineData("057.omml", "‖8‖")]
    [InlineData("103.omml", "(a)̂")]
    [InlineData("111.omml", "⏞(c)")]
    [InlineData("123.omml", "(under)̅")]
    [InlineData("185.omml", "(under)┴(_)")]
    [InlineData("186.omml", "(under)┬⏟")]
    public void MatchesPinnedFixtureUnicodeMath(string fixture, string expected)
    {
        string omml = File.ReadAllText(Path.Combine(OmmlTestData.UpstreamRoot, fixture));
        Assert.Equal(expected, DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Fact]
    public void ConvertsSdkDelimiterWithoutSerialization()
    {
        Mx.Delimiter delimiter = new(new Mx.Base(new Mx.Run(new Mx.Text("x"))));
        Mx.OfficeMath math = new(delimiter);
        Assert.Equal("(x)", DxpOmmlConverter.Convert(math, DxpOmmlOutputFormat.Text).Output);
    }

    private static XElement DelimiterRow(string omml, DxpOmmlConversionOptions? options = null) =>
        XElement.Parse(DxpOmmlConverter.ToMathMl(omml, options)).Descendants(MathMl + "mrow")
            .Single(e => e.Attribute("data-omml-shape") != null);

    private static string Delimiter(string begin, string end, string separator, string content, bool grow) =>
        Inline($"<m:d><m:dPr><m:begChr m:val=\"{begin}\"/><m:sepChr m:val=\"{separator}\"/><m:endChr m:val=\"{end}\"/><m:grow m:val=\"{(grow ? "on" : "off")}\"/></m:dPr><m:e>{content}</m:e></m:d>");
    private static string Accent(string character, string content) =>
        Inline($"<m:acc><m:accPr><m:chr m:val=\"{character}\"/></m:accPr><m:e>{content}</m:e></m:acc>");
    private static string GroupCharacter(string character, string position, string verticalJustification, string content) =>
        Inline($"<m:groupChr><m:groupChrPr><m:chr m:val=\"{character}\"/><m:pos m:val=\"{position}\"/><m:vertJc m:val=\"{verticalJustification}\"/></m:groupChrPr><m:e>{content}</m:e></m:groupChr>");
    private static string Inline(string content) => $"<m:oMath xmlns:m=\"{M}\">{content}</m:oMath>";
    private static string Run(string text) => $"<m:r><m:t>{text}</m:t></m:r>";
}
