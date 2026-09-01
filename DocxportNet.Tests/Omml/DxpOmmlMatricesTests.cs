using System.Xml.Linq;
using System.Xml;
using DocxportNet.Omml;
using Mx = DocumentFormat.OpenXml.Math;

namespace DocxportNet.Tests.Omml;

public sealed class DxpOmmlMatricesTests
{
    private const string M = OmmlTestData.MathNamespace;
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";

    [Fact]
    public void RendersRectangularMatrixInEveryOutput()
    {
        string omml = Matrix("", Row("a", "b"), Row("c", "d"));

        Assert.Equal(@"\begin{array}[c]{cc}a & b \\ c & d\end{array}", DxpOmmlConverter.ToLatex(omml));
        Assert.Equal("■(a&b@c&d)", DxpOmmlConverter.ToUnicodeMath(omml));
        Assert.Equal("[[a, b]; [c, d]]", DxpOmmlConverter.ToText(omml));
        XElement table = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mtable").Single();
        Assert.Equal(2, table.Elements(MathMl + "mtr").Count());
        Assert.All(table.Elements(MathMl + "mtr"), row => Assert.Equal(2, row.Elements(MathMl + "mtd").Count()));
    }

    [Fact]
    public void AppliesRepeatedColumnAlignmentAndRetainsRaggedRows()
    {
        string properties = "<m:mcs><m:mc><m:mcPr><m:count m:val=\"2\"/><m:mcJc m:val=\"left\"/></m:mcPr></m:mc>" +
            "<m:mc><m:mcPr><m:count/><m:mcJc m:val=\"right\"/></m:mcPr></m:mc></m:mcs>";
        string omml = Matrix(properties, Row("a", "b", "c"), Row("d"));

        XElement table = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mtable").Single();
        Assert.Equal("left left right", (string?)table.Attribute("columnalign"));
        Assert.Equal(new[] { 3, 1 }, table.Elements(MathMl + "mtr").Select(row => row.Elements(MathMl + "mtd").Count()));
        Assert.Equal(@"\begin{array}[c]{llr}a & b & c \\ d\end{array}", DxpOmmlConverter.ToLatex(omml));
        Assert.Equal("■(a&b&c@d)", DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Fact]
    public void DefaultsUndefinedColumnsToCenterWithoutDroppingCells()
    {
        string columns = "<m:mcs><m:mc><m:mcPr><m:mcJc m:val=\"right\"/></m:mcPr></m:mc></m:mcs>";
        string omml = Matrix(columns, Row("a", "b", "c"));

        XElement table = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mtable").Single();
        Assert.Equal("right center center", (string?)table.Attribute("columnalign"));
        Assert.Equal(@"\begin{array}[c]{rcc}a & b & c\end{array}", DxpOmmlConverter.ToLatex(omml));
    }

    [Fact]
    public void RetainsMatrixLayoutPropertiesAndPlaceholderVisibility()
    {
        string properties = "<m:baseJc m:val=\"bot\"/><m:plcHide/><m:rSpRule m:val=\"3\"/><m:rSp m:val=\"240\"/>" +
            "<m:cSp m:val=\"120\"/><m:cGp m:val=\"360\"/><m:cGpRule m:val=\"4\"/><m:ctrlPr/>";
        string hidden = Matrix(properties, "<m:mr><m:e/></m:mr>");
        XElement hiddenTable = XElement.Parse(DxpOmmlConverter.ToMathMl(hidden)).Descendants(MathMl + "mtable").Single();

        Assert.Equal("bottom", (string?)hiddenTable.Attribute("align"));
        Assert.Equal("true", (string?)hiddenTable.Attribute("data-omml-placeholder-hidden"));
        Assert.Equal("240", (string?)hiddenTable.Attribute("data-omml-row-spacing"));
        Assert.Equal("3", (string?)hiddenTable.Attribute("data-omml-row-spacing-rule"));
        Assert.Equal("120", (string?)hiddenTable.Attribute("data-omml-column-spacing"));
        Assert.Equal("360", (string?)hiddenTable.Attribute("data-omml-column-gap"));
        Assert.Equal("4", (string?)hiddenTable.Attribute("data-omml-column-gap-rule"));
        Assert.Empty(hiddenTable.Descendants(MathMl + "mspace"));
        Assert.Contains(@"\begin{array}[b]", DxpOmmlConverter.ToLatex(hidden), StringComparison.Ordinal);

        string visible = Matrix("", "<m:mr><m:e/></m:mr>");
        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(visible)).Descendants(MathMl + "mspace"),
            space => (string?)space.Attribute("data-omml-placeholder") == "true");
    }

    [Fact]
    public void MatrixCellsSupportNestedExpressions()
    {
        string fraction = $"<m:f><m:num>{Run("a")}</m:num><m:den>{Run("b")}</m:den></m:f>";
        string omml = Matrix("", $"<m:mr><m:e>{fraction}</m:e><m:e>{Run("x")}</m:e></m:mr>");

        Assert.Contains(@"\frac{a}{b}", DxpOmmlConverter.ToLatex(omml), StringComparison.Ordinal);
        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mfrac"));
    }

    [Fact]
    public void EquationArrayUsesGatheredOrAlignedAccordingToMarkers()
    {
        string gathered = EquationArray("", $"<m:e>{Run("x=1")}</m:e>", $"<m:e>{Run("y=2")}</m:e>");
        string aligned = EquationArray("", $"<m:e>{Run("x", true)}{Run("=1")}</m:e>", $"<m:e>{Run("y", true)}{Run("=2")}</m:e>");

        Assert.Equal(@"\begin{gathered}x=1 \\ y=2\end{gathered}", DxpOmmlConverter.ToLatex(gathered));
        Assert.Equal(@"\begin{aligned}&x=1 \\ &y=2\end{aligned}", DxpOmmlConverter.ToLatex(aligned));
        Assert.Equal("█(x=1@y=2)", DxpOmmlConverter.ToUnicodeMath(gathered));
        Assert.Equal("█(&x=1@&y=2)", DxpOmmlConverter.ToUnicodeMath(aligned));
        Assert.Equal("x=1; y=2", DxpOmmlConverter.ToText(gathered));
        Assert.Equal(2, XElement.Parse(DxpOmmlConverter.ToMathMl(aligned)).Descendants(MathMl + "malignmark").Count());
    }

    [Fact]
    public void RetainsEquationArrayPropertiesAndRows()
    {
        string properties = "<m:baseJc m:val=\"top\"/><m:maxDist/><m:objDist m:val=\"true\"/>" +
            "<m:rSpRule m:val=\"2\"/><m:rSp m:val=\"180\"/><m:ctrlPr/>";
        string omml = EquationArray(properties, $"<m:e>{Run("a")}</m:e>", "<m:e/>", $"<m:e>{Run("b")}</m:e>");
        XElement table = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mtable").Single();

        Assert.Equal("top", (string?)table.Attribute("align"));
        Assert.Equal("100%", (string?)table.Attribute("width"));
        Assert.Equal("true", (string?)table.Attribute("data-omml-max-distribution"));
        Assert.Equal("true", (string?)table.Attribute("data-omml-object-distribution"));
        Assert.Equal("180", (string?)table.Attribute("data-omml-row-spacing"));
        Assert.Equal("2", (string?)table.Attribute("data-omml-row-spacing-rule"));
        Assert.Equal(3, table.Elements(MathMl + "mtr").Count());
        Assert.StartsWith(@"\begin{gathered}[t]", DxpOmmlConverter.ToLatex(omml), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("156.omml", "■(a&b)")]
    [InlineData("157.omml", "█(a@b)")]
    [InlineData("163.omml", "■(1&2&3@4&5&6@7&8&9)")]
    public void MatchesPinnedFixtureUnicodeMath(string fixture, string expected)
    {
        string omml = File.ReadAllText(Path.Combine(OmmlTestData.UpstreamRoot, fixture));
        Assert.Equal(expected, DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Theory]
    [InlineData("156.omml")]
    [InlineData("157.omml")]
    [InlineData("158.omml")]
    [InlineData("159.omml")]
    [InlineData("176.omml")]
    [InlineData("177.omml")]
    public void ConvertsRepresentativePinnedFixturesWithoutFallback(string fixture)
    {
        string omml = File.ReadAllText(Path.Combine(OmmlTestData.UpstreamRoot, fixture));
        foreach (DxpOmmlOutputFormat format in Enum.GetValues<DxpOmmlOutputFormat>())
        {
            DxpOmmlConversionResult result = DxpOmmlConverter.Convert(omml, format);
            Assert.NotEmpty(result.Output);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.ElementName is "m:m" or "m:eqArr");
        }
    }

    [Fact]
    public void ConvertsEveryApplicablePinnedFixtureWithoutMatrixFallback()
    {
        string[] fixtures = OmmlTestData.UpstreamFixtures()
            .Where(path => !path.Contains($"{Path.DirectorySeparatorChar}line_break{Path.DirectorySeparatorChar}", StringComparison.Ordinal))
            .Where(path => { string xml = File.ReadAllText(path); return IsWellFormed(xml) && (xml.Contains("<m:m>", StringComparison.Ordinal) || xml.Contains("<m:eqArr>", StringComparison.Ordinal)); })
            .ToArray();
        Assert.NotEmpty(fixtures);
        foreach (string fixture in fixtures)
        foreach (DxpOmmlOutputFormat format in Enum.GetValues<DxpOmmlOutputFormat>())
            Assert.DoesNotContain(DxpOmmlConverter.Convert(File.ReadAllText(fixture), format).Diagnostics,
                diagnostic => diagnostic.ElementName is "m:m" or "m:eqArr");
    }

    [Fact]
    public void ConvertsSdkMatrixWithoutSerialization()
    {
        Mx.Matrix matrix = new(new Mx.MatrixRow(
            new Mx.Base(new Mx.Run(new Mx.Text("a"))),
            new Mx.Base(new Mx.Run(new Mx.Text("b")))));
        Mx.OfficeMath math = new(matrix);

        Assert.Equal("■(a&b)", DxpOmmlConverter.Convert(math, DxpOmmlOutputFormat.UnicodeMath).Output);
    }

    [Fact]
    public void ConvertsSdkEquationArrayWithoutSerialization()
    {
        Mx.EquationArray equationArray = new(
            new Mx.Base(new Mx.Run(new Mx.Text("a"))),
            new Mx.Base(new Mx.Run(new Mx.Text("b"))));
        Mx.OfficeMath math = new(equationArray);

        Assert.Equal("█(a@b)", DxpOmmlConverter.Convert(math, DxpOmmlOutputFormat.UnicodeMath).Output);
    }

    private static string Matrix(string properties, params string[] rows) =>
        Inline($"<m:m>{(properties.Length == 0 ? string.Empty : $"<m:mPr>{properties}</m:mPr>")}{string.Concat(rows)}</m:m>");
    private static string EquationArray(string properties, params string[] rows) =>
        Inline($"<m:eqArr>{(properties.Length == 0 ? string.Empty : $"<m:eqArrPr>{properties}</m:eqArrPr>")}{string.Concat(rows)}</m:eqArr>");
    private static string Row(params string[] cells) => $"<m:mr>{string.Concat(cells.Select(cell => $"<m:e>{Run(cell)}</m:e>"))}</m:mr>";
    private static string Run(string text, bool alignment = false) => $"<m:r>{(alignment ? "<m:rPr><m:aln/></m:rPr>" : string.Empty)}<m:t>{text}</m:t></m:r>";
    private static string Inline(string content) => $"<m:oMath xmlns:m=\"{M}\">{content}</m:oMath>";
    private static bool IsWellFormed(string xml)
    {
        try { XDocument.Parse(xml); return true; }
        catch (XmlException) { return false; }
    }
}
