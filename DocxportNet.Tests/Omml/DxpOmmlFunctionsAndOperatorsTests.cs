using System.Xml.Linq;
using DocxportNet.Omml;
using Mx = DocumentFormat.OpenXml.Math;

namespace DocxportNet.Tests.Omml;

public sealed class DxpOmmlFunctionsAndOperatorsTests
{
    private const string M = OmmlTestData.MathNamespace;
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";

    [Theory]
    [InlineData("sin", "sin")]
    [InlineData("cosh", "cosh")]
    [InlineData("log", "log")]
    [InlineData("arctan", "arctan")]
    [InlineData("det", "det")]
    [InlineData("lim", "lim")]
    public void MapsOnlyKnownSimpleFunctionNames(string name, string command)
    {
        string omml = Function(Run(name), Run("x"));

        Assert.Equal($"\\{command}{{x}}", DxpOmmlConverter.ToLatex(omml));
        Assert.Equal($"{name}⁡x", DxpOmmlConverter.ToUnicodeMath(omml));
        Assert.Equal($"{name}(x)", DxpOmmlConverter.ToText(omml));
        XElement row = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mrow")
            .Single(e => e.Elements(MathMl + "mo").Any(o => o.Value == "⁡"));
        Assert.Equal(new[] { name, "⁡", "x" }, row.Descendants().Where(e => !e.HasElements).Select(e => e.Value));
    }

    [Fact]
    public void PreservesArbitraryAndStructuredFunctionNamesWithoutGuessing()
    {
        Assert.Equal(@"\operatorname{velocity}{t}",
            DxpOmmlConverter.ToLatex(Function(Run("velocity"), Run("t"))));

        string fractionName = $"<m:f><m:num>{Run("a")}</m:num><m:den>{Run("b")}</m:den></m:f>";
        Assert.Equal(@"\mathop{\frac{a}{b}}{x}",
            DxpOmmlConverter.ToLatex(Function(fractionName, Run("x"))));

        string styledName = "<m:r><m:rPr><m:sty m:val=\"b\"/></m:rPr><m:t>F</m:t></m:r>";
        Assert.Equal(@"\mathop{\mathbf{F}}{x}",
            DxpOmmlConverter.ToLatex(Function(styledName, Run("x"))));
    }

    [Fact]
    public void FunctionApplicationSurvivesEmptyAndDelimitedArgumentsAndControlProperties()
    {
        DxpOmmlConversionOptions throwing = new() { FallbackPolicy = DxpOmmlFallbackPolicy.Throw };
        string empty = Function(Run("sin"), string.Empty, controlProperties: true);
        string delimited = Function(Run("f"), $"<m:d><m:e>{Run("x")}</m:e></m:d>");

        Assert.Equal(@"\sin{}", DxpOmmlConverter.ToLatex(empty, throwing));
        Assert.Equal(@"\operatorname{f}{\left(x\right)}", DxpOmmlConverter.ToLatex(delimited));
        Assert.Equal("f((x))", DxpOmmlConverter.ToText(delimited));
    }

    [Theory]
    [InlineData("limLow", "munder", @"{\lim}_{n→∞}", "lim_(n→∞)", "lim with lower limit n→∞")]
    [InlineData("limUpp", "mover", @"{x}^{2}", "x^(2)", "x with upper limit 2")]
    public void RendersStandaloneLimits(string element, string mathMlElement,
        string latex, string unicodeMath, string text)
    {
        string @base = element == "limLow" ? Run("lim") : Run("x");
        string value = element == "limLow" ? Run("n→∞") : Run("2");
        string omml = Limit(element, @base, value, controlProperties: true);
        DxpOmmlConversionOptions throwing = new() { FallbackPolicy = DxpOmmlFallbackPolicy.Throw };

        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(omml, throwing)).Descendants(MathMl + mathMlElement));
        Assert.Equal(latex, DxpOmmlConverter.ToLatex(omml, throwing));
        Assert.Equal(unicodeMath, DxpOmmlConverter.ToUnicodeMath(omml, throwing));
        Assert.Equal(text, DxpOmmlConverter.ToText(omml, throwing));
    }

    [Fact]
    public void LimitsPreserveNestedFunctionsAccentsAndScripts()
    {
        string accentedBase = $"<m:acc><m:e>{Function(Run("sin"), Run("x"), inline: true)}</m:e></m:acc>";
        string scriptedLimit = $"<m:sSup><m:e>{Run("n")}</m:e><m:sup>{Run("2")}</m:sup></m:sSup>";
        string omml = Limit("limLow", accentedBase, scriptedLimit);

        Assert.Equal(@"{\hat{\sin{x}}}_{n^{2}}", DxpOmmlConverter.ToLatex(omml));
        XElement math = XElement.Parse(DxpOmmlConverter.ToMathMl(omml));
        Assert.Single(math.Descendants(MathMl + "munder"));
        Assert.Single(math.Descendants(MathMl + "mover"), e => (string?)e.Attribute("accent") == "true");
        Assert.Single(math.Descendants(MathMl + "msup"));
    }

    [Theory]
    [InlineData("∑", @"\sum")]
    [InlineData("∏", @"\prod")]
    [InlineData("∐", @"\coprod")]
    [InlineData("∫", @"\int")]
    [InlineData("∬", @"\iint")]
    [InlineData("∭", @"\iiint")]
    [InlineData("∮", @"\oint")]
    [InlineData("∯", @"\oiint")]
    [InlineData("∰", @"\oiiint")]
    [InlineData("⋂", @"\bigcap")]
    [InlineData("⋃", @"\bigcup")]
    [InlineData("⋀", @"\bigwedge")]
    [InlineData("⋁", @"\bigvee")]
    public void MapsStandardNaryOperators(string character, string command)
    {
        string omml = Nary(character, "undOvr", Run("i=1"), Run("n"), Run("x"));
        Assert.Equal($"{command}\\limits_{{i=1}}^{{n}}\\,x", DxpOmmlConverter.ToLatex(omml));
        Assert.Equal($"{character}_(i=1)^(n)▒〖x〗", DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Fact]
    public void SupportsArbitraryNaryCharacterAndGrowth()
    {
        string omml = Nary("★", "subSup", Run("a"), Run("b"), Run("x"), grow: false);
        Assert.Equal(@"\mathop{\text{★}}\nolimits_{a}^{b}\,x", DxpOmmlConverter.ToLatex(omml));
        XElement op = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mo").Single(e => e.Value == "★");
        Assert.Equal("true", (string?)op.Attribute("largeop"));
        Assert.Equal("false", (string?)op.Attribute("stretchy"));
        Assert.Equal("msubsup", op.Parent?.Name.LocalName);
    }

    [Fact]
    public void NaryGrowWithoutAValueUsesTheEnabledDefault()
    {
        string omml = Inline($"<m:nary><m:naryPr><m:chr m:val=\"∑\"/><m:grow/></m:naryPr><m:sub>{Run("i")}</m:sub><m:sup>{Run("n")}</m:sup><m:e>{Run("x")}</m:e></m:nary>");
        XElement op = XElement.Parse(DxpOmmlConverter.ToMathMl(omml)).Descendants(MathMl + "mo").Single(e => e.Value == "∑");
        Assert.Equal("true", (string?)op.Attribute("stretchy"));
    }

    [Fact]
    public void AppliesHiddenLimitsIncludingPropertiesWithoutValues()
    {
        string omml = Nary("∑", "undOvr", Run("i"), Run("n"), Run("x"),
            extraProperties: "<m:subHide/><m:supHide m:val=\"1\"/><m:ctrlPr/>");
        DxpOmmlConversionOptions throwing = new() { FallbackPolicy = DxpOmmlFallbackPolicy.Throw };

        Assert.Equal(@"\sum\limits\,x", DxpOmmlConverter.ToLatex(omml, throwing));
        Assert.Equal("∑▒〖x〗", DxpOmmlConverter.ToUnicodeMath(omml, throwing));
        Assert.Equal("∑ of x", DxpOmmlConverter.ToText(omml, throwing));
        Assert.DoesNotContain(XElement.Parse(DxpOmmlConverter.ToMathMl(omml, throwing)).Descendants(),
            e => e.Name == MathMl + "munder" || e.Name == MathMl + "mover" || e.Name == MathMl + "munderover");
    }

    [Fact]
    public void UsesIntegralAndNaryDocumentDefaultsWhenLocalLocationIsAbsent()
    {
        string integral = Nary(null, null, Run("0"), Run("1"), Run("x"));
        string sum = Nary("∑", null, Run("0"), Run("1"), Run("x"));
        string displaySum = Nary("∑", null, Run("0"), Run("1"), Run("x"), display: true);

        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(integral)).Descendants(MathMl + "msubsup"));
        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(sum)).Descendants(MathMl + "msubsup"));
        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(displaySum)).Descendants(MathMl + "munderover"));
        Assert.StartsWith(@"\int\nolimits", DxpOmmlConverter.ToLatex(integral), StringComparison.Ordinal);
        Assert.StartsWith(@"\sum\nolimits", DxpOmmlConverter.ToLatex(sum), StringComparison.Ordinal);
        Assert.StartsWith(@"\sum\limits", DxpOmmlConverter.ToLatex(displaySum), StringComparison.Ordinal);

        DxpOmmlConversionOptions reversed = new()
        {
            IntegralLimitLocation = DxpOmmlLimitLocation.UnderOver,
            NaryLimitLocation = DxpOmmlLimitLocation.SubscriptSuperscript,
        };
        string displayIntegral = Nary(null, null, Run("0"), Run("1"), Run("x"), display: true);
        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(displayIntegral, reversed)).Descendants(MathMl + "munderover"));
        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(displaySum, reversed)).Descendants(MathMl + "msubsup"));
    }

    [Fact]
    public void LocalLimitLocationOverridesOptionsAndOrdinaryScriptsRemainOrdinary()
    {
        DxpOmmlConversionOptions options = new()
        {
            Display = true,
            NaryLimitLocation = DxpOmmlLimitLocation.SubscriptSuperscript,
        };
        string nary = Nary("∑", "undOvr", Run("i"), Run("n"), Run("x"));
        string script = Inline($"<m:sSubSup><m:e>{Run("x")}</m:e><m:sub>{Run("i")}</m:sub><m:sup>{Run("n")}</m:sup></m:sSubSup>");

        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(nary, options)).Descendants(MathMl + "munderover"));
        Assert.Single(XElement.Parse(DxpOmmlConverter.ToMathMl(script, options)).Descendants(MathMl + "msubsup"));
    }

    [Fact]
    public void MissingArgumentsRecoverAsEmptyAndDuplicatePropertiesUseTheFirstValue()
    {
        string function = Inline("<m:func/>");
        string limit = Inline("<m:limLow/>");
        string nary = Inline($"<m:nary><m:naryPr><m:chr m:val=\"∑\"/><m:chr m:val=\"∏\"/><m:subHide/><m:supHide/></m:naryPr><m:e>{Run("x")}</m:e></m:nary>");

        Assert.Equal(@"\operatorname{}{}", DxpOmmlConverter.ToLatex(function));
        Assert.Equal("{}_{}", DxpOmmlConverter.ToLatex(limit));
        Assert.Equal(@"\sum\nolimits\,x", DxpOmmlConverter.ToLatex(nary));
    }

    [Theory]
    [InlineData("013.omml", "∫▒〖1〗")]
    [InlineData("014.omml", "∫_(2)^(1)▒〖3〗")]
    [InlineData("016.omml", "∬▒〖1〗")]
    [InlineData("035.omml", "∑_(3)^(1)▒〖2〗")]
    [InlineData("073.omml", "sin⁡x")]
    [InlineData("085.omml", "sinh⁡x")]
    public void MatchesRepresentativePinnedFixtureUnicodeMath(string fixture, string expected)
    {
        string omml = File.ReadAllText(Path.Combine(OmmlTestData.UpstreamRoot, fixture));
        Assert.Equal(expected, DxpOmmlConverter.ToUnicodeMath(omml));
    }

    [Fact]
    public void ConvertsEveryApplicablePinnedFixtureWithoutFallback()
    {
        string[] markers = ["<m:func>", "<m:limLow>", "<m:limUpp>", "<m:nary>"];
        string[] fixtures = OmmlTestData.UpstreamFixtures()
            .Where(path => !path.Contains($"{Path.DirectorySeparatorChar}line_break{Path.DirectorySeparatorChar}", StringComparison.Ordinal))
            .Where(path => markers.Any(marker => File.ReadAllText(path).Contains(marker, StringComparison.Ordinal)))
            .ToArray();

        Assert.NotEmpty(fixtures);
        foreach (string fixture in fixtures)
        {
            string omml = File.ReadAllText(fixture);
            foreach (DxpOmmlOutputFormat format in Enum.GetValues<DxpOmmlOutputFormat>())
            {
                DxpOmmlConversionResult result = DxpOmmlConverter.Convert(omml, format);
                Assert.NotEmpty(result.Output);
                Assert.DoesNotContain(result.Diagnostics,
                    diagnostic => diagnostic.ElementName is "m:func" or "m:limLow" or "m:limUpp" or "m:nary");
            }
        }
    }

    [Fact]
    public void ConvertsSdkFunctionWithoutSerialization()
    {
        Mx.MathFunction function = new(
            new Mx.FunctionName(new Mx.Run(new Mx.Text("sin"))),
            new Mx.Base(new Mx.Run(new Mx.Text("x"))));
        Mx.OfficeMath math = new(function);

        Assert.Equal(@"\sin{x}", DxpOmmlConverter.Convert(math, DxpOmmlOutputFormat.Latex).Output);
    }

    private static string Function(string name, string argument, bool controlProperties = false, bool inline = false)
    {
        string value = $"<m:func>{(controlProperties ? "<m:funcPr><m:ctrlPr/></m:funcPr>" : string.Empty)}<m:fName>{name}</m:fName><m:e>{argument}</m:e></m:func>";
        return inline ? value : Inline(value);
    }

    private static string Limit(string element, string @base, string limit, bool controlProperties = false) =>
        Inline($"<m:{element}>{(controlProperties ? $"<m:{element}Pr><m:ctrlPr/></m:{element}Pr>" : string.Empty)}<m:e>{@base}</m:e><m:lim>{limit}</m:lim></m:{element}>");

    private static string Nary(string? character, string? location, string subscript,
        string superscript, string argument, bool grow = true, string extraProperties = "", bool display = false)
    {
        string characterProperty = character == null ? string.Empty : $"<m:chr m:val=\"{character}\"/>";
        string locationProperty = location == null ? string.Empty : $"<m:limLoc m:val=\"{location}\"/>";
        string nary = $"<m:nary><m:naryPr>{characterProperty}{locationProperty}<m:grow m:val=\"{(grow ? "1" : "0")}\"/>{extraProperties}</m:naryPr><m:sub>{subscript}</m:sub><m:sup>{superscript}</m:sup><m:e>{argument}</m:e></m:nary>";
        return display ? $"<m:oMathPara xmlns:m=\"{M}\"><m:oMath>{nary}</m:oMath></m:oMathPara>" : Inline(nary);
    }

    private static string Inline(string content) => $"<m:oMath xmlns:m=\"{M}\">{content}</m:oMath>";
    private static string Run(string text) => $"<m:r><m:t>{text}</m:t></m:r>";
}
