using System.Security.Cryptography;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;

namespace DocxportNet.Tests.Omml;

public sealed class OmmlTestHarnessTests
{
    [Fact]
    public void PinnedUpstreamCorpusHasExpectedInventory()
    {
        IReadOnlyList<string> fixtures = OmmlTestData.UpstreamFixtures();

        Assert.Equal(279, fixtures.Count);
        Assert.Equal(90, fixtures.Count(path => Path.GetDirectoryName(path)!
            .EndsWith("line_break", StringComparison.Ordinal)));
        Assert.All(fixtures, path => Assert.NotEmpty(File.ReadAllText(path)));
    }

    [Fact]
    public void PinnedUpstreamCorpusMatchesRecordedHashes()
    {
        string manifestPath = Path.Combine(OmmlTestData.UpstreamRoot, "corpus-manifest.json");
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(manifestPath));
        JsonElement files = manifest.RootElement.GetProperty("files");

        Assert.Equal(279, files.GetArrayLength());
        foreach (JsonElement entry in files.EnumerateArray())
        {
            string relativePath = entry.GetProperty("path").GetString()!;
            string expectedHash = entry.GetProperty("sha256").GetString()!;
            string actualHash = Convert.ToHexStringLower(SHA256.HashData(
                File.ReadAllBytes(Path.Combine(OmmlTestData.UpstreamRoot, relativePath))));
            Assert.Equal(expectedHash, actualHash);
        }
    }

    [Fact]
    public void CanonicalizationIgnoresPrefixesAttributeOrderAndFormatting()
    {
        const string first = """
            <m:math xmlns:m="urn:math" xmlns:x="urn:x" x:b="2" a="1">
              <m:run>value</m:run>
            </m:math>
            """;
        const string second = "<q:math a=\"1\" p:b=\"2\" xmlns:p=\"urn:x\" xmlns:q=\"urn:math\"><q:run>value</q:run></q:math>";

        Assert.Equal(XmlCanonicalizer.Canonicalize(first), XmlCanonicalizer.Canonicalize(second));
    }

    [Fact]
    public void CanonicalizationPreservesMeaningfulTextWhitespace()
    {
        const string first = "<math><text>a b</text></math>";
        const string second = "<math><text>ab</text></math>";

        Assert.NotEqual(XmlCanonicalizer.Canonicalize(first), XmlCanonicalizer.Canonicalize(second));
        Assert.NotEqual(
            XmlCanonicalizer.Canonicalize("<math><text> </text></math>"),
            XmlCanonicalizer.Canonicalize("<math><text /></math>"));
    }

    [Fact]
    public void BuildersComposeFocusedNestedInlineAndDisplayExpressions()
    {
        XElement fraction = OmmlTestData.Fraction(OmmlTestData.Run("1"), OmmlTestData.Run("2"));
        XElement inline = OmmlTestData.Inline(fraction);
        XElement display = OmmlTestData.Display(new XElement(inline));
        XNamespace m = OmmlTestData.MathNamespace;

        Assert.Equal(m + "oMathPara", display.Name);
        Assert.NotNull(display.Descendants(m + "f").Single());
        Assert.Equal(["1", "2"], display.Descendants(m + "t").Select(node => node.Value));
    }

    [Theory]
    [MemberData(nameof(MalformedFragments))]
    public void MalformedCorpusCasesAreRejectedBySecureCanonicalizer(string _, string xml)
    {
        Assert.ThrowsAny<XmlException>(() => XmlCanonicalizer.Canonicalize(xml));
    }

    public static IEnumerable<object[]> MalformedFragments() => OmmlTestData.MalformedFragments();
}
