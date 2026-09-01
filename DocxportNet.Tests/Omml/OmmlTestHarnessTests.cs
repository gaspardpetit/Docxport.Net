using System.Security.Cryptography;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using DocxportNet.Omml;

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
    public void GeneratedOracleManifestMatchesAvailableOutputs()
    {
        string oracleRoot = Path.Combine(OmmlTestData.FixtureRoot, "OracleGenerated");
        using JsonDocument manifest = JsonDocument.Parse(
            File.ReadAllText(Path.Combine(oracleRoot, "manifest.json")));
        JsonElement root = manifest.RootElement;

        Assert.Equal("00c52783877b38f6b8e6e109f1803f96bb34fc62",
            root.GetProperty("oracle").GetProperty("plurimath_commit").GetString());
        Assert.Equal("51d4abe5df58fe33a92df094971c5828c3459ffb",
            root.GetProperty("oracle").GetProperty("omml_commit").GetString());

        JsonElement fixtures = root.GetProperty("fixtures");
        Assert.Equal(279, fixtures.GetArrayLength());
        Assert.Equal(260, fixtures.EnumerateArray().Count(entry =>
            entry.GetProperty("status").GetString() == "converted"));
        Assert.Equal(19, fixtures.EnumerateArray().Count(entry =>
            entry.GetProperty("status").GetString() == "partial"));

        IReadOnlyDictionary<string, string> extensions = new Dictionary<string, string>
        {
            ["mathml"] = ".mathml",
            ["latex"] = ".tex",
            ["unicodemath"] = ".txt",
        };
        foreach (JsonElement fixture in fixtures.EnumerateArray())
        {
            string relativeSource = fixture.GetProperty("source").GetString()!
                .Replace('/', Path.DirectorySeparatorChar);
            JsonElement outputs = fixture.GetProperty("outputs");
            foreach ((string format, string extension) in extensions)
            {
                bool converted = outputs.GetProperty(format).GetProperty("status").GetString() == "converted";
                string outputPath = Path.Combine(
                    oracleRoot,
                    format,
                    Path.ChangeExtension(relativeSource, extension));
                Assert.Equal(converted, File.Exists(outputPath));
            }
        }
    }

    [Fact]
    public void CompletePinnedCorpusConvertsThroughEveryWriterWithoutUnsupportedSemanticNodes()
    {
        List<string> failures = new();
        foreach (string fixture in OmmlTestData.UpstreamFixtures())
        {
            string relativePath = Path.GetRelativePath(OmmlTestData.UpstreamRoot, fixture)
                .Replace(Path.DirectorySeparatorChar, '/');
            string source = File.ReadAllText(fixture);
            if (relativePath is "187.omml" or "issue-158.omml")
            {
                Assert.Throws<DxpOmmlParseException>(() => DxpOmmlConverter.ToText(source));
                continue;
            }
            XDocument sourceDocument = XDocument.Parse(source);

            foreach (DxpOmmlOutputFormat format in Enum.GetValues<DxpOmmlOutputFormat>())
            {
                try
                {
                    DxpOmmlConversionResult result = DxpOmmlConverter.Convert(source, format);
                    if (string.IsNullOrEmpty(result.Output))
                        failures.Add($"{relativePath} [{format}]: empty output");
                    if (result.Diagnostics.Any(diagnostic => diagnostic.Code == "OMML001"))
                    {
                        string elements = string.Join(", ", result.Diagnostics
                            .Where(diagnostic => diagnostic.Code == "OMML001")
                            .Select(diagnostic => diagnostic.ElementName)
                            .Distinct(StringComparer.Ordinal));
                        failures.Add($"{relativePath} [{format}]: unsupported {elements}");
                    }

                    if (format == DxpOmmlOutputFormat.MathMl)
                        _ = XDocument.Parse(result.Output);
                    if (format == DxpOmmlOutputFormat.Text)
                    {
                        XNamespace math = OmmlTestData.MathNamespace;
                        foreach (IGrouping<string, XElement> literals in sourceDocument.Descendants(math + "t")
                                     .Where(element => !element.Ancestors(math + "phant").Any(phantom =>
                                         phantom.Element(math + "phantPr")?.Element(math + "show")?
                                             .Attribute(math + "val")?.Value is "0" or "false" or "off"))
                                     .Where(element => element.Value.Length != 0 && element.Value != "\u200B")
                                     .GroupBy(element => element.Value, StringComparer.Ordinal))
                        {
                            if (CountOccurrences(result.Output, literals.Key) < literals.Count())
                                failures.Add($"{relativePath} [Text]: missing literal occurrence {JsonSerializer.Serialize(literals.Key)}");
                        }
                    }
                }
                catch (Exception exception) when (exception is DxpOmmlParseException or XmlException)
                {
                    failures.Add($"{relativePath} [{format}]: {exception.GetType().Name}: {exception.Message}");
                }
            }
        }

        Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));

        static int CountOccurrences(string value, string search)
        {
            int count = 0;
            for (int index = 0; (index = value.IndexOf(search, index, StringComparison.Ordinal)) >= 0;
                 index += search.Length)
                count++;
            return count;
        }
    }

    [Fact]
    public void NamedNormativeFixturesHaveReviewedReadableTextExpectations()
    {
        IReadOnlyList<string> fixtures = OmmlTestData.NormativeFixtures();
        Assert.NotEmpty(fixtures);
        foreach (string fixture in fixtures)
        {
            string expectation = Path.ChangeExtension(fixture, ".text");
            Assert.True(File.Exists(expectation), $"Missing expectation for {Path.GetFileName(fixture)}");
            Assert.Equal(File.ReadAllText(expectation).TrimEnd('\r', '\n'),
                DxpOmmlConverter.ToText(File.ReadAllText(fixture)));
        }
    }

    [Fact]
    public void NamedInvalidNormativeFixturesAreRejected()
    {
        IReadOnlyList<string> fixtures = OmmlTestData.InvalidNormativeFixtures();
        Assert.NotEmpty(fixtures);
        Assert.All(fixtures, fixture => Assert.Throws<DxpOmmlParseException>(() =>
            DxpOmmlConverter.ToText(File.ReadAllText(fixture))));
    }

    [Fact]
    public void OracleComparisonInventoryIsExplicitlyAudited()
    {
        Dictionary<DxpOmmlOutputFormat, (string Folder, string Extension)> formats = new()
        {
            [DxpOmmlOutputFormat.MathMl] = ("mathml", ".mathml"),
            [DxpOmmlOutputFormat.Latex] = ("latex", ".tex"),
            [DxpOmmlOutputFormat.UnicodeMath] = ("unicodemath", ".txt"),
        };
        Dictionary<DxpOmmlOutputFormat, int> available = formats.Keys.ToDictionary(format => format, _ => 0);
        Dictionary<DxpOmmlOutputFormat, int> exact = formats.Keys.ToDictionary(format => format, _ => 0);
        Dictionary<DxpOmmlOutputFormat, int> invalidOracle = formats.Keys.ToDictionary(format => format, _ => 0);

        foreach (string fixture in OmmlTestData.UpstreamFixtures())
        {
            string relative = Path.GetRelativePath(OmmlTestData.UpstreamRoot, fixture);
            if (relative is "187.omml" or "issue-158.omml") continue;
            string source = File.ReadAllText(fixture);
            foreach ((DxpOmmlOutputFormat format, (string folder, string extension)) in formats)
            {
                string oracle = Path.Combine(OmmlTestData.FixtureRoot, "OracleGenerated", folder,
                    Path.ChangeExtension(relative, extension));
                if (!File.Exists(oracle)) continue;
                available[format]++;
                string actual = DxpOmmlConverter.Convert(source, format).Output;
                string expected = File.ReadAllText(oracle).TrimEnd('\r', '\n');
                bool matches;
                try
                {
                    matches = format == DxpOmmlOutputFormat.MathMl
                        ? XmlCanonicalizer.Canonicalize(actual) == XmlCanonicalizer.Canonicalize(expected)
                        : actual == expected;
                }
                catch (XmlException)
                {
                    invalidOracle[format]++;
                    continue;
                }
                if (matches) exact[format]++;
            }
        }

        Assert.Equal(276, available[DxpOmmlOutputFormat.MathMl]);
        Assert.Equal(277, available[DxpOmmlOutputFormat.Latex]);
        Assert.Equal(259, available[DxpOmmlOutputFormat.UnicodeMath]);
        Assert.Equal(1, invalidOracle[DxpOmmlOutputFormat.MathMl]);
        Assert.Equal(0, invalidOracle[DxpOmmlOutputFormat.Latex]);
        Assert.Equal(0, invalidOracle[DxpOmmlOutputFormat.UnicodeMath]);
        Assert.True(exact[DxpOmmlOutputFormat.Latex] >= 11);
        Assert.True(exact[DxpOmmlOutputFormat.UnicodeMath] >= 61);
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
