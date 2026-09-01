using System.Diagnostics;
using System.Xml.Linq;
using DocxportNet.Omml;

namespace DocxportNet.Tests.Omml;

public sealed class DxpOmmlHardeningTests
{
    private const string M = OmmlTestData.MathNamespace;

    [Fact]
    public void DeterministicMalformedMutationCorpusFailsSafely()
    {
        string seed = $"<m:oMath xmlns:m=\"{M}\"><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>";
        List<string> mutations =
        [
            .. Enumerable.Range(0, seed.Length).Where(index => index % 7 == 0).Select(index => seed[..index]),
            seed.Replace("m:oMath", "x:oMath", StringComparison.Ordinal),
            seed.Replace("<m:t>a</m:t>", "<m:t>&bogus;</m:t>", StringComparison.Ordinal),
            seed.Replace("<m:f>", "<!DOCTYPE f [<!ENTITY x SYSTEM 'file:///C:/Windows/win.ini'>]><m:f>", StringComparison.Ordinal),
            "\0" + seed,
        ];

        foreach (string mutation in mutations)
        {
            bool converted = DxpOmmlConverter.TryConvert(
                mutation, DxpOmmlOutputFormat.MathMl, out DxpOmmlConversionResult? result,
                out DxpOmmlException? error);

            Assert.False(converted);
            Assert.Null(result);
            Assert.IsAssignableFrom<DxpOmmlException>(error);
        }
    }

    [Fact]
    public void UnknownExtensionsCannotInjectMarkupIntoAnyOutput()
    {
        string input = $"<m:oMath xmlns:m=\"{M}\" xmlns:x=\"urn:future\"><x:future><m:t>&lt;script&gt;alert(1)&lt;/script&gt;</m:t></x:future></m:oMath>";

        string mathml = DxpOmmlConverter.ToMathMl(input);
        XElement parsed = XElement.Parse(mathml);
        Assert.DoesNotContain("<script>", mathml, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("<script>alert(1)</script>", parsed.Value);
        Assert.DoesNotContain("\\script", DxpOmmlConverter.ToLatex(input), StringComparison.Ordinal);
        Assert.Equal("<script>alert(1)</script>", DxpOmmlConverter.ToText(input));
    }

    [Fact]
    public void ThousandsOfSiblingEquationsRemainLinearEnoughForInteractiveUse()
    {
        static string Document(int count) => $"<m:oMathPara xmlns:m=\"{M}\">{string.Concat(Enumerable.Repeat("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>", count))}</m:oMathPara>";
        DxpOmmlConversionOptions options = new() { MaxElementCount = 20_000 };

        _ = DxpOmmlConverter.ToText(Document(10), options);
        Stopwatch timer = Stopwatch.StartNew();
        string output = DxpOmmlConverter.ToText(Document(2_000), options);
        timer.Stop();

        Assert.Equal(2_000, output.Length);
        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(10), $"Conversion took {timer.Elapsed}.");
    }
}
