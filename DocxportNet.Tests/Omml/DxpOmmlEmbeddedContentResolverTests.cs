using DocxportNet.Omml;
using DocxportNet.Walker;

namespace DocxportNet.Tests.Omml;

public sealed class DxpOmmlEmbeddedContentResolverTests
{
    private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void LatexCanResolveEmbeddedWordprocessingMlThroughInjectedResolver()
    {
        RecordingResolver resolver = new("resolved_#");
        string omml = $"""
            <m:oMath xmlns:m="{MathNamespace}" xmlns:w="{WordNamespace}">
              <w:hyperlink><w:r><w:t>fallback</w:t></w:r></w:hyperlink>
            </m:oMath>
            """;

        DxpOmmlConversionResult result = DxpOmmlConverter.Convert(
            omml,
            DxpOmmlOutputFormat.Latex,
            new DxpOmmlConversionOptions { EmbeddedContentResolver = resolver });

        Assert.Equal(@"resolved\_\#", result.Output);
        Assert.NotNull(resolver.Request);
        Assert.Equal("w:hyperlink", resolver.Request!.ElementName);
        Assert.Equal(DxpOmmlOutputFormat.Latex, resolver.Request.OutputFormat);
        Assert.Contains("fallback", resolver.Request.XmlElement!.Value, StringComparison.Ordinal);
    }

    [Fact]
    public void WalkerResolverUsesNormalRunProcessingForLatexText()
    {
        string omml = $"""
            <m:oMath xmlns:m="{MathNamespace}" xmlns:w="{WordNamespace}">
              <w:r><w:t>A_B</w:t><w:tab/><w:t>C</w:t><w:noBreakHyphen/><w:t>D</w:t></w:r>
            </m:oMath>
            """;

        string result = DxpOmmlConverter.ToLatex(
            omml,
            new DxpOmmlConversionOptions {
                EmbeddedContentResolver = new DxpWalkerOmmlEmbeddedContentResolver()
            });

        Assert.Equal("A\\_B\tC-D", result);
    }

    [Fact]
    public void StandaloneConversionKeepsLightweightVisibleTextFallback()
    {
        string omml = $"""
            <m:oMath xmlns:m="{MathNamespace}" xmlns:w="{WordNamespace}">
              <w:hyperlink><w:r><w:t>A_B</w:t></w:r></w:hyperlink>
            </m:oMath>
            """;

        Assert.Equal(@"A\_B", DxpOmmlConverter.ToLatex(omml));
    }

    private sealed class RecordingResolver(string result) : IDxpOmmlEmbeddedContentResolver
    {
        public DxpOmmlEmbeddedContentRequest? Request { get; private set; }

        public string? Resolve(DxpOmmlEmbeddedContentRequest request)
        {
            Request = request;
            return result;
        }
    }
}
