using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker.Context;

namespace DocxportNet.Tests;

public class ParagraphSpacingTests
{
    [Fact]
    public void DirectAfterSpacingRetainsInheritedLineSpacing()
    {
        using var document = OpenSample("file-sample_1MB.docx");
        var paragraph = document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .First(candidate => candidate.InnerText.StartsWith("Vestibulum neque massa", StringComparison.Ordinal));

        var spacing = new DxpStyleResolver(document).GetEffectiveParagraphSpacing(paragraph);

        Assert.Equal("225", spacing?.After?.Value);
        Assert.Equal("288", spacing?.Line?.Value);
        Assert.Equal(LineSpacingRuleValues.Auto, spacing?.LineRule?.Value);
    }

    [Fact]
    public void ExplicitSpacingOverridesInheritedAutoFlags()
    {
        using var document = OpenSample("Right to Repair.docx");
        var paragraphs = document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>();
        var paragraph = paragraphs.First(candidate =>
            candidate.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "NormalWeb");
        var resolver = new DxpStyleResolver(document);

        var spacing = resolver.GetEffectiveParagraphSpacing(paragraph);
        Assert.Equal("120", spacing?.Before?.Value);
        Assert.Equal("120", spacing?.After?.Value);
        Assert.Null(spacing?.BeforeAutoSpacing);
        Assert.Null(spacing?.AfterAutoSpacing);
        Assert.Equal("240", spacing?.Line?.Value);
        Assert.Equal("1", DxpParagraphStyleComputer.ComputeLineHeightCss(spacing));

        var heading = paragraphs.First(candidate => candidate.InnerText == "What it means");
        var headingSpacing = resolver.GetEffectiveParagraphSpacing(heading);
        Assert.True(headingSpacing?.BeforeAutoSpacing?.Value);
        Assert.True(headingSpacing?.AfterAutoSpacing?.Value);

        var explicitAutoOff = (Paragraph)heading.CloneNode(true);
        explicitAutoOff.ParagraphProperties!.SpacingBetweenLines = new SpacingBetweenLines {
            BeforeAutoSpacing = false,
            AfterAutoSpacing = false
        };
        var explicitSpacing = resolver.GetEffectiveParagraphSpacing(explicitAutoOff);
        Assert.Equal("100", explicitSpacing?.Before?.Value);
        Assert.Equal("100", explicitSpacing?.After?.Value);
        Assert.False(explicitSpacing?.BeforeAutoSpacing?.Value);
        Assert.False(explicitSpacing?.AfterAutoSpacing?.Value);
    }

    [Fact]
    public void DirectLineRuleRetainsInheritedLineWithoutMutatingSource()
    {
        using var document = OpenSample("file-sample_1MB.docx");
        var original = document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>()
            .First(candidate => candidate.InnerText.StartsWith("Vestibulum neque massa", StringComparison.Ordinal));
        var paragraph = (Paragraph)original.CloneNode(true);
        paragraph.ParagraphProperties!.SpacingBetweenLines = new SpacingBetweenLines {
            LineRule = LineSpacingRuleValues.Exact
        };

        var spacing = new DxpStyleResolver(document).GetEffectiveParagraphSpacing(paragraph);

        Assert.Equal("288", spacing?.Line?.Value);
        Assert.Equal(LineSpacingRuleValues.Exact, spacing?.LineRule?.Value);
        Assert.Null(paragraph.ParagraphProperties.SpacingBetweenLines.Line);
    }

    private static WordprocessingDocument OpenSample(string name)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "fixtures", "docx", "samples", name));
        return WordprocessingDocument.Open(path, false);
    }
}
