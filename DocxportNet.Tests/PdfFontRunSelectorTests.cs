using System.Text;
using DocxportNet.Visitors.Pdf;

namespace DocxportNet.Tests;

public sealed class PdfFontRunSelectorTests
{
    [Fact]
    public void SplitsAtCharacterFallbackAndPreservesUnicodeScalars()
    {
        var runs = DxpPdfFontRunSelector.Split("A日😀本B", "Latin", false, false, new Provider());

        Assert.Equal(new[] { "A", "日", "😀", "本", "B" }, runs.Select(run => run.Text));
        Assert.Equal(new[] { "Latin", "Japanese", "Emoji", "Japanese", "Latin" },
            runs.Select(run => run.Family));
    }

    [Fact]
    public void KeepsOneRunWhenAllCharactersUseTheRequestedFace()
    {
        var runs = DxpPdfFontRunSelector.Split("ordinary text", "Latin", false, false, new Provider());

        var run = Assert.Single(runs);
        Assert.Equal("ordinary text", run.Text);
        Assert.Equal("Latin", run.Family);
    }

    private sealed class Provider : IDxpPdfFontProvider
    {
        public DxpPdfFontFace? Resolve(string familyName, bool bold, bool italic) =>
            new(familyName, []);

        public DxpPdfFontMatch? ResolveForCharacter(string familyName, bool bold, bool italic, Rune character)
        {
            string family = character.Value switch
            {
                0x65E5 or 0x672C => "Japanese",
                0x1F600 => "Emoji",
                _ => familyName
            };
            return new DxpPdfFontMatch(family, new DxpPdfFontFace(family, []));
        }
    }
}
