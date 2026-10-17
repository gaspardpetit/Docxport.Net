namespace DocxportNet.Tests;

public class FixturePathTests
{
    [Fact]
    public void DocPairUsesDocFormatDirectory()
    {
        var path = FixturePath.Combine("root", "fixtures", "docx", "legacy", "example.doc");

        Assert.Equal(Path.Combine("root", "fixtures", "doc", "legacy", "example.doc"), path);
    }

    [Fact]
    public void DocxPairUsesDocxFormatDirectory()
    {
        var path = FixturePath.Resolve(Path.Combine("root", "fixtures", "doc", "legacy", "example.docx"));

        Assert.Equal(Path.Combine("root", "fixtures", "docx", "legacy", "example.docx"), path);
    }

    [Fact]
    public void SamplePathsAreUnchanged()
    {
        var path = Path.Combine("root", "samples", "example.doc");

        Assert.Equal(path, FixturePath.Resolve(path));
    }
}
