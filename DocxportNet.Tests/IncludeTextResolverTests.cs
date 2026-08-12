using DocxportNet.Fields;
using DocxportNet.Fields.Resolution;
using System.Text;

namespace DocxportNet.Tests;

public sealed class IncludeTextResolverTests
{
    [Fact]
    public async Task ResolvesRelativePathBelowConfiguredRoot()
    {
        using var fixture = new ResolverFixture();
        string file = fixture.Write([1, 2, 3], "Headers", "Example.docx");
        var resolver = new DxpFileSystemIncludeTextResolver([fixture.Root]);

        var result = await resolver.ResolveAsync(
            new DxpIncludeTextRequest(Path.Combine("Headers", "Example.docx")),
            new DxpFieldEvalContext());

        Assert.NotNull(result);
        Assert.Equal(Path.GetFullPath(file), result!.Identity);
        Assert.Equal([1, 2, 3], result.Content);
        Assert.Equal(DxpIncludeTextSourceFormat.Docx, result.Format);
    }

    [Theory]
    [InlineData("Signature.htm")]
    [InlineData("Signature.HTML")]
    public async Task ReportsHtmlSourceFormat(string name)
    {
        using var fixture = new ResolverFixture();
        fixture.Write(Encoding.UTF8.GetBytes("<p>signature</p>"), name);
        var resolver = new DxpFileSystemIncludeTextResolver([fixture.Root]);

        var result = await resolver.ResolveAsync(new DxpIncludeTextRequest(name), new DxpFieldEvalContext());

        Assert.NotNull(result);
        Assert.Equal(DxpIncludeTextSourceFormat.Html, result!.Format);
    }

    [Fact]
    public async Task RemapsLegacyAbsolutePathUsingTrailingSegments()
    {
        using var fixture = new ResolverFixture();
        string file = fixture.Write([4, 5, 6], "Headers", "Example.docx");
        var resolver = new DxpFileSystemIncludeTextResolver([fixture.Root]);

        var result = await resolver.ResolveAsync(
            new DxpIncludeTextRequest(@"Z:\Legacy\Templates\Headers\Example.docx"),
            new DxpFieldEvalContext());

        Assert.NotNull(result);
        Assert.Equal(Path.GetFullPath(file), result!.Identity);
        Assert.Equal([4, 5, 6], result.Content);
    }

    [Fact]
    public async Task RemapsLegacyUncPathUsingTrailingSegments()
    {
        using var fixture = new ResolverFixture();
        string file = fixture.Write([10, 11, 12], "Bodies", "Example.docx");
        var resolver = new DxpFileSystemIncludeTextResolver([fixture.Root]);

        var result = await resolver.ResolveAsync(
            new DxpIncludeTextRequest(@"\\legacy-server\templates\Bodies\Example.docx"),
            new DxpFieldEvalContext());

        Assert.NotNull(result);
        Assert.Equal(Path.GetFullPath(file), result!.Identity);
        Assert.Equal([10, 11, 12], result.Content);
    }

    [Fact]
    public async Task DoesNotReadExistingFileOutsideConfiguredRoots()
    {
        using var allowed = new ResolverFixture();
        using var outside = new ResolverFixture();
        string outsideFile = outside.Write([7, 8, 9], "Private.docx");
        var resolver = new DxpFileSystemIncludeTextResolver([allowed.Root]);

        var result = await resolver.ResolveAsync(
            new DxpIncludeTextRequest(outsideFile),
            new DxpFieldEvalContext());

        Assert.Null(result);
    }

    [Fact]
    public async Task ReadsTemplateThatIsOpenForEditing()
    {
        using var fixture = new ResolverFixture();
        string file = fixture.Write([13, 14, 15], "Bodies", "Open.docx");
        var resolver = new DxpFileSystemIncludeTextResolver([fixture.Root]);
        using var editingStream = new FileStream(
            file,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.ReadWrite);

        var result = await resolver.ResolveAsync(
            new DxpIncludeTextRequest(Path.Combine("Bodies", "Open.docx")),
            new DxpFieldEvalContext());

        Assert.NotNull(result);
        Assert.Equal([13, 14, 15], result!.Content);
    }

    private sealed class ResolverFixture : IDisposable
    {
        public ResolverFixture()
        {
            Root = Path.Combine(Path.GetTempPath(), "docxport-includetext-tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Root);
        }

        public string Root { get; }

        public string Write(byte[] content, params string[] relativeParts)
        {
            string path = relativeParts.Aggregate(Root, Path.Combine);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllBytes(path, content);
            return path;
        }

        public void Dispose()
        {
            if (Directory.Exists(Root))
                Directory.Delete(Root, recursive: true);
        }
    }
}
