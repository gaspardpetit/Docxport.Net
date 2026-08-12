namespace DocxportNet.Fields.Resolution;

public sealed record DxpIncludeTextRequest(string Path);

public enum DxpIncludeTextSourceFormat
{
    Auto,
    Docx,
    Html
}

public sealed record DxpIncludeTextSource(string Identity, byte[] Content)
{
    public DxpIncludeTextSourceFormat Format { get; init; } = DxpIncludeTextSourceFormat.Auto;
}

public interface IDxpIncludeTextResolver
{
    Task<DxpIncludeTextSource?> ResolveAsync(
        DxpIncludeTextRequest request,
        DxpFieldEvalContext context,
        CancellationToken cancellationToken = default);
}
