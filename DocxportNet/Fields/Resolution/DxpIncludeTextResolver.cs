namespace DocxportNet.Fields.Resolution;

public sealed record DxpIncludeTextRequest(string Path);

public sealed record DxpIncludeTextSource(string Identity, byte[] Content);

public interface IDxpIncludeTextResolver
{
    Task<DxpIncludeTextSource?> ResolveAsync(
        DxpIncludeTextRequest request,
        DxpFieldEvalContext context,
        CancellationToken cancellationToken = default);
}
