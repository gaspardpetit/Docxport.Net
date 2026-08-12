namespace DocxportNet.Fields.Resolution;

public sealed class DxpFileSystemIncludeTextResolver : IDxpIncludeTextResolver
{
    private readonly IReadOnlyList<string> _roots;

    public DxpFileSystemIncludeTextResolver(IEnumerable<string> includePaths)
    {
        if (includePaths == null)
            throw new ArgumentNullException(nameof(includePaths));
        _roots = includePaths
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Select(Path.GetFullPath)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (_roots.Count == 0)
            throw new ArgumentException("At least one include path is required.", nameof(includePaths));
    }

    public IReadOnlyList<string> IncludePaths => _roots;

    public Task<DxpIncludeTextSource?> ResolveAsync(
        DxpIncludeTextRequest request,
        DxpFieldEvalContext context,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(request.Path))
            return Task.FromResult<DxpIncludeTextSource?>(null);

        foreach (string candidate in CandidatePaths(request.Path))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!File.Exists(candidate))
                continue;

            return Task.FromResult<DxpIncludeTextSource?>(
                new DxpIncludeTextSource(candidate, ReadAllBytesShared(candidate)) {
                    Format = IsHtmlPath(candidate)
                        ? DxpIncludeTextSourceFormat.Html
                        : DxpIncludeTextSourceFormat.Docx
                });
        }

        return Task.FromResult<DxpIncludeTextSource?>(null);
    }

    private static byte[] ReadAllBytesShared(string path)
    {
        using var input = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        using var output = input.Length <= int.MaxValue
            ? new MemoryStream((int)input.Length)
            : new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private IEnumerable<string> CandidatePaths(string requestedPath)
    {
        var emitted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        string normalized = requestedPath.Trim().Replace('\\', Path.DirectorySeparatorChar).Replace('/', Path.DirectorySeparatorChar);
        bool hasRootedSyntax = HasRootedSyntax(requestedPath, normalized);

        if (Path.IsPathRooted(normalized))
        {
            string absolute = Path.GetFullPath(normalized);
            foreach (string root in _roots)
            {
                if (IsUnderRoot(absolute, root) && emitted.Add(absolute))
                    yield return absolute;
            }
        }

        string suffixSource = StripRoot(requestedPath, normalized);
        string[] segments = suffixSource
            .Split([Path.DirectorySeparatorChar], StringSplitOptions.RemoveEmptyEntries);
        foreach (string root in _roots)
        {
            int firstSuffix = hasRootedSyntax ? 0 : -1;
            for (int index = firstSuffix; index < segments.Length; index++)
            {
                string[] relativeSegments = index < 0 ? segments : segments.Skip(index).ToArray();
                if (relativeSegments.Length == 0)
                    continue;
                string candidate = Path.GetFullPath(relativeSegments.Aggregate(root, Path.Combine));
                if (IsUnderRoot(candidate, root) && emitted.Add(candidate))
                    yield return candidate;
            }
        }
    }

    private static bool HasRootedSyntax(string original, string normalized)
        => Path.IsPathRooted(normalized)
            || (original.Length >= 3 && char.IsLetter(original[0]) && original[1] == ':' && IsSeparator(original[2]))
            || (original.Length >= 2 && IsSeparator(original[0]) && IsSeparator(original[1]));

    private static string StripRoot(string original, string normalized)
    {
        if (Path.IsPathRooted(normalized))
        {
            string? pathRoot = Path.GetPathRoot(normalized);
            if (!string.IsNullOrEmpty(pathRoot))
                return normalized.Substring(pathRoot.Length);
        }

        if (original.Length >= 3 && char.IsLetter(original[0]) && original[1] == ':' && IsSeparator(original[2]))
            return normalized.Substring(3);

        return normalized.TrimStart(Path.DirectorySeparatorChar);
    }

    private static bool IsSeparator(char value) => value is '\\' or '/';

    private static bool IsHtmlPath(string path)
        => Path.GetExtension(path).Equals(".htm", StringComparison.OrdinalIgnoreCase)
            || Path.GetExtension(path).Equals(".html", StringComparison.OrdinalIgnoreCase);

    private static bool IsUnderRoot(string candidate, string root)
    {
        string normalizedRoot = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return candidate.StartsWith(normalizedRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
    }
}
