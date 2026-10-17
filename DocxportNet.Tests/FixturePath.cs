namespace DocxportNet.Tests;

internal static class FixturePath
{
    public static string Combine(params string[] parts) => Resolve(Path.Combine(parts));

    public static string Resolve(string path)
    {
        var docxSegment = Path.DirectorySeparatorChar + "fixtures" + Path.DirectorySeparatorChar +
            "docx" + Path.DirectorySeparatorChar;
        var docSegment = Path.DirectorySeparatorChar + "fixtures" + Path.DirectorySeparatorChar +
            "doc" + Path.DirectorySeparatorChar;
        return Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".doc" => path.Replace(docxSegment, docSegment, StringComparison.OrdinalIgnoreCase),
            ".docx" => path.Replace(docSegment, docxSegment, StringComparison.OrdinalIgnoreCase),
            _ => path,
        };
    }
}
