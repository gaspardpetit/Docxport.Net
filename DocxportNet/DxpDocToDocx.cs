using DocxportNet.Doc;

namespace DocxportNet;

/// <summary>
/// Builds a basic DOCX directly from a binary DOC by indexing its text and projecting the index.
/// This does not run the DOCX visitor pipeline.
/// </summary>
public static class DxpDocToDocx
{
    public static DocxProjectionResult Project(string path)
    {
        using var index = new DocTextIndexWalker().Index(path);
        return new DocToDocxProjector().Project(index);
    }

    public static DocxProjectionResult Project(byte[] bytes)
    {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var input = new MemoryStream(bytes, writable: false);
        return Project(input);
    }

    /// <summary>Projects a DOC stream without closing the caller's stream.</summary>
    public static DocxProjectionResult Project(Stream input)
    {
        using var index = new DocTextIndexWalker().Index(input);
        return new DocToDocxProjector().Project(index);
    }
}
