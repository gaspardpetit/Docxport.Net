namespace DocxportNet.Doc;

/// <summary>Converts a binary DOC source to the DOCX input expected by the existing exporter.</summary>
internal static class DocInputProjection
{
    private static readonly byte[] CompoundSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    public static byte[] ProjectIfDoc(byte[] input)
    {
        if (!IsBinaryDoc(input)) return input;
        return DxpDocToDocx.Project(input).DocxBytes;
    }

    /// <summary>Returns null for DOCX paths so the existing path-based reader remains in use.</summary>
    public static byte[]? ProjectIfDoc(string path)
    {
        using var input = new FileStream(path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        var signature = new byte[CompoundSignature.Length];
        var read = 0;
        while (read < signature.Length)
        {
            var count = input.Read(signature, read, signature.Length - read);
            if (count == 0) return null;
            read += count;
        }
        if (!IsBinaryDoc(signature)) return null;
        input.Position = 0;
        return DxpDocToDocx.Project(input).DocxBytes;
    }

    private static bool IsBinaryDoc(byte[] input) =>
        input != null && input.AsSpan().StartsWith(CompoundSignature);
}
