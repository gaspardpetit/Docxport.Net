namespace DocxportNet.Doc;

/// <summary>Minimal default structures for an unformatted main-story document.</summary>
internal static class DocDefaultStructures
{
    public static byte[] CreateStyleSheet()
    {
        // LPStshi: 18-byte Stshif and fifteen fixed-index empty LPStd entries.
        var bytes = new byte[20 + 15 * 2];
        new DocStyleSheetHeader(15, 10, 1, 0, 15).Write().CopyTo(bytes, 0);
        return bytes;
    }

    public static byte[] CreateSectionTable(int characterCount)
    {
        // PlcfSed with one section and no exception properties (fcSepx=-1).
        return new DocSectionTable(checked((uint)characterCount), -1).Write();
    }
}
