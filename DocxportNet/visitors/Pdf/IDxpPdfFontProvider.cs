namespace DocxportNet.Visitors.Pdf;

/// <summary>Supplies physical font data for PDF export. Hosts may preload fonts from any source,
/// including browser downloads, before starting an export.</summary>
public interface IDxpPdfFontProvider
{
    /// <summary>Returns a font face for the requested family and style, or null when unavailable.</summary>
    DxpPdfFontFace? Resolve(string familyName, bool bold, bool italic);
}

/// <param name="Key">A stable, unique identifier for this physical face.</param>
/// <param name="Data">The complete font file bytes.</param>
/// <param name="SimulateBold">Whether the renderer should synthesize bold.</param>
/// <param name="SimulateItalic">Whether the renderer should synthesize italic.</param>
public sealed record DxpPdfFontFace(string Key, byte[] Data,
    bool SimulateBold = false, bool SimulateItalic = false);
