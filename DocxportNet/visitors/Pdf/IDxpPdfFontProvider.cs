using System.Text;

namespace DocxportNet.Visitors.Pdf;

/// <summary>Supplies physical font data for PDF export. Hosts may preload fonts from any source,
/// including browser downloads, before starting an export.</summary>
public interface IDxpPdfFontProvider
{
    /// <summary>Returns a font face for the requested family and style, or null when unavailable.</summary>
    DxpPdfFontFace? Resolve(string familyName, bool bold, bool italic);

    /// <summary>Checks whether a family and style can be resolved. Override to avoid loading
    /// font bytes for a simple availability check.</summary>
    bool IsAvailable(string familyName, bool bold, bool italic) => Resolve(familyName, bold, italic) != null;

    /// <summary>Chooses a face containing the character. The returned family must also
    /// resolve through <see cref="Resolve"/> so the renderer can load it. Returns null
    /// when glyph coverage is unknown or no suitable face exists.</summary>
    DxpPdfFontMatch? ResolveForCharacter(string familyName, bool bold, bool italic, Rune character) => null;
}

/// <summary>A physical face selected for a character, with the family name to use during layout.</summary>
public sealed record DxpPdfFontMatch(string FamilyName, DxpPdfFontFace Face);

/// <param name="Key">A stable, unique identifier for this physical face.</param>
/// <param name="Data">The complete font file bytes.</param>
/// <param name="SimulateBold">Whether the renderer should synthesize bold.</param>
/// <param name="SimulateItalic">Whether the renderer should synthesize italic.</param>
public sealed record DxpPdfFontFace(string Key, byte[] Data,
    bool SimulateBold = false, bool SimulateItalic = false);
