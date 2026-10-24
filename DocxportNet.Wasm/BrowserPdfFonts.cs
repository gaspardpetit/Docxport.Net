using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;
using DocxportNet.Visitors.Pdf;
using PdfSharp.Internal.OpenType;

namespace DocxportNet.Wasm;

internal sealed class BrowserPdfFonts : IDxpPdfFontProvider
{
    private readonly ConcurrentDictionary<string, DxpPdfFontFace> _faces =
        new(StringComparer.OrdinalIgnoreCase);
    private string? _defaultFamily;
    private readonly ConcurrentDictionary<string, Lazy<OpenTypeFontDescriptor?>> _descriptors = new();

    public IReadOnlyCollection<string> Families => _defaultFamily is { } first
        ? [first, .. _faces.Keys.Where(name => !name.Equals(first,
            StringComparison.OrdinalIgnoreCase)).OrderBy(name => name, StringComparer.OrdinalIgnoreCase)]
        : [];

    public void Register(string family, byte[] data)
    {
        if (string.IsNullOrWhiteSpace(family))
            throw new ArgumentException("A PDF font family is required.", nameof(family));
        if (data == null || data.Length == 0)
            throw new ArgumentException("Non-empty OpenType font bytes are required.", nameof(data));
        string name = family.Trim();
        var face = new DxpPdfFontFace(Convert.ToHexString(SHA256.HashData(data)), data.ToArray());
        if (!_faces.TryAdd(name, face) &&
            !_faces[name].Key.Equals(face.Key, StringComparison.Ordinal))
            throw new InvalidOperationException($"PDF font family '{family}' is already registered with different data.");
        Interlocked.CompareExchange(ref _defaultFamily, name, null);
    }

    public bool IsAvailable(string familyName, bool bold, bool italic) =>
        _faces.ContainsKey(familyName);

    public DxpPdfFontFace? Resolve(string familyName, bool bold, bool italic)
    {
        if (!_faces.TryGetValue(familyName, out var face) &&
            (_defaultFamily is null || !_faces.TryGetValue(_defaultFamily, out face)))
            return null;
        return face with { SimulateBold = bold, SimulateItalic = italic };
    }

    public DxpPdfFontMatch? ResolveForCharacter(string familyName, bool bold, bool italic, Rune character)
    {
        IEnumerable<string> candidates = _faces.ContainsKey(familyName)
            ? [familyName, .. _faces.Keys.Where(name => !name.Equals(familyName,
                StringComparison.OrdinalIgnoreCase)).OrderBy(name => name, StringComparer.OrdinalIgnoreCase)]
            : _faces.Keys.OrderBy(name => name, StringComparer.OrdinalIgnoreCase);
        foreach (string candidate in candidates)
        {
            var face = _faces[candidate];
            if (HasGlyph(face, character))
                return new DxpPdfFontMatch(candidate, face with
                { SimulateBold = bold, SimulateItalic = italic });
        }
        return null;
    }

    private bool HasGlyph(DxpPdfFontFace face, Rune character)
    {
        var descriptor = _descriptors.GetOrAdd(face.Key, _ =>
            new Lazy<OpenTypeFontDescriptor?>(() =>
            {
                try
                {
                    var source = OpenTypeFontSource.GetOrCreateFrom(face.Data);
                    return OpenTypeFontFace.GetOrCreateFrom(source).OTDescriptor;
                }
                catch (InvalidOperationException) { return null; }
                catch (InvalidDataException) { return null; }
            })).Value;
        return descriptor?.CodepointToGlyphIndex(character.Value) > 0;
    }
}
