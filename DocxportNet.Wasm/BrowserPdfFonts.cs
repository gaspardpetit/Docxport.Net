using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;
using DocxportNet.Visitors.Pdf;
using PdfSharp.Internal.OpenType;

namespace DocxportNet.Wasm;

internal sealed class BrowserPdfFonts : IDxpPdfFontProvider
{
    private sealed record RegisteredFace(string Family, bool Bold, bool Italic, DxpPdfFontFace Face);
    private readonly ConcurrentDictionary<string, RegisteredFace> _faces =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, string> _aliases = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, string> _currentAliases = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, int> _familyVersions = new(StringComparer.OrdinalIgnoreCase);
    private string? _defaultFamily;
    private readonly ConcurrentDictionary<string, Lazy<OpenTypeFontDescriptor?>> _descriptors = new();

    public IReadOnlyCollection<string> Families => _defaultFamily is { } first
        ? [first, .. _faces.Values.Select(item => item.Family).Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(name => !name.Equals(first, StringComparison.OrdinalIgnoreCase))
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)]
        : [];

    public void Register(string family, byte[] data) => Register(family, false, false, data);

    public void Register(string family, bool bold, bool italic, byte[] data)
    {
        if (string.IsNullOrWhiteSpace(family))
            throw new ArgumentException("A PDF font family is required.", nameof(family));
        if (data == null || data.Length == 0)
            throw new ArgumentException("Non-empty OpenType font bytes are required.", nameof(data));
        string name = family.Trim();
        var face = new DxpPdfFontFace(Convert.ToHexString(SHA256.HashData(data)), data.ToArray());
        var key = Key(name, bold, italic);
        if (!_faces.TryAdd(key, new RegisteredFace(name, bold, italic, face)))
        {
            if (!_faces[key].Face.Key.Equals(face.Key, StringComparison.Ordinal))
                throw new InvalidOperationException($"PDF font face '{family}' ({bold}, {italic}) is already registered with different data.");
            return;
        }
        int version = _familyVersions.AddOrUpdate(name, 1, (_, previous) => previous + 1);
        string familyHash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(name)).AsSpan(0, 8));
        string alias = $"DxpLocal-{familyHash}-{version}";
        _aliases[alias] = name;
        _currentAliases[name] = alias;
        Interlocked.CompareExchange(ref _defaultFamily, name, null);
    }

    public bool IsAvailable(string familyName, bool bold, bool italic) =>
        _faces.ContainsKey(Key(OriginalFamily(familyName), bold, italic)) ||
        _faces.Values.Any(item => item.Family.Equals(OriginalFamily(familyName), StringComparison.OrdinalIgnoreCase));

    public DxpPdfFontFace? Resolve(string familyName, bool bold, bool italic)
    {
        var face = FaceFor(OriginalFamily(familyName), bold, italic) ??
            (_defaultFamily is { } fallback ? FaceFor(fallback, bold, italic) : null);
        return face is null ? null : WithSimulation(face, bold, italic);
    }

    public DxpPdfFontMatch? ResolveForCharacter(string familyName, bool bold, bool italic, Rune character)
    {
        string requested = OriginalFamily(familyName);
        IEnumerable<string> candidates = IsAvailable(requested, bold, italic)
            ? [requested, .. Families.Where(name => !name.Equals(requested, StringComparison.OrdinalIgnoreCase))]
            : Families;
        foreach (string candidate in candidates)
        {
            var face = FaceFor(candidate, bold, italic);
            if (face != null && HasGlyph(face.Face, character))
                return new DxpPdfFontMatch(AliasFor(candidate), WithSimulation(face, bold, italic));
        }
        return null;
    }

    private static string Key(string family, bool bold, bool italic) =>
        $"{family}\u001f{(bold ? 1 : 0)}{(italic ? 1 : 0)}";

    private string OriginalFamily(string family) => _aliases.GetValueOrDefault(family) ?? family;

    private string AliasFor(string family) => _currentAliases.GetValueOrDefault(family) ?? family;

    private RegisteredFace? FaceFor(string family, bool bold, bool italic) =>
        _faces.GetValueOrDefault(Key(family, bold, italic)) ??
        _faces.GetValueOrDefault(Key(family, false, false)) ??
        _faces.Values.FirstOrDefault(item => item.Family.Equals(family, StringComparison.OrdinalIgnoreCase));

    private static DxpPdfFontFace WithSimulation(RegisteredFace face, bool bold, bool italic) =>
        face.Face with { SimulateBold = bold && !face.Bold, SimulateItalic = italic && !face.Italic };

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
