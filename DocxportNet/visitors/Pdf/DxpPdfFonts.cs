using System.Collections.Concurrent;
using PdfSharp.Fonts;

namespace DocxportNet.Visitors.Pdf;

/// <summary>Connects a Docxport font provider to Pdfxport's process-wide PDF font resolver.
/// Configure once, before creating any PDF fonts or rendering a PDF.</summary>
public static class DxpPdfFonts
{
    /// <summary>Installs a provider for PDF export. A null provider uses system fonts.</summary>
    /// <remarks>Pdfxport owns global font state. Call this once during application startup;
    /// do not change providers between concurrent exports.</remarks>
    public static void Configure(IDxpPdfFontProvider? provider = null)
    {
        if (OperatingSystem.IsWindows())
            GlobalFontSettings.UseWindowsFontsUnderWindows = true;
        GlobalFontSettings.FontResolver = new Resolver(provider ?? new DxpSystemPdfFontProvider());
    }

    private sealed class Resolver(IDxpPdfFontProvider provider) : IFontResolver
    {
        private readonly ConcurrentDictionary<string, byte[]> _faces = new(StringComparer.Ordinal);

        public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            // Pdfxport's Windows resolver handles the installed-font inventory and TTC faces.
            // A caller-supplied provider takes precedence, including on Windows.
            if (OperatingSystem.IsWindows() && provider is DxpSystemPdfFontProvider)
                return PlatformFontResolver.ResolveTypeface(familyName, isBold, isItalic);

            var face = provider.Resolve(familyName, isBold, isItalic);
            if (face is null)
                return null;
            _faces.TryAdd(face.Key, face.Data);
            return new FontResolverInfo(face.Key, face.SimulateBold, face.SimulateItalic);
        }

        public byte[]? GetFont(string faceName) =>
            _faces.TryGetValue(faceName, out var data) ? data : null;
    }
}
