using System.IO.Compression;
using System.Xml.Linq;

namespace DocxportNet.Wasm;

internal static class BrowserPdfFontPreflight
{
    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public static string[] Families(byte[] input)
    {
        byte[] docx = input.AsSpan().StartsWith(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 })
            ? DocxportNet.DxpDocToDocx.Project(input).DocxBytes : input;
        var families = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var themes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var themeSchemes = new List<XElement>();
        using var stream = new MemoryStream(docx, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        foreach (var entry in archive.Entries)
        {
            if (!entry.FullName.StartsWith("word/", StringComparison.OrdinalIgnoreCase) ||
                !entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;
            using var part = entry.Open();
            var root = XDocument.Load(part).Root;
            if (root == null) continue;
            foreach (var fonts in root.Descendants(Word + "rFonts"))
            {
                foreach (var attribute in fonts.Attributes())
                {
                    if (attribute.Name.Namespace != Word || string.IsNullOrWhiteSpace(attribute.Value)) continue;
                    if (attribute.Name.LocalName is "ascii" or "hAnsi" or "eastAsia" or "cs")
                        families.Add(attribute.Value.Trim());
                    else if (attribute.Name.LocalName.EndsWith("Theme", StringComparison.Ordinal))
                        themes.Add(attribute.Value);
                }
            }
            if (entry.FullName.StartsWith("word/theme/", StringComparison.OrdinalIgnoreCase))
                themeSchemes.AddRange(root.Descendants(Drawing + "fontScheme"));
        }
        foreach (var scheme in themeSchemes)
        {
            foreach (var theme in themes)
            {
                string group = theme.StartsWith("major", StringComparison.OrdinalIgnoreCase) ? "majorFont" : "minorFont";
                var typeface = scheme.Element(Drawing + group)?.Element(Drawing + "latin")?.Attribute("typeface")?.Value;
                if (!string.IsNullOrWhiteSpace(typeface)) families.Add(typeface);
            }
        }
        // The browser sample supplies Noto Sans for documents with no declared family.
        if (families.Count == 0) families.Add("Noto Sans");
        return families.OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray();
    }
}
