using System.Collections.Concurrent;
using System.Buffers.Binary;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace DocxportNet.Doc;

/// <summary>Resolves installed or explicitly supplied font files for advance measurement.</summary>
internal static class DocSystemFontAdvances
{
    private static readonly ConcurrentDictionary<string, Lazy<DocOpenTypeAdvances?>> Cache =
        new(StringComparer.OrdinalIgnoreCase);

    public static DocOpenTypeAdvances? Find(string? family, bool bold = false,
        bool italic = false, bool includeCloudFonts = true)
    {
        if (string.IsNullOrWhiteSpace(family)) return null;
        var face = family + (bold && italic ? " Bold Italic" :
            bold ? " Bold" : italic ? " Italic" : "");
        var directory = Environment.GetEnvironmentVariable("DOCXPORT_FONT_DIRECTORY");
        return Cache.GetOrAdd((directory ?? "") + "|" + includeCloudFonts + "|" + face, _ =>
            new Lazy<DocOpenTypeAdvances?>(() => Load(face, directory, includeCloudFonts))).Value;
    }

    private static DocOpenTypeAdvances? Load(string family, string? overrideDirectory,
        bool includeCloudFonts)
    {
        try
        {
            // An explicitly supplied font file can supply metrics without
            // installing the face or silently accepting a system substitute.
            if (Path.IsPathRooted(overrideDirectory) && family.All(ch =>
                char.IsLetterOrDigit(ch) || ch is ' ' or '-'))
            {
                var file = Path.Combine(overrideDirectory!, family.Replace(' ', '-') + ".ttf");
                if (File.Exists(file))
                    return new DocOpenTypeAdvances(File.ReadAllBytes(file));
            }
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return null;
            const string fontsKey = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts";
            foreach (var (hive, directory) in new[]
            {
                (Registry.LocalMachine, Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts")),
                (Registry.CurrentUser, Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Microsoft", "Windows", "Fonts"))
            })
            {
                using var key = hive.OpenSubKey(fontsKey);
                if (key == null) continue;
                foreach (var name in key.GetValueNames())
                {
                    var suffix = name.EndsWith(" (TrueType)", StringComparison.OrdinalIgnoreCase)
                        ? " (TrueType)" : name.EndsWith(" (OpenType)", StringComparison.OrdinalIgnoreCase)
                            ? " (OpenType)" : null;
                    if (suffix == null) continue;
                    var faces = name.Substring(0, name.Length - suffix.Length)
                        .Split(new[] { " & " }, StringSplitOptions.None);
                    var faceIndex = Array.FindIndex(faces, face =>
                        face.Equals(family, StringComparison.OrdinalIgnoreCase));
                    if (faceIndex < 0 || key.GetValue(name) is not string file) continue;
                    var path = Path.IsPathRooted(file) ? file : Path.Combine(directory, file);
                    if (File.Exists(path))
                        return new DocOpenTypeAdvances(File.ReadAllBytes(path), faceIndex);
                }
            }
            var cache = Path.Combine(Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData), "Microsoft", "FontCache");
            if (includeCloudFonts && Directory.Exists(cache))
                foreach (var version in Directory.GetDirectories(cache))
                {
                    var folder = Path.Combine(version, "CloudFonts", family.Replace(" Bold Italic", "")
                        .Replace(" Bold", "").Replace(" Italic", ""));
                    if (!Directory.Exists(folder)) continue;
                    foreach (var path in Directory.GetFiles(folder, "*.ttf").OrderBy(path => path))
                    {
                        var bytes = File.ReadAllBytes(path);
                        if (MatchesStyle(bytes, family.Contains(" Bold", StringComparison.Ordinal),
                            family.EndsWith(" Italic", StringComparison.Ordinal)))
                            return new DocOpenTypeAdvances(bytes);
                    }
                }
            return null;
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or
            PlatformNotSupportedException or InvalidDataException or NotSupportedException)
        {
            return null;
        }
    }

    private static bool MatchesStyle(byte[] bytes, bool bold, bool italic)
    {
        if (bytes.Length < 12) return false;
        var count = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(4, 2));
        for (var index = 0; index < count && 12 + (index + 1) * 16 <= bytes.Length; index++)
        {
            var record = bytes.AsSpan(12 + index * 16, 16);
            if (!record.Slice(0, 4).SequenceEqual("OS/2"u8)) continue;
            var offset = BinaryPrimitives.ReadUInt32BigEndian(record.Slice(8, 4));
            if (offset > bytes.Length - 64) return false;
            var weight = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan((int)offset + 4, 2));
            var flags = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan((int)offset + 62, 2));
            return weight == (bold ? 700 : 400) && ((flags & 1) != 0) == italic;
        }
        return false;
    }
}
