using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace DocxportNet.Doc;

/// <summary>Resolves installed or explicitly supplied font files for advance measurement.</summary>
internal static class DocSystemFontAdvances
{
    private static readonly ConcurrentDictionary<string, Lazy<DocOpenTypeAdvances?>> Cache =
        new(StringComparer.OrdinalIgnoreCase);

    public static DocOpenTypeAdvances? Find(string? family, bool bold = false,
        bool italic = false)
    {
        if (string.IsNullOrWhiteSpace(family)) return null;
        var face = family + (bold && italic ? " Bold Italic" :
            bold ? " Bold" : italic ? " Italic" : "");
        var directory = Environment.GetEnvironmentVariable("DOCXPORT_FONT_DIRECTORY");
        return Cache.GetOrAdd((directory ?? "") + "|" + face, _ =>
            new Lazy<DocOpenTypeAdvances?>(() => Load(face, directory))).Value;
    }

    private static DocOpenTypeAdvances? Load(string family, string? overrideDirectory)
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
                var matchingName = key?.GetValueNames().FirstOrDefault(name =>
                    name.Equals(family + " (TrueType)", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals(family + " (OpenType)", StringComparison.OrdinalIgnoreCase));
                if (matchingName == null || key!.GetValue(matchingName) is not string file)
                    continue;
                var path = Path.IsPathRooted(file) ? file : Path.Combine(directory, file);
                if (File.Exists(path))
                    return new DocOpenTypeAdvances(File.ReadAllBytes(path));
            }
            return null;
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or
            PlatformNotSupportedException or InvalidDataException or NotSupportedException)
        {
            return null;
        }
    }
}
