using System.Collections.Concurrent;
using System.Diagnostics;
using PdfSharp.Fonts;

internal sealed class DxpCliSystemFontResolver : IFontResolver
{
    private readonly ConcurrentDictionary<(string Family, bool Bold, bool Italic), string> _resolved = new();
    private readonly ConcurrentDictionary<string, string> _paths = new(StringComparer.Ordinal);

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        string style = isBold && isItalic ? "Bold Italic" : isBold ? "Bold" : isItalic ? "Italic" : "Regular";
        string path = _resolved.GetOrAdd((familyName, isBold, isItalic), _ =>
            FindWithFontconfig(familyName, style) ?? FindFallbackFont() ??
            throw new InvalidOperationException("No usable system font found. Install fontconfig and a TrueType font."));
        _paths.TryAdd(path, path);
        return new FontResolverInfo(path, false, false);
    }

    public byte[] GetFont(string faceName) => File.ReadAllBytes(_paths[faceName]);

    private static string? FindWithFontconfig(string familyName, string style)
    {
        try
        {
            var start = new ProcessStartInfo("fc-match")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            start.ArgumentList.Add("--format=%{file}\n");
            start.ArgumentList.Add($"{familyName}:style={style}");
            using var process = Process.Start(start);
            if (process == null)
                return null;
            string path = process.StandardOutput.ReadLine()?.Trim() ?? "";
            process.WaitForExit();
            return process.ExitCode == 0 && File.Exists(path) ? path : null;
        }
        catch (System.ComponentModel.Win32Exception)
        {
            return null;
        }
    }

    private static string? FindFallbackFont()
    {
        foreach (string path in new[]
        {
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
            "/Library/Fonts/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Arial.ttf"
        })
        {
            if (File.Exists(path))
                return path;
        }
        return null;
    }
}
