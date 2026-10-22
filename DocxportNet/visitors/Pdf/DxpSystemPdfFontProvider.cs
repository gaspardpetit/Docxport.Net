using System.Collections.Concurrent;
using System.Diagnostics;

namespace DocxportNet.Visitors.Pdf;

/// <summary>Discovers installed fonts through fontconfig on Linux and macOS.
/// On Windows, <see cref="DxpPdfFonts"/> delegates to Pdfxport's platform resolver.</summary>
public sealed class DxpSystemPdfFontProvider : IDxpPdfFontProvider
{
    private readonly ConcurrentDictionary<(string Family, bool Bold, bool Italic), string> _resolved = new();

    public DxpPdfFontFace? Resolve(string familyName, bool bold, bool italic)
    {
        string style = bold && italic ? "Bold Italic" : bold ? "Bold" : italic ? "Italic" : "Regular";
        string path = _resolved.GetOrAdd((familyName, bold, italic), _ =>
            FindWithFontconfig(familyName, style) ?? FindFallbackFont() ?? string.Empty);
        return path.Length == 0 ? null : new DxpPdfFontFace(path, File.ReadAllBytes(path));
    }

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
