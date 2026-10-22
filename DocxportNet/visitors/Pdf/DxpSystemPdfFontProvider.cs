using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;

namespace DocxportNet.Visitors.Pdf;

/// <summary>Discovers installed OpenType fonts and checks their glyph coverage.</summary>
public sealed class DxpSystemPdfFontProvider : IDxpPdfFontProvider
{
    private readonly Lazy<DxpDirectoryPdfFontProvider> _catalog = new(() =>
        new DxpDirectoryPdfFontProvider(SystemFontDirectories()));
    private readonly ConcurrentDictionary<(string Family, bool Bold, bool Italic), string> _fontconfig = new();

    /// <summary>Installed family names known to the default PDF renderer.</summary>
    public IReadOnlyCollection<string> Families => _catalog.Value.Families;

    /// <summary>An installed font with broad East Asian coverage for runs without an
    /// explicit East Asian face. Document-specified fonts still take precedence.</summary>
    public string? PreferredCjkFamily => new[] { "Noto Sans JP", "Noto Sans CJK JP", "Yu Mincho",
        "Hiragino Sans", "Microsoft YaHei" }
        .FirstOrDefault(family => _catalog.Value.IsAvailable(family, false, false));

    public bool IsAvailable(string familyName, bool bold, bool italic) =>
        _catalog.Value.IsAvailable(familyName, bold, italic);

    public DxpPdfFontFace? Resolve(string familyName, bool bold, bool italic)
    {
        var face = _catalog.Value.Resolve(familyName, bold, italic);
        if (face != null) return face;
        // Fontconfig also searches font locations outside the standard directories.
        if (OperatingSystem.IsWindows()) return null;
        string path = _fontconfig.GetOrAdd((familyName, bold, italic), key =>
            FindWithFontconfig(key.Family, key.Bold, key.Italic) ?? string.Empty);
        return path.Length == 0 ? null : new DxpPdfFontFace(path, File.ReadAllBytes(path));
    }

    public DxpPdfFontMatch? ResolveForCharacter(string familyName, bool bold, bool italic, Rune character) =>
        _catalog.Value.ResolveForCharacter(familyName, bold, italic, character) ??
        _catalog.Value.ResolveAnyForCharacter(bold, italic, character);

    private static IEnumerable<string> SystemFontDirectories()
    {
        if (OperatingSystem.IsWindows())
        {
            string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            if (windows.Length > 0) yield return Path.Combine(windows, "Fonts");
            string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (local.Length > 0)
            {
                yield return Path.Combine(local, "Microsoft", "Windows", "Fonts");
                yield return Path.Combine(local, "Microsoft", "FontCache");
            }
        }
        else if (OperatingSystem.IsMacOS())
        {
            yield return "/System/Library/Fonts";
            yield return "/Library/Fonts";
            yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library", "Fonts");
        }
        else
        {
            yield return "/usr/share/fonts";
            yield return "/usr/local/share/fonts";
            string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            yield return Path.Combine(home, ".local", "share", "fonts");
            yield return Path.Combine(home, ".fonts");
        }
    }

    private static string? FindWithFontconfig(string family, bool bold, bool italic)
    {
        try
        {
            string style = bold && italic ? "Bold Italic" : bold ? "Bold" : italic ? "Italic" : "Regular";
            var start = new ProcessStartInfo("fc-match")
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            start.ArgumentList.Add("--format=%{file}\n");
            start.ArgumentList.Add($"{family}:style={style}");
            using var process = Process.Start(start);
            if (process == null) return null;
            string path = process.StandardOutput.ReadLine()?.Trim() ?? "";
            process.WaitForExit();
            return process.ExitCode == 0 && File.Exists(path) ? path : null;
        }
        catch (System.ComponentModel.Win32Exception) { return null; }
    }
}
