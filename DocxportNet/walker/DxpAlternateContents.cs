using DocumentFormat.OpenXml;

namespace DocxportNet.Walker;

public class DxpAlternateContents
{
    // Helper: split the whitespace-delimited Requires value into prefixes.
    public static IReadOnlyList<string> GetRequiredPrefixes(AlternateContentChoice ch)
    {
        string? val = ch.Requires?.Value; // <-- correct source for mc:Requires
        if (string.IsNullOrWhiteSpace(val))
            return [];

        char[]? NullSeparator = null!;
        return val!.Split(NullSeparator, StringSplitOptions.RemoveEmptyEntries);
    }

}
