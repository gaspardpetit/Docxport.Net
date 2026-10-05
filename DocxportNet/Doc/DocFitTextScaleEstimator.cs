namespace DocxportNet.Doc;

internal sealed record DocFitTextMeasureRun(string Text, string? AsciiFont,
    string? EastAsiaFont, double FontSizePoints, string? HighAnsiFont = null,
    bool Bold = false, bool Italic = false);

/// <summary>Estimates a DOC character-scale fallback for Word's fitted text region.</summary>
internal static class DocFitTextScaleEstimator
{
    public static bool TryEstimate(int widthTwips, string text, string? fontFamily,
        double fontSizePoints, bool autoSpaceEastAsianLatin,
        bool autoSpaceEastAsianNumbers, out ushort scale) =>
        TryEstimate(widthTwips,
            [new DocFitTextMeasureRun(text, fontFamily, fontFamily, fontSizePoints)],
            autoSpaceEastAsianLatin, autoSpaceEastAsianNumbers, out scale);

    public static bool TryEstimate(int widthTwips,
        IReadOnlyList<DocFitTextMeasureRun> runs,
        bool autoSpaceEastAsianLatin, bool autoSpaceEastAsianNumbers,
        out ushort scale) => TryEstimate(widthTwips, runs,
            autoSpaceEastAsianLatin, autoSpaceEastAsianNumbers, out scale, out _, out _);

    public static bool TryEstimate(int widthTwips,
        IReadOnlyList<DocFitTextMeasureRun> runs,
        bool autoSpaceEastAsianLatin, bool autoSpaceEastAsianNumbers,
        out ushort scale, out short spacingTwips,
        out short finalSpacingTwips)
    {
        scale = 0;
        spacingTwips = 0;
        finalSpacingTwips = 0;
        if (widthTwips <= 0 || runs.Count == 0) return false;
        double naturalWidth = 0;
        double letterSpacing = 0;
        double numberSpacing = 0;
        var faces = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        char previous = '\0';
        double previousSize = 0;
        foreach (var run in runs)
        {
            if (run.FontSizePoints <= 0) return false;
            var cursor = 0;
            while (cursor < run.Text.Length)
            {
                string? Family(char ch) => IsEastAsian(ch) ? run.EastAsiaFont :
                    ch > 127 ? run.HighAnsiFont ?? run.AsciiFont : run.AsciiFont;
                var family = Family(run.Text[cursor]);
                if (family == null || DocSystemFontAdvances.Find(family, run.Bold, run.Italic) is not { } metrics)
                    return false;
                var end = cursor + 1;
                while (end < run.Text.Length &&
                    string.Equals(Family(run.Text[end]), family,
                        StringComparison.OrdinalIgnoreCase))
                    end++;
                var part = run.Text.Substring(cursor, end - cursor);
                if (!metrics.TryMeasure(part, run.FontSizePoints, out var advance))
                    return false;
                naturalWidth += advance;
                faces.Add(family);
                for (var i = cursor; i < end; i++)
                {
                    var ch = run.Text[i];
                    var spacing = (previousSize + run.FontSizePoints) / 8;
                    if (IsEastAsian(previous) && IsLatinLetter(ch) ||
                        IsLatinLetter(previous) && IsEastAsian(ch))
                        letterSpacing += spacing;
                    if (IsEastAsian(previous) && IsAsciiDigit(ch) ||
                        IsAsciiDigit(previous) && IsEastAsian(ch))
                        numberSpacing += spacing;
                    previous = ch;
                    previousSize = run.FontSizePoints;
                }
                cursor = end;
            }
        }
        // Word's native DOC fit scale groups numeric boundaries with Latin ones
        // when a single face serves the whole region. Across different faces,
        // the paragraph's DE and DN controls contribute separately.
        if (faces.Count == 1)
        {
            if (autoSpaceEastAsianLatin)
                naturalWidth += letterSpacing + numberSpacing;
        }
        else
        {
            if (autoSpaceEastAsianLatin) naturalWidth += letterSpacing;
            if (autoSpaceEastAsianNumbers) naturalWidth += numberSpacing;
        }
        if (naturalWidth <= 0) return false;
        var percentage = Math.Round(widthTwips / 20.0 / naturalWidth * 100,
            MidpointRounding.AwayFromZero);
        if (percentage is < 1 or > 600) return false;
        scale = (ushort)percentage;
        var count = runs.Sum(run => run.Text.Length);
        if (count > 0)
        {
            var residual = Math.Round(widthTwips - naturalWidth * 20,
                MidpointRounding.AwayFromZero);
            var spacing = Math.Round(residual / count,
                MidpointRounding.AwayFromZero);
            var final = residual - spacing * (count - 1);
            if (spacing is < short.MinValue or > short.MaxValue ||
                final is < short.MinValue or > short.MaxValue) return false;
            spacingTwips = (short)spacing;
            finalSpacingTwips = (short)final;
        }
        return true;
    }

    private static bool IsLatinLetter(char ch) => ch is >= 'A' and <= 'Z' or
        >= 'a' and <= 'z';

    private static bool IsAsciiDigit(char ch) => ch is >= '0' and <= '9';

    private static bool IsEastAsian(char ch) => ch is
        >= '\u3000' and <= '\u303F' or // CJK symbols and punctuation
        >= '\u3040' and <= '\u30FF' or // Japanese kana
        >= '\u3400' and <= '\u9FFF' or // CJK ideographs
        >= '\uAC00' and <= '\uD7AF';  // Hangul syllables
}
