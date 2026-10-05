namespace DocxportNet.Doc;

/// <summary>Supplies a measured character scale when Word needs it to render DOC fit-text.</summary>
internal static class DocFitTextScaleResolver
{
    public static void Apply(List<DocPlainTextFormatRun> runs, string text,
        IReadOnlyList<DocPlainTextParagraphStyleRun> paragraphs,
        IReadOnlyList<DocStyleDefinition> styles, DocCharacterFormatting defaults,
        DocParagraphFormatting? paragraphDefaults)
    {
        if (runs.All(x => x.Formatting.FitText == null &&
            x.Formatting.CharacterStyleIndex == null)) return;
        var ordered = runs.Select((run, index) => (Run: run, Index: index))
            .OrderBy(x => x.Run.Start).ToArray();
        var styleByIndex = styles.ToDictionary(x => x.Index);
        T? InStyle<T>(int? index, Func<DocCharacterFormatting, T?> value)
            where T : struct
        {
            var visited = new HashSet<int>();
            while (index is int current && visited.Add(current) &&
                styleByIndex.TryGetValue(current, out var style))
            {
                if (value(style.CharacterFormatting) is T found) return found;
                index = style.BasedOnIndex;
            }
            return null;
        }
        string? StringInStyle(int? index, Func<DocCharacterFormatting, string?> value)
        {
            var visited = new HashSet<int>();
            while (index is int current && visited.Add(current) &&
                styleByIndex.TryGetValue(current, out var style))
            {
                if (value(style.CharacterFormatting) is string found) return found;
                index = style.BasedOnIndex;
            }
            return null;
        }
        DocFitText? Fit(DocPlainTextFormatRun run, int? paragraphStyle)
        {
            if (run.Formatting.FitText is { } direct) return direct;
            var visited = new HashSet<int>();
            foreach (var first in new[] { run.Formatting.CharacterStyleIndex, paragraphStyle })
            {
                var index = first;
                while (index is int current && visited.Add(current) &&
                    styleByIndex.TryGetValue(current, out var style))
                {
                    if (style.CharacterFormatting.FitText is { } fit) return fit;
                    index = style.BasedOnIndex;
                }
            }
            return defaults.FitText;
        }
        DocPlainTextParagraphStyleRun? Paragraph(int cp) => paragraphs
            .LastOrDefault(x => x.Start <= cp && cp < x.End);
        bool AutoSpace(DocPlainTextParagraphStyleRun? paragraph,
            Func<DocParagraphFormatting, bool?> value)
        {
            if (paragraph?.Formatting is { } direct && value(direct) is bool enabled)
                return enabled;
            var visited = new HashSet<int>();
            var index = paragraph?.StyleIndex;
            while (index is int current && visited.Add(current) &&
                styleByIndex.TryGetValue(current, out var style))
            {
                if (style.ParagraphFormatting is { } formatting &&
                    value(formatting) is bool inherited)
                    return inherited;
                index = style.BasedOnIndex;
            }
            return paragraphDefaults is { } defaults && value(defaults) is bool fallback
                ? fallback : true;
        }
        for (var start = 0; start < ordered.Length;)
        {
            var paragraph = Paragraph(ordered[start].Run.Start);
            var fit = Fit(ordered[start].Run, paragraph?.StyleIndex);
            if (fit is not { WidthTwips: > 0 }) { start++; continue; }
            var end = start + 1;
            while (end < ordered.Length &&
                ordered[end - 1].Run.End == ordered[end].Run.Start &&
                Fit(ordered[end].Run, Paragraph(ordered[end].Run.Start)?.StyleIndex) == fit &&
                !text.Substring(ordered[end - 1].Run.Start,
                    ordered[end - 1].Run.End - ordered[end - 1].Run.Start)
                    .Any(ch => ch is '\r' or '\f' or '\u0007'))
                end++;
            var chunk = text.Substring(ordered[start].Run.Start,
                ordered[end - 1].Run.End - ordered[start].Run.Start);
            if (chunk.Length == 0 || chunk.Any(ch => ch is '\r' or '\f' or '\u0007') ||
                Enumerable.Range(start, end - start).Any(i =>
                    ordered[i].Run.Formatting.CharacterScalePercent != null))
            { start = end; continue; }
            var measured = new List<DocFitTextMeasureRun>();
            for (var i = start; i < end; i++)
            {
                var run = ordered[i].Run;
                var paraStyle = Paragraph(run.Start)?.StyleIndex;
                var characterStyle = run.Formatting.CharacterStyleIndex;
                var ascii = run.Formatting.AsciiFontName ??
                    StringInStyle(characterStyle, x => x.AsciiFontName) ??
                    StringInStyle(paraStyle, x => x.AsciiFontName) ?? defaults.AsciiFontName;
                var eastAsian = run.Formatting.EastAsiaFontName ??
                    StringInStyle(characterStyle, x => x.EastAsiaFontName) ??
                    StringInStyle(paraStyle, x => x.EastAsiaFontName) ??
                    defaults.EastAsiaFontName ?? ascii;
                var runSize = run.Formatting.SizeHalfPoints ??
                    InStyle(characterStyle, x => x.SizeHalfPoints) ??
                    InStyle(paraStyle, x => x.SizeHalfPoints) ??
                    defaults.SizeHalfPoints ?? (ushort)20;
                var highAnsi = run.Formatting.HighAnsiFontName ??
                    StringInStyle(characterStyle, x => x.HighAnsiFontName) ??
                    StringInStyle(paraStyle, x => x.HighAnsiFontName) ??
                    defaults.HighAnsiFontName ?? ascii;
                bool Weight(Func<DocCharacterFormatting, bool?> property,
                    Func<DocCharacterFormatting, bool?> toggle)
                {
                    var inherited = InStyle(characterStyle, property) ??
                        InStyle(paraStyle, property) ?? property(defaults) ?? false;
                    if (toggle(run.Formatting) is bool invert) return inherited ^ invert;
                    return property(run.Formatting) ?? inherited;
                }
                var bold = Weight(x => x.Bold, x => x.BoldStyleToggle);
                var italic = Weight(x => x.Italic, x => x.ItalicStyleToggle);
                measured.Add(new DocFitTextMeasureRun(text.Substring(run.Start,
                    run.End - run.Start), ascii, eastAsian, runSize / 2.0, highAnsi,
                    bold, italic));
            }
            if (DocFitTextScaleEstimator.TryEstimate(fit.WidthTwips, measured,
                    AutoSpace(paragraph, x => x.AutoSpaceDE),
                    AutoSpace(paragraph, x => x.AutoSpaceDN), out var scale,
                    out var spacingTwips, out var finalSpacingTwips))
            {
                bool CharacterStyleHasFit(int? index)
                {
                    var visited = new HashSet<int>();
                    while (index is int current && visited.Add(current) &&
                        styleByIndex.TryGetValue(current, out var style))
                    {
                        if (style.CharacterFormatting.FitText != null) return true;
                        index = style.BasedOnIndex;
                    }
                    return false;
                }
                var useSpacing = Enumerable.Range(start, end - start).Any(i =>
                    ordered[i].Run.Formatting is { FitText: not null,
                        CharacterStyleIndex: not null } direct &&
                    !CharacterStyleHasFit(direct.CharacterStyleIndex));
                if (useSpacing && Enumerable.Range(start, end - start).Any(i =>
                    ordered[i].Run.Formatting.CharacterSpacingTwips != null))
                { start = end; continue; }
                for (var i = start; i < end; i++)
                {
                    var (run, index) = ordered[i];
                    runs[index] = run with { Formatting = useSpacing
                        ? run.Formatting with { CharacterSpacingTwips = spacingTwips }
                        : run.Formatting with { CharacterScalePercent = scale } };
                }
                if (useSpacing && finalSpacingTwips != spacingTwips)
                {
                    var (last, index) = ordered[end - 1];
                    var final = runs[index] with
                    {
                        Start = last.End - 1,
                        Formatting = runs[index].Formatting with
                            { CharacterSpacingTwips = finalSpacingTwips }
                    };
                    if (last.End - last.Start > 1)
                    {
                        runs[index] = runs[index] with { End = last.End - 1 };
                        runs.Add(final);
                    }
                    else runs[index] = final;
                }
            }
            start = end;
        }
    }
}
