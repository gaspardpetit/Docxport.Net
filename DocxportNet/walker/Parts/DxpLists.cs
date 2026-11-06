using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Formatting;
using DocxportNet.Walker.Context;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace DocxportNet.Walker;


public sealed class DxpNumberingResolver
{
    private readonly Numbering? _numbering;
    private readonly Dictionary<int, NumberingInstance> _numById = new();
    private readonly Dictionary<int, AbstractNum> _absById = new();
    private readonly Dictionary<string, Style> _styleById = new(StringComparer.Ordinal);

    public DxpNumberingResolver(WordprocessingDocument doc)
    {
        _numbering = doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        if (_numbering != null)
        {
            foreach (var n in _numbering.Elements<NumberingInstance>())
                if (n.NumberID?.Value != null)
                    _numById[n.NumberID.Value] = n;

            foreach (var a in _numbering.Elements<AbstractNum>())
                if (a.AbstractNumberId?.Value != null)
                    _absById[a.AbstractNumberId.Value] = a;
        }

        var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles != null)
        {
            foreach (var s in styles.Elements<Style>())
            {
                var id = s.StyleId?.Value;
                if (!string.IsNullOrEmpty(id))
                    _styleById[id!] = s;
            }
        }
    }

    public (AbstractNum abs, Level lvl, int startAt)? ResolveLevel(int numId, int ilvl)
    {
        if (!_numById.TryGetValue(numId, out var num))
            return null;

        var absId = num.AbstractNumId?.Val?.Value;
        if (absId == null || !_absById.TryGetValue(absId.Value, out var abs))
            return null;

        // IMPORTANT: some abstractNum are just a link (numStyleLink/styleLink) and contain no <w:lvl>
        abs = FollowAbstractNumLinks(abs);
        if (abs == null) return null;

        // clamp + Word-like fallback to level 0 when missing
        if (ilvl < 0)
            ilvl = 0;
        if (ilvl > 8)
            ilvl = 8;

        var abstractLevel = abs.Elements<Level>().FirstOrDefault(x => x.LevelIndex?.Value == ilvl);

        var ov = num.Elements<LevelOverride>()
                    .FirstOrDefault(o => o.LevelIndex?.Value == ilvl);
        var replacement = ov?.GetFirstChild<Level>();
        var lvl = replacement ?? abstractLevel ?? new Level(new LevelText { Val = "" }) { LevelIndex = ilvl };
        // At the root, Word activates the replacement start only with startOverride.
        // At deeper levels its start takes precedence even without startOverride.
        var overrideStart = ov?.StartOverrideNumberingValue?.Val?.Value;
        var startAt = ilvl == 0 && overrideStart == null
            ? abstractLevel?.StartNumberingValue?.Val?.Value ?? 0
            : replacement?.StartNumberingValue?.Val?.Value
            ?? ov?.StartOverrideNumberingValue?.Val?.Value
            // A replacement without a start has its own default zero;
            // only the absence of a replacement inherits the abstract start.
            ?? (replacement != null ? 0 : abstractLevel?.StartNumberingValue?.Val?.Value ?? 0);

        return (abs, lvl, startAt);
    }

    public bool HasStartOverride(int numId, int ilvl) =>
        _numById.TryGetValue(numId, out var num) &&
        num.Elements<LevelOverride>().Any(level => level.LevelIndex?.Value == ilvl &&
            level.StartOverrideNumberingValue?.Val?.Value != null);

    public int GetCounterAbstractId(int numId) =>
        _numById.TryGetValue(numId, out var num) ? num.AbstractNumId?.Val?.Value ?? numId : numId;

    public int GetRestartTrigger(int numId, int ilvl)
    {
        var abs = ResolveLevel(numId, ilvl)?.abs;
        // Word ignores lvlRestart inside a replacement level.
        var level = abs?.Elements<Level>().FirstOrDefault(x => x.LevelIndex?.Value == ilvl);
        var restart = level?.GetFirstChild<LevelRestart>()?.Val?.Value;
        // OOXML uses a one-based level index; zero disables automatic restart.
        return restart == null ? ilvl - 1 : restart.Value - 1;
    }

    internal AbstractNum? FollowAbstractNumLinks(AbstractNum abs)
    {
        // If it already has levels, nothing to do.
        if (abs.Elements<Level>().Any())
            return abs;
        if (abs.StyleLink != null)
            return null; // A style-associated definition must supply its own levels in Word.

        // 1) numStyleLink (abstractNumId may point to a linked style without levels)
        var numStyleLinkId = abs.NumberingStyleLink?.Val?.Value; // maps to <w:numStyleLink>
        if (!string.IsNullOrEmpty(numStyleLinkId))
        {
            var linked = ResolveAbstractFromNumberingStyleId(numStyleLinkId!);
            if (linked?.StyleLink?.Val?.Value == numStyleLinkId && linked.Elements<Level>().Any())
                return linked;
        }

        // styleLink is a reverse association, not a forward numbering reference.
        // Word does not resolve multi-hop abstract links or a missing reverse link.
        return string.IsNullOrEmpty(numStyleLinkId) ? abs : null;
    }

    private AbstractNum? ResolveAbstractFromNumberingStyleId(string numberingStyleId)
    {
        // styles.xml: numbering style PatentSpecificationList has numId=2
        if (!_styleById.TryGetValue(numberingStyleId, out var style))
            return null;

        int? numId = null;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (visited.Add(style.StyleId?.Value ?? ""))
        {
            numId = style.StyleParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
            if (numId != null)
                break;
            var parent = style.BasedOn?.Val?.Value;
            if (parent == null || !_styleById.TryGetValue(parent, out style))
                break;
        }

        if (numId == null || numId.Value == 0)
            return null;

        if (!_numById.TryGetValue(numId.Value, out var num))
            return null;

        var absId = num.AbstractNumId?.Val?.Value;
        if (absId == null)
            return null;

        return _absById.TryGetValue(absId.Value, out var abs) ? abs : null;
    }
}



public sealed class DxpListTracker
{
    private readonly Dictionary<int, int?[]> _counters = new();
    private readonly Dictionary<int, int?[]> _activeCounters = new();
    private readonly Dictionary<int, bool[]> _referenceLevels = new();
    private readonly HashSet<int> _visitedInstances = new();
    private readonly HashSet<int> _appliedStartOverrideInstances = new();
    private readonly HashSet<(int NumId, int Level)> _startedLevels = new();
    private readonly Dictionary<int, int?[]> _restartStarts = new();
    private readonly HashSet<(int NumId, int Level)> _seededAncestorLevels = new();
    private readonly HashSet<(int NumId, int Level)> _consumedAncestorOverrides = new();
    internal DxpListTracker CopyForReference()
    {
        var copy = new DxpListTracker();
        foreach (var entry in _counters) copy._counters.Add(entry.Key, (int?[])entry.Value.Clone());
        foreach (var entry in _activeCounters) copy._activeCounters.Add(entry.Key, (int?[])entry.Value.Clone());
        foreach (var entry in _referenceLevels) copy._referenceLevels.Add(entry.Key, (bool[])entry.Value.Clone());
        copy._visitedInstances.UnionWith(_visitedInstances);
        copy._appliedStartOverrideInstances.UnionWith(_appliedStartOverrideInstances);
        copy._startedLevels.UnionWith(_startedLevels);
        foreach (var entry in _restartStarts) copy._restartStarts.Add(entry.Key, (int?[])entry.Value.Clone());
        copy._seededAncestorLevels.UnionWith(_seededAncestorLevels);
        copy._consumedAncestorOverrides.UnionWith(_consumedAncestorOverrides);
        return copy;
    }

    private int?[] GetArr(int abstractId)
    {
        if (!_counters.TryGetValue(abstractId, out var arr))
        {
            arr = Enumerable.Repeat<int?>(null, 9).ToArray();
            _counters[abstractId] = arr;
        }
        return arr;
    }

    public int NextIndex(int numId, int abstractId, int ilvl, int startAt, bool hasStartOverride,
        IReadOnlyList<int>? restartTriggers = null)
        => NextIndex(numId, abstractId, ilvl, startAt, hasStartOverride, restartTriggers, null);

    internal int NextIndex(int numId, int abstractId, int ilvl, int startAt, bool hasStartOverride,
        IReadOnlyList<int>? restartTriggers, IReadOnlyList<int>? restartStarts,
        IReadOnlyList<bool>? ancestorStartOverrides = null,
        IReadOnlyList<int>? declaredStarts = null)
    {
        var arr = GetArr(abstractId);
        if (!_activeCounters.TryGetValue(abstractId, out var active))
            _activeCounters[abstractId] = active = Enumerable.Repeat<int?>(null, 9).ToArray();
        // A heading can establish an instance before its first body paragraph.
        // The override applies when the instance first appears, not per level.
        var firstInInstance = _visitedInstances.Add(numId);
        var key = (numId, ilvl);
        var firstAtLevel = _startedLevels.Add(key);
        var seededAncestor = _seededAncestorLevels.Remove(key);
        var consumedAncestorOverride = _consumedAncestorOverrides.Remove(key);
        // Word shares progress across instances of the same abstract list,
        // including returns to instances used before an explicit restart.
        // A start override changes that shared progress when first applied;
        // it does not create a private counter for the restarted instance.
        // Word wraps signed 32-bit counters, even when callers enable checked arithmetic.
        // Replacement and override starts initialize the instance's level.
        // After an ancestor clears an already-used level, Word uses the start
        // belonging to that ancestor's instance, including across instances.
        var restartStartAt = _restartStarts.TryGetValue(abstractId, out var pending)
            ? pending[ilvl] ?? startAt : startAt;
        var applyStartOverride = hasStartOverride &&
            (firstInInstance || seededAncestor && firstAtLevel && !consumedAncestorOverride ||
                arr[ilvl] == null && firstAtLevel &&
                    !(restartTriggers != null && ilvl < restartTriggers.Count && restartTriggers[ilvl] == -1 &&
                        _appliedStartOverrideInstances.Contains(numId)));
        if (applyStartOverride) _appliedStartOverrideInstances.Add(numId);
        ResetDescendants(numId, abstractId, ilvl, restartTriggers, restartStarts,
            applyStartOverride ? declaredStarts : null);
        arr[ilvl] = applyStartOverride ? startAt : arr[ilvl] == null ? restartStartAt
            : unchecked(arr[ilvl]!.Value + 1);
        // Applying a child's override also consumes unvisited ancestor overrides.
        // Their future counters use effective starts, while the active placeholder
        // keeps its existing/abstract display until the ancestor actually appears.
        // A skipped never-restarting level below an actual ancestor is different:
        // its unconsumed start-1 seed survives, although its override is consumed.
        if (applyStartOverride && ancestorStartOverrides != null && restartStarts != null)
        {
            for (var ancestor = 0; ancestor < ilvl; ancestor++)
            {
                if (ancestor >= ancestorStartOverrides.Count || !ancestorStartOverrides[ancestor] ||
                    ancestor >= restartStarts.Count || _startedLevels.Contains((numId, ancestor))) continue;
                _consumedAncestorOverrides.Add((numId, ancestor));
                var never = restartTriggers != null && ancestor < restartTriggers.Count && restartTriggers[ancestor] == -1;
                var hasActualAncestors = _referenceLevels.TryGetValue(abstractId, out var actualAncestors);
                var alreadyActual = hasActualAncestors && actualAncestors![ancestor];
                var belowActual = hasActualAncestors && actualAncestors!.Take(ancestor).Any(value => value);
                // Another instance can already own the shared actual ancestor.
                // Consume this instance's override without replacing that progress.
                if (!alreadyActual && (!never || !belowActual)) arr[ancestor] = restartStarts[ancestor];
            }
        }
        active[ilvl] = arr[ilvl];
        if (!_referenceLevels.TryGetValue(abstractId, out var visited))
            _referenceLevels[abstractId] = visited = new bool[9];
        visited[ilvl] = true;
        return arr[ilvl]!.Value;
    }

    public void ObserveNoDisplayLevel(int numId, int abstractId, int ilvl,
        IReadOnlyList<int>? restartTriggers = null) =>
        ResetDescendants(numId, abstractId, ilvl, restartTriggers);

    private void ResetDescendants(int numId, int abstractId, int ilvl,
        IReadOnlyList<int>? restartTriggers, IReadOnlyList<int>? restartStarts = null,
        IReadOnlyList<int>? declaredStarts = null)
    {
        var arr = GetArr(abstractId);
        _referenceLevels.TryGetValue(abstractId, out var visited);
        if (!_activeCounters.TryGetValue(abstractId, out var active))
            _activeCounters[abstractId] = active = Enumerable.Repeat<int?>(null, 9).ToArray();
        for (int d = ilvl + 1; d < arr.Length; d++)
        {
            // A new ancestor starts a new paragraph chain even when lvlRestart
            // keeps the descendant's numeric counter running.
            if (visited != null) visited[d] = false;
            var trigger = restartTriggers != null && d < restartTriggers.Count
                ? restartTriggers[d] : d - 1;
            if (!_restartStarts.TryGetValue(abstractId, out var pending))
                _restartStarts[abstractId] = pending = new int?[9];
            if (trigger < ilvl)
            {
                active[d] = null;
                // Even a never-restarting level acquires its initial start
                // from the first ancestor's instance, before any child visit.
                if (pending[d] == null && restartStarts != null && d < restartStarts.Count)
                    // A never-restarting child first reached after this instance
                    // began at an ancestor has skipped its initial override.
                    // Replacement-level starts still belong to the level itself.
                    pending[d] = trigger == -1 && declaredStarts != null && d < declaredStarts.Count
                        ? declaredStarts[d] : restartStarts[d];
                continue;
            }
            active[d] = null;
            arr[d] = null;
            pending[d] = restartStarts != null && d < restartStarts.Count ? restartStarts[d] : null;
        }
    }

    public int SeedUnstartedAncestor(int numId, int abstractId, int ilvl, int abstractStart)
        => SeedUnstartedAncestor(numId, abstractId, ilvl, abstractStart, restartNever: false);

    internal int SeedUnstartedAncestor(int numId, int abstractId, int ilvl, int abstractStart,
        bool restartNever, int? effectiveStart = null)
    {
        _seededAncestorLevels.Add((numId, ilvl));
        var arr = GetArr(abstractId);
        var belowActualAncestor = _referenceLevels.TryGetValue(abstractId, out var visited) &&
            visited.Take(ilvl).Any(value => value);
        if (arr[ilvl] == null)
        {
            // Under an actual ancestor, a skipped never-restarting level has not
            // consumed its first value. Word exposes start-1 in child placeholders
            // and advances to start on that level's first actual paragraph.
            arr[ilvl] = restartNever && belowActualAncestor ? abstractStart - 1
                : _startedLevels.Contains((numId, ilvl)) || _consumedAncestorOverrides.Contains((numId, ilvl))
                    ? effectiveStart ?? abstractStart : abstractStart;
        }
        if (!_activeCounters.TryGetValue(abstractId, out var active))
            _activeCounters[abstractId] = active = Enumerable.Repeat<int?>(null, 9).ToArray();
        var displayStart = _startedLevels.Contains((numId, ilvl)) &&
            _restartStarts.TryGetValue(abstractId, out var pending) ? pending[ilvl] ?? abstractStart : abstractStart;
        active[ilvl] ??= restartNever && belowActualAncestor ? abstractStart - 1 : displayStart;
        return active[ilvl]!.Value;
    }

    public int? GetActiveCurrent(int abstractId, int ilvl) =>
        _activeCounters.TryGetValue(abstractId, out var active) ? active[ilvl] : null;

    public int? GetCurrent(int numId, int abstractId, int ilvl) => GetArr(abstractId)[ilvl];

    public void ClearAll()
    {
        _counters.Clear();
        _activeCounters.Clear();
        _referenceLevels.Clear();
        _visitedInstances.Clear();
        _appliedStartOverrideInstances.Clear();
        _startedLevels.Clear();
        _restartStarts.Clear();
        _seededAncestorLevels.Clear();
        _consumedAncestorOverrides.Clear();
    }
}

public static class DxpListMarkerFormatter
{
    private static int? TryGetCustomDecimalWidth(Level lvl)
    {
        // The "custom" numFmt is in mc:AlternateContent and may not be surfaced by the SDK as lvl.NumberingFormat.
        // So parse from the raw XML.
        var xml = lvl.OuterXml;

        // Look for: w:numFmt w:val="custom" w:format="0001, 0002, ..."
        const string key = "w:format=\"";
        var idx = xml.IndexOf("w:numFmt", StringComparison.Ordinal);
        if (idx < 0)
            return null;

        // quick/robust-enough scan
        var customIdx = xml.IndexOf("w:val=\"custom\"", idx, StringComparison.Ordinal);
        if (customIdx < 0)
            return null;

        var fmtIdx = xml.IndexOf(key, customIdx, StringComparison.Ordinal);
        if (fmtIdx < 0)
            return null;

        fmtIdx += key.Length;
        var end = xml.IndexOf('"', fmtIdx);
        if (end < 0)
            return null;

        var format = xml.Substring(fmtIdx, end - fmtIdx); // e.g. "0001, 0002, 0003, ."
                                                          // Take the first token up to comma/space; "0001" => width 4
        var token = format.Split(',', ' ', ';').FirstOrDefault(t => t.Length > 0);
        if (string.IsNullOrEmpty(token))
            return null;

        // Count leading zeros in first token (0001 => width 4). If no zeros, ignore.
        var width = token.Length;
        if (width <= 1)
            return null;

        // Ensure it's a numeric mask like 0001
        if (token.Any(ch => ch != '0' && ch != '1' && ch != '2' && ch != '3' && ch != '4' && ch != '5' && ch != '6' && ch != '7' && ch != '8' && ch != '9'))
            return null;

        return width;
    }

    public static string? TryBuildMarkerText(
        DxpStyleEffectiveNumPr numPr,
        DxpNumberingResolver nr,
        DxpListTracker tracker,
        CultureInfo? culture = null)
    {
        var resolved = nr.ResolveLevel(numPr.NumId, numPr.Ilvl);
        if (resolved == null)
            return null;

        var (abs, lvl, startAt) = resolved.Value;

        var fmt = lvl.NumberingFormat?.Val?.Value;     // w:numFmt (may be null)
        var lvlText = lvl.LevelText?.Val?.Value ?? ""; // w:lvlText ("" is meaningful!)
        // Word stores at most 31 label units; each of its first nine numeric
        // references occupies one unit, regardless of the expanded number width.
        var labelEnd = 0;
        var referenceCount = 0;
        for (var units = 0; units < 31 && labelEnd < lvlText.Length; units++)
        {
            if (lvlText[labelEnd] == '%' && labelEnd + 1 < lvlText.Length &&
                lvlText[labelEnd + 1] is >= '0' and <= '9' && referenceCount < 9)
            {
                referenceCount++;
                labelEnd++;
            }
            labelEnd++;
        }
        lvlText = lvlText.Substring(0, labelEnd);
        var references = System.Text.RegularExpressions.Regex.Matches(lvlText, "%[0-9]")
            .Cast<System.Text.RegularExpressions.Match>().Take(9).ToArray();

        // Increment current level (normal list item / bullet / numbered)
        var abstractId = nr.GetCounterAbstractId(numPr.NumId);
        var restartTriggers = Enumerable.Range(0, 9)
            .Select(level => nr.GetRestartTrigger(numPr.NumId, level)).ToArray();
        var effectiveStarts = Enumerable.Range(0, 9)
            .Select(level => nr.ResolveLevel(numPr.NumId, level)?.startAt ?? 0).ToArray();
        var ancestorStartOverrides = Enumerable.Range(0, 9)
            .Select(level => nr.HasStartOverride(numPr.NumId, level)).ToArray();
        var declaredStarts = Enumerable.Range(0, 9)
            .Select(level => nr.ResolveLevel(numPr.NumId, level)?.lvl.StartNumberingValue?.Val?.Value ?? 0).ToArray();
        // Before an actual ancestor establishes the chain, an unactivated
        // replacement uses the abstract start for its first visible item.
        if (!nr.HasStartOverride(numPr.NumId, numPr.Ilvl) && numPr.Ilvl > 0 &&
            tracker.GetActiveCurrent(abstractId, numPr.Ilvl - 1) == null &&
            abs.Elements<Level>().Any(level => level.LevelIndex?.Value == numPr.Ilvl - 1))
            startAt = abs.Elements<Level>().FirstOrDefault(level => level.LevelIndex?.Value == numPr.Ilvl)
                ?.StartNumberingValue?.Val?.Value ?? startAt;
        tracker.NextIndex(numPr.NumId, abstractId, numPr.Ilvl, startAt,
            nr.HasStartOverride(numPr.NumId, numPr.Ilvl), restartTriggers,
            effectiveStarts, ancestorStartOverrides, declaredStarts);

        // Word establishes defined ancestors on visiting a child, even when
        // the child's label has no placeholder referring to those ancestors.
        for (var ancestorIndex = numPr.Ilvl - 1; ancestorIndex >= 0; ancestorIndex--)
        {
            var ancestor = abs.Elements<Level>().FirstOrDefault(level => level.LevelIndex?.Value == ancestorIndex);
            if (ancestor == null)
                break;
            if (tracker.GetActiveCurrent(abstractId, ancestorIndex) == null)
                tracker.SeedUnstartedAncestor(numPr.NumId, abstractId, ancestorIndex,
                    ancestor.StartNumberingValue?.Val?.Value ?? 0,
                    restartNever: nr.GetRestartTrigger(numPr.NumId, ancestorIndex) == -1,
                    effectiveStart: nr.ResolveLevel(numPr.NumId, ancestorIndex)?.startAt);
        }

        // Hidden numeric paragraphs advance; an invalid/future reference suppresses
        // the label without visiting a descendant counter.
        if (string.IsNullOrEmpty(lvlText) || references.Any(reference =>
            reference.Value[1] == '0' || reference.Value[1] - '1' > numPr.Ilvl))
            return null;

        // If bullet, Word stores the glyph directly in lvlText (and font is in lvl.rPr).
        if (fmt == NumberFormatValues.Bullet)
        {
            var font = GetBulletFont(lvl);
            if (!string.IsNullOrEmpty(font) && lvlText.Length > 0)
                return $"""<span style="font-family:{font}">{lvlText}</span>""";
            return lvlText;
        }

        var substitutions = new Dictionary<string, string>(StringComparer.Ordinal);

        for (int i = 0; i < 9; i++)
        {
            var placeholder = "%" + (i + 1).ToString(CultureInfo.InvariantCulture);
            if (!references.Any(reference => reference.Value == placeholder))
                continue;

            var cur = i < numPr.Ilvl
                ? tracker.GetActiveCurrent(abstractId, i)
                : tracker.GetCurrent(numPr.NumId, abstractId, i);
            if (cur == null)
            {
                // Word displays the abstract start value before the ancestor emits.
                // It also seeds the shared counter, so its first later item advances.
                var abstractLevel = abs.Elements<Level>()
                    .FirstOrDefault(level => level.LevelIndex?.Value == i);
                cur = tracker.SeedUnstartedAncestor(numPr.NumId, abstractId, i,
                    abstractLevel?.StartNumberingValue?.Val?.Value ?? 0,
                    restartTriggers[i] == -1, effectiveStarts[i]);
            }

            var refLvl = nr.ResolveLevel(numPr.NumId, i)?.lvl;
            var refFmt = refLvl?.NumberingFormat?.Val?.Value ?? NumberFormatValues.Decimal;
            var legal = lvl.GetFirstChild<IsLegalNumberingStyle>();
            if (legal != null && (legal.Val == null || legal.Val.Value) &&
                refFmt != NumberFormatValues.None && refFmt != NumberFormatValues.Bullet &&
                refFmt != NumberFormatValues.DecimalZero)
                refFmt = NumberFormatValues.Decimal;



            // Custom "0001" padding (from mc:AlternateContent) if present
            int? customWidth = legal?.Val?.Value != false && legal != null ? null
                : refLvl != null ? TryGetCustomDecimalWidth(refLvl) : null;

            string formatted = customWidth != null
                ? cur.Value.ToString(CultureInfo.InvariantCulture).PadLeft(customWidth.Value, '0')
                : FormatListNumber(cur.Value, refFmt, culture);
            if ((refFmt == NumberFormatValues.CardinalText || refFmt == NumberFormatValues.OrdinalText) &&
                formatted.Length > 0)
                formatted = char.ToUpper(formatted[0], culture ?? CultureInfo.CurrentCulture) + formatted.Substring(1);

            substitutions[placeholder] = formatted;

        }

        var referenceIndex = 0;
        return System.Text.RegularExpressions.Regex.Replace(lvlText, "%[1-9]", match =>
            referenceIndex++ < 9 && substitutions.TryGetValue(match.Value, out var value) ? value : match.Value);
    }

    private static string FormatListNumber(int n, NumberFormatValues fmt, CultureInfo? culture)
    {
        // Word lists repeat a single Latin letter (AA, BB). Field formatting
        // uses positional letters (AA, AB) and continues through FormatNumber.
        if (n > 0 && (fmt == NumberFormatValues.UpperLetter || fmt == NumberFormatValues.LowerLetter))
        {
            var first = fmt == NumberFormatValues.UpperLetter ? 'A' : 'a';
            return new string((char)(first + (n - 1) % 26), (n - 1) / 26 % 30 + 1);
        }
        return FormatNumber(n, fmt, culture);
    }

    public static string FormatNumber(int n, NumberFormatValues fmt, CultureInfo? culture = null)
    {
        if (fmt == NumberFormatValues.None)
            return "";
        culture ??= CultureInfo.CurrentCulture;
        if (fmt == NumberFormatValues.Decimal)
            return n.ToString(CultureInfo.InvariantCulture);
        if (fmt == NumberFormatValues.DecimalZero)
            return n < 10 ? "0" + n.ToString(CultureInfo.InvariantCulture) : n.ToString(CultureInfo.InvariantCulture);
        if (fmt == NumberFormatValues.DecimalEnclosedCircle)
            return n switch
            {
                0 => "\u24ea",
                >= 1 and <= 20 => ((char)('\u2460' + n - 1)).ToString(),
                >= 21 and <= 35 => ((char)('\u3251' + n - 21)).ToString(),
                >= 36 and <= 50 => ((char)('\u32b1' + n - 36)).ToString(),
                _ => n.ToString(CultureInfo.InvariantCulture)
            };
        if (fmt == NumberFormatValues.NumberInDash)
            return "- " + n.ToString(CultureInfo.InvariantCulture) + " -";
        if (fmt == NumberFormatValues.DecimalFullWidth || fmt == NumberFormatValues.DecimalFullWidth2)
            return string.Concat(n.ToString(CultureInfo.InvariantCulture)
                .Select(digit => digit is >= '0' and <= '9' ? (char)('０' + digit - '0') : digit));
        if (fmt == NumberFormatValues.IdeographDigital)
            return string.Concat(n.ToString(CultureInfo.InvariantCulture)
                .Select(digit => digit is >= '0' and <= '9' ? "〇一二三四五六七八九"[digit - '0'] : digit));
        if (fmt == NumberFormatValues.Hex)
            return n.ToString("X", CultureInfo.InvariantCulture);
        if (fmt == NumberFormatValues.Chicago && n > 0)
        {
            const string symbols = "*†‡§";
            return FormatRepeatedSymbol(n, symbols);
        }
        if (fmt == NumberFormatValues.UpperRoman)
            return ToRoman(n).ToUpperInvariant();
        if (fmt == NumberFormatValues.LowerRoman)
            return ToRoman(n).ToLowerInvariant();
        if (fmt == NumberFormatValues.UpperLetter)
            return ToAlpha(n).ToUpperInvariant();
        if (fmt == NumberFormatValues.LowerLetter)
            return ToAlpha(n).ToLowerInvariant();
        if (fmt == NumberFormatValues.RussianLower)
            return ToRepeatedAlphabet(n, "абвгдежзиклмнопрстуфхцчшщыэюя");
        if (fmt == NumberFormatValues.RussianUpper)
            return ToRepeatedAlphabet(n, "АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЫЭЮЯ");
        if (fmt == NumberFormatValues.CardinalText)
            return ResolveNumberProvider(culture).ToCardinal(n);
        if (fmt == NumberFormatValues.OrdinalText)
            return ResolveNumberProvider(culture).ToOrdinalWords(n);
        if (fmt == NumberFormatValues.Ordinal)
            return n.ToString(CultureInfo.InvariantCulture) + OrdinalSuffix(n);

        return n.ToString(CultureInfo.InvariantCulture);
    }

    private static DxpINumberToWordsProvider ResolveNumberProvider(CultureInfo culture)
    {
        return DxpNumberToWordsRegistry.Default.Resolve(culture);
    }

    private static string OrdinalSuffix(int n)
    {
        int abs = Math.Abs(n);
        int mod100 = abs % 100;
        if (mod100 >= 11 && mod100 <= 13)
            return "th";
        return (abs % 10) switch {
            1 => "st",
            2 => "nd",
            3 => "rd",
            _ => "th"
        };
    }

    private static string ToAlpha(int n)
    {
        // 1 -> A, 26 -> Z, 27 -> AA ...
        if (n <= 0)
            return n.ToString(CultureInfo.InvariantCulture);
        var s = "";
        while (n > 0)
        {
            n--; // 1-based
            s = (char)('A' + n % 26) + s;
            n /= 26;
        }
        return s;
    }

    private static string ToRepeatedAlphabet(int number, string alphabet)
    {
        if (number <= 0)
            return number.ToString(CultureInfo.InvariantCulture);
        return FormatRepeatedSymbol(number, alphabet);
    }

    private static string FormatRepeatedSymbol(int number, string symbols)
    {
        int repetitions = (number - 1) / symbols.Length + 1;
        return repetitions > 256
            ? number.ToString(CultureInfo.InvariantCulture)
            : new string(symbols[(number - 1) % symbols.Length], repetitions);
    }

    private static string ToRoman(int number)
    {
        if (number <= 0)
            return number.ToString(CultureInfo.InvariantCulture);
        var map = new (int v, string s)[] {
            (1000,"M"),(900,"CM"),(500,"D"),(400,"CD"),
            (100,"C"),(90,"XC"),(50,"L"),(40,"XL"),
            (10,"X"),(9,"IX"),(5,"V"),(4,"IV"),(1,"I")
        };
        var result = "";
        foreach (var (v, s) in map)
            while (number >= v)
            { result += s; number -= v; }
        return result;
    }

    private static string? GetBulletFont(Level lvl)
    {
        var fonts = lvl.NumberingSymbolRunProperties?.RunFonts;
        if (fonts == null)
            return null;

        return fonts.Ascii?.Value
            ?? fonts.HighAnsi?.Value;
    }
}

public class DxpLists
{
    private sealed class OpenXmlElementReferenceComparer : IEqualityComparer<OpenXmlElement>
    {
        public bool Equals(OpenXmlElement? x, OpenXmlElement? y) => ReferenceEquals(x, y);
        public int GetHashCode(OpenXmlElement obj) => RuntimeHelpers.GetHashCode(obj);
    }

    private DxpNumberingResolver? _num;
    private readonly Dictionary<OpenXmlElement, DxpListTracker> _listsByStory =
        new(new OpenXmlElementReferenceComparer());
    private readonly DxpListTracker _marginLists = new();

    public DxpLists()
    {
    }

    public static string FormatNumber(int n, NumberFormatValues fmt, CultureInfo? culture = null)
    {
        return DxpListMarkerFormatter.FormatNumber(n, fmt, culture);
    }

    internal void Init(WordprocessingDocument doc)
    {
        _num = new DxpNumberingResolver(doc);
        _listsByStory.Clear();
        _marginLists.ClearAll();
    }

    internal DxpMarker MaterializeMarker(Paragraph p, DxpIStyleResolver s)
    {
        string? marker = null;
        int? numId = null;
        int? iLvl = null;
        var separator = " ";

        if (_num != null)
        {
            var numPr = s.ResolveEffectiveNumPr(p);
            if (numPr != null)
            {
                var culture = ResolveListCulture(p, s);
                var ancestors = p.Ancestors().ToArray();
                var lists = _marginLists;
                if (!ancestors.Any(element => element is Header or Footer))
                {
                    var story = ancestors.LastOrDefault() ?? p;
                    if (!_listsByStory.TryGetValue(story, out lists))
                        _listsByStory[story] = lists = new DxpListTracker();
                }
                marker = DxpListMarkerFormatter.TryBuildMarkerText(numPr, _num, lists, culture);
                numId = numPr.NumId;
                iLvl = numPr.Ilvl;
                if (_num.ResolveLevel(numPr.NumId, numPr.Ilvl)?.lvl.LevelSuffix?.Val?.Value == LevelSuffixValues.Nothing)
                    separator = string.Empty;
            }
        }
        return new DxpMarker(marker, numId, iLvl) { TextSeparator = separator };
    }

    public DxpStyleEffectiveIndentTwips GetIndentation(Paragraph p, DxpIStyleResolver s)
    {
        DxpStyleEffectiveIndentTwips indent = s.GetIndentation(p, _num);
        return indent;
    }

    private static CultureInfo ResolveListCulture(Paragraph p, DxpIStyleResolver s)
    {
        if (s is DxpStyleResolver resolver)
        {
            var lang = resolver.ResolveParagraphLanguage(p) ?? resolver.GetDefaultLanguage();
            if (!string.IsNullOrWhiteSpace(lang))
            {
                try
                {
                    return new CultureInfo(lang);
                }
                catch (CultureNotFoundException)
                {
                }
            }
        }
        return CultureInfo.CurrentCulture;
    }
}
