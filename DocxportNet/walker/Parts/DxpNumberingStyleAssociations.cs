using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Walker;

internal static class DxpNumberingStyleAssociations
{
    internal static Styles? Normalize(Styles? source, Numbering? numbering)
    {
        if (source == null || numbering == null) return source;
        // Word imports style/list associations in definition order. Keep the
        // resulting cancellations private to numbering; layout uses the source.
        var styles = (Styles)source.CloneNode(true);
        var ordered = styles.Elements<Style>().Where(s => s.Type?.Value == StyleValues.Paragraph).ToArray();
        var byId = ordered.Where(s => s.StyleId?.Value != null)
            .GroupBy(s => s.StyleId!.Value!, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.Last(), StringComparer.Ordinal);
        var instances = numbering.Elements<NumberingInstance>().Where(n => n.NumberID?.Value != null && n.AbstractNumId?.Val?.Value != null)
            .GroupBy(n => n.NumberID!.Value).ToDictionary(g => g.Key, g => g.Last().AbstractNumId!.Val!.Value);
        var abstracts = numbering.Elements<AbstractNum>().Where(a => a.AbstractNumberId?.Value != null)
            .GroupBy(a => a.AbstractNumberId!.Value).ToDictionary(g => g.Key, g => g.Last());
        var bindings = new Dictionary<(int Abstract, int Level), string>();
        foreach (var pair in abstracts)
            foreach (var level in pair.Value.Elements<Level>())
                if (level.LevelIndex?.Value is int index && level.ParagraphStyleIdInLevel?.Val?.Value is string id &&
                    byId.TryGetValue(id, out var target))
                {
                    var associatedId = Chain(target).Select(style =>
                        style.StyleParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value)
                        .FirstOrDefault(value => value != null);
                    if (associatedId is > 0 && instances.TryGetValue(associatedId.Value, out var owner) && owner == pair.Key)
                        bindings[(pair.Key, index)] = id;
                }

        IEnumerable<Style> Chain(Style? style)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            while (style?.StyleId?.Value is string id && seen.Add(id))
            {
                yield return style;
                var parent = style.BasedOn?.Val?.Value;
                style = parent != null && byId.TryGetValue(parent, out var next) ? next : null;
            }
        }

        var changed = false;
        foreach (var style in ordered)
        {
            int? numId = null, index = null;
            foreach (var ancestor in Chain(style))
            {
                var properties = ancestor.StyleParagraphProperties?.NumberingProperties;
                numId ??= properties?.NumberingId?.Val?.Value;
                index ??= properties?.NumberingLevelReference?.Val?.Value;
            }
            if (numId is not > 0 || !instances.TryGetValue(numId.Value, out var abstractId) ||
                !abstracts.TryGetValue(abstractId, out var definition) || !definition.Elements<Level>().Any()) continue;
            var key = (abstractId, index ?? 0);
            var id = style.StyleId?.Value;
            if (id == null) continue;
            if (!bindings.TryGetValue(key, out var bound))
            {
                bindings[key] = id;
                continue;
            }
            if (bound == id) continue;
            var own = style.StyleParagraphProperties?.NumberingProperties;
            if (own == null || !own.ChildElements.Any()) continue;
            // An ancestor losing its explicit instance to a derived style loses
            // numPr. Other competing direct properties become an explicit stop.
            if (own.NumberingId != null && byId.TryGetValue(bound, out var owner) && Chain(owner).Any(s => ReferenceEquals(s, style)))
                own.Remove();
            else
                style.StyleParagraphProperties!.NumberingProperties = new NumberingProperties(new NumberingId { Val = 0 });
            changed = true;
        }
        return changed ? styles : source;
    }
}
