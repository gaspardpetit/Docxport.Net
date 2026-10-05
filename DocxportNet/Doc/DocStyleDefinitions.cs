using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>A named DOC style and the properties currently decoded from its UPX.</summary>
public sealed record DocStyleDefinition(int Index, string Name, int Type,
    int? BasedOnIndex, int? NextIndex, DocCharacterFormatting CharacterFormatting,
    DocParagraphFormatting? ParagraphFormatting = null,
    DocParagraphFormatting? TableFormatting = null,
    int? InvariantStyleId = null,
    int? LinkedStyleIndex = null,
    DocCharacterFormatting? DirectCharacterFormatting = null,
    DocParagraphFormatting? DirectParagraphFormatting = null,
    IReadOnlyDictionary<ushort, DocParagraphFormatting>? ConditionalParagraphFormatting = null,
    IReadOnlyDictionary<ushort, DocCharacterFormatting>? ConditionalCharacterFormatting = null,
    IReadOnlyDictionary<ushort, DocCellShading>? ConditionalTableShading = null,
    IReadOnlyDictionary<ushort, DocConditionalTableBorders>? ConditionalTableBorders = null,
    IReadOnlyDictionary<ushort, byte>? ConditionalTableVerticalAlignment = null,
    IReadOnlyDictionary<ushort, bool>? ConditionalTableNoWrap = null)
{
    public string StyleId => $"DocStyle{Index}";
}

internal static class DocStyleDefinitionsReader
{
    public static IReadOnlyList<DocStyleDefinition> Read(DocTextIndex index)
    {
        var result = new List<DocStyleDefinition>();
        foreach (var entry in index.Styles)
        {
            var definition = entry.Children.FirstOrDefault(x => x.Kind == "STD");
            if (definition == null) continue;
            DocStyleDefinitionNavigator.Expand(index.Structure, definition);
            var fields = Descendants(definition).First(x => x.Kind == "StdfBase");
            var name = definition.Children.First(x => x.Kind == "Xstz").Attributes["text"];
            var upx = Descendants(definition).FirstOrDefault(x => x.Kind == "UpxChpx");
            var formatting = upx?.StreamName is string stream &&
                upx.Offset is long offset && upx.Length is long length && length > 0
                ? DocCharacterFormattingReader.ParseSprms(index.Structure.ReadRange(stream,
                    offset, checked((int)length)))
                : DocCharacterFormatting.Empty;
            if (formatting.AsciiFontIndex != null || formatting.EastAsiaFontIndex != null ||
                formatting.HighAnsiFontIndex != null ||
                formatting.ComplexScriptFontIndex != null)
                formatting = DocCharacterFormattingReader.ResolveFonts(formatting, index.Fonts);
            var paragraphUpx = Descendants(definition).FirstOrDefault(x => x.Kind == "UpxPapx");
            var tableUpx = Descendants(definition).FirstOrDefault(x => x.Kind == "UpxTapx");
            var tableFormatting = tableUpx?.StreamName is string tableStream &&
                tableUpx.Offset is long tableOffset && tableUpx.Length is long tableLength &&
                tableLength > 0
                ? DocParagraphFormatting.Parse(index.Structure.ReadRange(tableStream,
                    tableOffset, checked((int)tableLength)))
                : DocParagraphFormatting.Empty;
            var paragraphFormatting = DocParagraphFormatting.Empty;
            if (paragraphUpx?.StreamName is string paragraphStream &&
                paragraphUpx.Offset is long paragraphOffset &&
                paragraphUpx.Length is long paragraphLength && paragraphLength > 2)
            {
                var bytes = index.Structure.ReadRange(paragraphStream, paragraphOffset,
                    checked((int)paragraphLength));
                paragraphFormatting = DocParagraphFormatting.Parse(bytes.AsSpan(2));
            }
            Dictionary<ushort, DocParagraphFormatting>? conditionalParagraphs = null;
            Dictionary<ushort, DocCharacterFormatting>? conditionalCharacters = null;
            Dictionary<ushort, DocCellShading>? conditionalShadings = null;
            Dictionary<ushort, DocConditionalTableBorders>? conditionalBorders = null;
            Dictionary<ushort, byte>? conditionalVerticalAlignments = null;
            Dictionary<ushort, bool>? conditionalNoWraps = null;
            if (tableUpx != null)
            {
                DocPropertyNavigator.Expand(index.Structure, tableUpx);
                foreach (var prl in Descendants(tableUpx).Where(x =>
                    x.Kind == "Prl" && x.Children.Any(child =>
                        child.Kind == "Sprm" && child.Attributes.TryGetValue("code",
                            out var code) && code == "0xD66A")))
                {
                    if (prl.StreamName == null || prl.Offset == null ||
                        prl.Length is not long prlLength || prlLength < 5)
                        throw new InvalidDataException("A conditional table rule is truncated.");
                    var operand = index.Structure.ReadRange(prl.StreamName,
                        prl.Offset.Value + 2, checked((int)prlLength - 2));
                    if (operand.Length < 3 || operand[0] != operand.Length - 1)
                        throw new InvalidDataException("A conditional table rule has an invalid length.");
                    var condition = BinaryPrimitives.ReadUInt16LittleEndian(
                        operand.AsSpan(1));
                    if (!DocTableStyleCondition.IsValid(condition)) continue;
                    var conditionalTableFormatting = DocParagraphFormatting.Parse(
                        operand.AsSpan(3));
                    if (conditionalTableFormatting.TableStyleShading is { } shade)
                    {
                        conditionalShadings ??= new();
                        conditionalShadings[condition] = shade;
                    }
                    if (DocConditionalTableBorders.Parse(operand.AsSpan(3))
                        is { } borders)
                    {
                        conditionalBorders ??= new();
                        conditionalBorders[condition] = borders;
                    }
                    if (conditionalTableFormatting.TableStyleVerticalAlignment
                        is byte alignment)
                    {
                        conditionalVerticalAlignments ??= new();
                        conditionalVerticalAlignments[condition] = alignment;
                    }
                    if (conditionalTableFormatting.TableStyleNoWrap is bool noWrap)
                    {
                        conditionalNoWraps ??= new();
                        conditionalNoWraps[condition] = noWrap;
                    }
                }
            }
            if (tableUpx != null && paragraphUpx != null)
            {
                DocPropertyNavigator.Expand(index.Structure, paragraphUpx);
                foreach (var prl in Descendants(paragraphUpx).Where(x =>
                    x.Kind == "Prl" && x.Children.Any(child =>
                        child.Kind == "Sprm" && child.Attributes.TryGetValue("code",
                            out var code) && code == "0xC666")))
                {
                    if (prl.StreamName == null || prl.Offset == null ||
                        prl.Length is not long prlLength || prlLength < 5)
                        throw new InvalidDataException("A conditional paragraph rule is truncated.");
                    var operand = index.Structure.ReadRange(prl.StreamName,
                        prl.Offset.Value + 2, checked((int)prlLength - 2));
                    if (operand[0] != operand.Length - 1 || operand.Length < 3)
                        throw new InvalidDataException("A conditional paragraph rule has an invalid length.");
                    var condition = BinaryPrimitives.ReadUInt16LittleEndian(
                        operand.AsSpan(1));
                    if (!DocTableStyleCondition.IsValid(condition)) continue;
                    conditionalParagraphs ??= new();
                    conditionalParagraphs[condition] =
                        DocParagraphFormatting.Parse(operand.AsSpan(3));
                }
            }
            if (tableUpx != null && upx != null)
            {
                DocPropertyNavigator.Expand(index.Structure, upx);
                foreach (var prl in Descendants(upx).Where(x =>
                    x.Kind == "Prl" && x.Children.Any(child =>
                        child.Kind == "Sprm" && child.Attributes.TryGetValue("code",
                            out var code) && code == "0xCA85")))
                {
                    if (prl.StreamName == null || prl.Offset == null ||
                        prl.Length is not long prlLength || prlLength < 5)
                        throw new InvalidDataException("A conditional run rule is truncated.");
                    var operand = index.Structure.ReadRange(prl.StreamName,
                        prl.Offset.Value + 2, checked((int)prlLength - 2));
                    if (operand.Length < 3 || operand[0] != operand.Length - 1)
                        throw new InvalidDataException("A conditional run rule has an invalid length.");
                    var condition = BinaryPrimitives.ReadUInt16LittleEndian(
                        operand.AsSpan(1));
                    if (!DocTableStyleCondition.IsValid(condition)) continue;
                    var character = DocCharacterFormattingReader.ParseSprms(
                        operand.AsSpan(3));
                    conditionalCharacters ??= new();
                    conditionalCharacters[condition] =
                        DocCharacterFormattingReader.ResolveFonts(character, index.Fonts);
                }
            }
            var basedOn = int.Parse(fields.Attributes["basedOnStyleIndex"],
                CultureInfo.InvariantCulture);
            var next = int.Parse(fields.Attributes["nextStyleIndex"],
                CultureInfo.InvariantCulture);
            var post = Descendants(definition).FirstOrDefault(x => x.Kind == "StdfPost2000");
            var linked = post != null && post.Attributes.TryGetValue("linkedStyleIndex", out var value)
                ? int.Parse(value, CultureInfo.InvariantCulture) : 0;
            result.Add(new DocStyleDefinition(
                int.Parse(entry.Attributes["styleIndex"], CultureInfo.InvariantCulture),
                name,
                int.Parse(fields.Attributes["styleType"], CultureInfo.InvariantCulture),
                basedOn == 0x0FFF ? null : basedOn,
                next == 0x0FFF ? null : next,
                formatting, paragraphFormatting, tableFormatting,
                int.Parse(fields.Attributes["invariantStyleId"], CultureInfo.InvariantCulture),
                linked == 0 ? null : linked,
                ConditionalParagraphFormatting: conditionalParagraphs,
                ConditionalCharacterFormatting: conditionalCharacters,
                ConditionalTableShading: conditionalShadings,
                ConditionalTableBorders: conditionalBorders,
                ConditionalTableVerticalAlignment: conditionalVerticalAlignments,
                ConditionalTableNoWrap: conditionalNoWraps));
        }
        var source = result.ToDictionary(x => x.Index);
        var resolved = new Dictionary<int, DocStyleDefinition>();
        var visiting = new HashSet<int>();
        DocStyleDefinition? Resolve(int index)
        {
            if (resolved.TryGetValue(index, out var ready)) return ready;
            if (!source.TryGetValue(index, out var style)) return null;
            if (!visiting.Add(index))
                throw new InvalidDataException("The DOC style inheritance graph contains a cycle.");
            var parentStyle = style.BasedOnIndex is int parentIndex
                ? Resolve(parentIndex) : null;
            var parent = parentStyle?.CharacterFormatting ?? DocCharacterFormatting.Empty;
            var own = style.CharacterFormatting.ResolveStyleToggles(parent);
            var effective = own with
            {
                Bold = own.Bold ?? parent.Bold,
                Italic = own.Italic ?? parent.Italic,
                Strike = own.Strike ?? parent.Strike,
                DoubleStrike = own.DoubleStrike ?? parent.DoubleStrike,
                Shadow = own.Shadow ?? parent.Shadow,
                Outline = own.Outline ?? parent.Outline,
                Emboss = own.Emboss ?? parent.Emboss,
                Imprint = own.Imprint ?? parent.Imprint,
                Caps = own.Caps ?? parent.Caps,
                SmallCaps = own.SmallCaps ?? parent.SmallCaps,
                Hidden = own.Hidden ?? parent.Hidden,
                SizeHalfPoints = own.SizeHalfPoints ?? parent.SizeHalfPoints,
                UnderlineCode = own.UnderlineCode ?? parent.UnderlineCode,
                UnderlineColorRef = own.UnderlineColorRef ?? parent.UnderlineColorRef,
                ColorRef = own.ColorRef ?? parent.ColorRef,
                HighlightCode = own.HighlightCode ?? parent.HighlightCode,
                ScriptCode = own.ScriptCode ?? parent.ScriptCode,
                Shading = own.Shading ?? parent.Shading,
                Border = own.Border ?? parent.Border,
                AsciiFontIndex = own.AsciiFontIndex ?? parent.AsciiFontIndex,
                EastAsiaFontIndex = own.EastAsiaFontIndex ?? parent.EastAsiaFontIndex,
                HighAnsiFontIndex = own.HighAnsiFontIndex ?? parent.HighAnsiFontIndex,
                ComplexScriptFontIndex = own.ComplexScriptFontIndex ??
                    parent.ComplexScriptFontIndex,
                AsciiFontName = own.AsciiFontName ?? parent.AsciiFontName,
                EastAsiaFontName = own.EastAsiaFontName ?? parent.EastAsiaFontName,
                HighAnsiFontName = own.HighAnsiFontName ?? parent.HighAnsiFontName,
                ComplexScriptFontName = own.ComplexScriptFontName ??
                    parent.ComplexScriptFontName,
                CharacterSpacingTwips = own.CharacterSpacingTwips ??
                    parent.CharacterSpacingTwips,
                KerningThresholdHalfPoints = own.KerningThresholdHalfPoints ??
                    parent.KerningThresholdHalfPoints,
                CharacterScalePercent = own.CharacterScalePercent ??
                    parent.CharacterScalePercent,
                SnapToGrid = own.SnapToGrid ?? parent.SnapToGrid,
                EmphasisMarkCode = own.EmphasisMarkCode ?? parent.EmphasisMarkCode,
                FitText = own.FitText ?? parent.FitText,
                BaselineOffsetHalfPoints = own.BaselineOffsetHalfPoints ??
                    parent.BaselineOffsetHalfPoints,
                LanguageId = own.LanguageId ?? parent.LanguageId,
                EastAsiaLanguageId = own.EastAsiaLanguageId ?? parent.EastAsiaLanguageId,
                ComplexScriptLanguageId = own.ComplexScriptLanguageId ??
                    parent.ComplexScriptLanguageId,
                ComplexScriptBold = own.ComplexScriptBold ?? parent.ComplexScriptBold,
                ComplexScriptItalic = own.ComplexScriptItalic ?? parent.ComplexScriptItalic,
                ComplexScriptSizeHalfPoints = own.ComplexScriptSizeHalfPoints ??
                    parent.ComplexScriptSizeHalfPoints,
                RightToLeftText = own.RightToLeftText ?? parent.RightToLeftText,
                ForceComplexScript = own.ForceComplexScript ?? parent.ForceComplexScript
            };
            var paragraph = style.ParagraphFormatting ?? DocParagraphFormatting.Empty;
            var parentParagraph = parentStyle?.ParagraphFormatting ??
                DocParagraphFormatting.Empty;
            var (clearedTabs, tabStops) = ComposeTabs(parentParagraph, paragraph);
            var effectiveParagraph = paragraph with
            {
                Justification = paragraph.Justification ?? parentParagraph.Justification,
                BeforeTwips = paragraph.BeforeTwips ?? parentParagraph.BeforeTwips,
                AfterTwips = paragraph.AfterTwips ?? parentParagraph.AfterTwips,
                LeftTwips = paragraph.LeftTwips ?? parentParagraph.LeftTwips,
                RightTwips = paragraph.RightTwips ?? parentParagraph.RightTwips,
                FirstLineTwips = paragraph.FirstLineTwips ?? parentParagraph.FirstLineTwips,
                KeepLines = paragraph.KeepLines ?? parentParagraph.KeepLines,
                KeepWithNext = paragraph.KeepWithNext ?? parentParagraph.KeepWithNext,
                PageBreakBefore = paragraph.PageBreakBefore ??
                    parentParagraph.PageBreakBefore,
                WidowControl = paragraph.WidowControl ?? parentParagraph.WidowControl,
                ParagraphRightToLeft = paragraph.ParagraphRightToLeft ??
                    parentParagraph.ParagraphRightToLeft,
                OutlineLevel = paragraph.OutlineLevel ?? parentParagraph.OutlineLevel,
                Kinsoku = paragraph.Kinsoku ?? parentParagraph.Kinsoku,
                WordWrap = paragraph.WordWrap ?? parentParagraph.WordWrap,
                SnapToGrid = paragraph.SnapToGrid ?? parentParagraph.SnapToGrid,
                AutoSpaceDE = paragraph.AutoSpaceDE ?? parentParagraph.AutoSpaceDE,
                AutoSpaceDN = paragraph.AutoSpaceDN ?? parentParagraph.AutoSpaceDN,
                ContextualSpacing = paragraph.ContextualSpacing ??
                    parentParagraph.ContextualSpacing,
                MirrorIndents = paragraph.MirrorIndents ?? parentParagraph.MirrorIndents,
                SuppressAutoHyphens = paragraph.SuppressAutoHyphens ??
                    parentParagraph.SuppressAutoHyphens,
                TextAlignmentCode = paragraph.TextAlignmentCode ??
                    parentParagraph.TextAlignmentCode,
                SuppressLineNumbers = paragraph.SuppressLineNumbers ??
                    parentParagraph.SuppressLineNumbers,
                AdjustRightIndent = paragraph.AdjustRightIndent ??
                    parentParagraph.AdjustRightIndent,
                BeforeAutoSpacing = paragraph.BeforeAutoSpacing ??
                    parentParagraph.BeforeAutoSpacing,
                AfterAutoSpacing = paragraph.AfterAutoSpacing ??
                    parentParagraph.AfterAutoSpacing,
                BeforeLines = paragraph.BeforeLines ?? parentParagraph.BeforeLines,
                AfterLines = paragraph.AfterLines ?? parentParagraph.AfterLines,
                LeftChars = paragraph.LeftChars ?? parentParagraph.LeftChars,
                RightChars = paragraph.RightChars ?? parentParagraph.RightChars,
                FirstLineChars = paragraph.FirstLineChars ??
                    parentParagraph.FirstLineChars,
                LineValue = paragraph.LineValue ?? parentParagraph.LineValue,
                LineIsMultiple = paragraph.LineValue != null
                    ? paragraph.LineIsMultiple : parentParagraph.LineIsMultiple,
                ClearedTabPositions = clearedTabs,
                TabStops = tabStops,
                FillRgb = paragraph.FillRgb ?? parentParagraph.FillRgb,
                ShadingForegroundRgb = paragraph.ShadingForegroundRgb ??
                    parentParagraph.ShadingForegroundRgb,
                ShadingPattern = paragraph.ShadingPattern ?? parentParagraph.ShadingPattern,
                TopBorder = paragraph.TopBorder ?? parentParagraph.TopBorder,
                LeftBorder = paragraph.LeftBorder ?? parentParagraph.LeftBorder,
                BottomBorder = paragraph.BottomBorder ?? parentParagraph.BottomBorder,
                RightBorder = paragraph.RightBorder ?? parentParagraph.RightBorder,
                BetweenBorder = paragraph.BetweenBorder ?? parentParagraph.BetweenBorder,
                ListOverrideIndex = paragraph.ListOverrideIndex ??
                    parentParagraph.ListOverrideIndex,
                ListLevel = paragraph.ListLevel ?? parentParagraph.ListLevel
            };
            var value = style with { CharacterFormatting = effective,
                ParagraphFormatting = effectiveParagraph,
                DirectCharacterFormatting = own,
                DirectParagraphFormatting = paragraph };
            visiting.Remove(index);
            resolved.Add(index, value);
            return value;
        }
        return result.Select(x => Resolve(x.Index)!).ToArray();
    }

    internal static (short[] Cleared, DocTabStop[] Added) ComposeTabs(
        DocParagraphFormatting parent, DocParagraphFormatting child)
    {
        var childClears = new HashSet<short>(child.ClearedTabPositions ?? []);
        var ranges = child.TabClearRanges ?? [];
        var added = new Dictionary<short, DocTabStop>();
        foreach (var stop in parent.TabStops ?? [])
            if (!childClears.Contains(stop.PositionTwips) &&
                !ranges.Any(x => Math.Abs(stop.PositionTwips - x.PositionTwips)
                    <= x.ToleranceTwips))
                added[stop.PositionTwips] = stop;
        var cleared = (parent.ClearedTabPositions ?? [])
            .Concat(child.ClearedTabPositions ?? [])
            .Concat((parent.TabStops ?? [])
                .Where(x => !added.ContainsKey(x.PositionTwips))
                .Select(x => x.PositionTwips))
            .Distinct().OrderBy(x => x).ToArray();
        foreach (var stop in child.TabStops ?? [])
            added[stop.PositionTwips] = stop;
        return (cleared, added.Values.OrderBy(x => x.PositionTwips).ToArray());
    }

    private static IEnumerable<DocStructureNode> Descendants(DocStructureNode node)
    {
        foreach (var child in node.Children)
        {
            yield return child;
            foreach (var descendant in Descendants(child)) yield return descendant;
        }
    }
}
