using System.Buffers.Binary;
using System.Text;

namespace DocxportNet.Doc;

/// <summary>Minimal default structures for an unformatted main-story document.</summary>
internal static class DocDefaultStructures
{
    internal const string PreservedRunDefaultsStyleName = "DocxportNet Preserved Run Defaults";

    public static byte[] CreateStyleSheet()
        => CreateStyleSheet(Array.Empty<DocStyleDefinition>());

    public static byte[] CreateStyleSheet(IReadOnlyList<DocStyleDefinition> styles)
        => CreateStyleSheet(styles, new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase));

    public static byte[] CreateStyleSheet(IReadOnlyList<DocStyleDefinition> styles,
        IReadOnlyDictionary<string, int> fontIndexes,
        DocCharacterFormatting? defaults = null,
        DocParagraphFormatting? paragraphDefaults = null)
    {
        var count = Math.Max(15, styles.Count == 0 ? 15 : styles.Max(x => x.Index) + 1);
        if (count >= 0x0FFE || styles.Any(x =>
            (x.Index < 15 && (x.Index != 0 || x.Type != 1)) ||
            x.Type is not (1 or 2 or 3)))
            throw new InvalidDataException("The DOC stylesheet contains an invalid style index or type.");
        var byIndex = styles.ToDictionary(x => x.Index);
        if (defaults is { } characterDefaults &&
            !(characterDefaults with
            {
                AsciiFontName = null,
                EastAsiaFontName = null,
                HighAnsiFontName = null,
                ComplexScriptFontName = null
            }).IsEmpty ||
            paragraphDefaults is { IsEmpty: false })
            if (!byIndex.ContainsKey(0))
                byIndex.Add(0, new DocStyleDefinition(0, "Normal", 1, null, 0,
                    defaults ?? DocCharacterFormatting.Empty, paragraphDefaults,
                    InvariantStyleId: 0));
        var baseSize = byIndex.Values.Any(x => x.LinkedStyleIndex is int linked &&
            linked != 0 && byIndex.ContainsKey(linked)) ? (ushort)18 : (ushort)10;
        using var output = new MemoryStream();
        short FontIndex(string? name) => name != null && fontIndexes.TryGetValue(name, out var index)
            ? checked((short)index) : (short)0;
        var header = new DocStyleSheetHeader(checked((ushort)count), baseSize, 1, 0, 15,
            FontIndex(defaults?.AsciiFontName), FontIndex(defaults?.EastAsiaFontName),
            FontIndex(defaults?.HighAnsiFontName),
            FontIndex(defaults?.ComplexScriptFontName)).Write();
        output.Write(header, 0, header.Length);
        for (var index = 0; index < count; index++)
        {
            if (!byIndex.TryGetValue(index, out var style))
            {
                output.WriteByte(0);
                output.WriteByte(0);
                continue;
            }
            var body = CreateStyleDefinition(style, byIndex, fontIndexes, baseSize);
            WriteU16(output, checked((ushort)body.Length));
            output.Write(body, 0, body.Length);
            if ((body.Length & 1) != 0) output.WriteByte(0);
        }
        return output.ToArray();
    }

    private static byte[] CreateStyleDefinition(DocStyleDefinition style,
        IReadOnlyDictionary<int, DocStyleDefinition> styles,
        IReadOnlyDictionary<string, int> fontIndexes, ushort baseSize)
    {
        var baseFields = new byte[baseSize];
        BinaryPrimitives.WriteUInt16LittleEndian(baseFields,
            checked((ushort)(style.InvariantStyleId ?? 0x0FFE)));
        var parent = style.BasedOnIndex is int basedOn &&
            (styles.ContainsKey(basedOn) || (style.Type == 3 && basedOn == 11))
            ? basedOn : 0x0FFF;
        var next = style.NextIndex is int nextIndex && styles.ContainsKey(nextIndex)
            ? nextIndex : style.Index;
        BinaryPrimitives.WriteUInt16LittleEndian(baseFields.AsSpan(2),
            checked((ushort)((parent << 4) | style.Type)));
        BinaryPrimitives.WriteUInt16LittleEndian(baseFields.AsSpan(4),
            checked((ushort)((next << 4) | (style.Type == 1 ? 2 : style.Type == 3 ? 3 : 1))));
        if (style.Type == 2 && style.Name == PreservedRunDefaultsStyleName)
            // GRFSTD.fHidden and fSemiHidden keep the internal defaults
            // carrier out of Word's style picker.
            BinaryPrimitives.WriteUInt16LittleEndian(baseFields.AsSpan(8), 0x0102);
        if (baseSize == 18 && style.LinkedStyleIndex is int linked &&
            linked != 0 && styles.ContainsKey(linked))
            BinaryPrimitives.WriteUInt16LittleEndian(baseFields.AsSpan(10),
                checked((ushort)linked));
        using var output = new MemoryStream();
        output.Write(baseFields, 0, baseFields.Length);
        var name = Encoding.Unicode.GetBytes(style.Name);
        WriteU16(output, checked((ushort)style.Name.Length));
        output.Write(name, 0, name.Length);
        WriteU16(output, 0);
        if (style.Type == 3)
        {
            using var tapx = new MemoryStream();
            tapx.Write((style.TableFormatting ?? DocParagraphFormatting.Empty).Encode());
            foreach (var condition in (style.ConditionalTableShading?.Keys ??
                Enumerable.Empty<ushort>()).Concat(
                style.ConditionalTableBorders?.Keys ?? Enumerable.Empty<ushort>())
                .Concat(style.ConditionalTableVerticalAlignment?.Keys ??
                    Enumerable.Empty<ushort>())
                .Concat(style.ConditionalTableNoWrap?.Keys ??
                    Enumerable.Empty<ushort>())
                .Distinct().OrderBy(x => x))
                {
                    using var inner = new MemoryStream();
                    if (style.ConditionalTableShading?.TryGetValue(condition,
                        out var shading) == true)
                        inner.Write((DocParagraphFormatting.Empty with
                        { TableStyleShading = shading }).Encode());
                    if (style.ConditionalTableBorders?.TryGetValue(condition,
                        out var borders) == true)
                        inner.Write(borders.Encode());
                    if (style.ConditionalTableVerticalAlignment?.TryGetValue(
                        condition, out var verticalAlignment) == true)
                        inner.Write((DocParagraphFormatting.Empty with
                        { TableStyleVerticalAlignment = verticalAlignment }).Encode());
                    if (style.ConditionalTableNoWrap?.TryGetValue(
                        condition, out var noWrap) == true)
                        inner.Write((DocParagraphFormatting.Empty with
                        { TableStyleNoWrap = noWrap }).Encode());
                    var modifiers = inner.ToArray();
                    if (modifiers.Length + 2 > byte.MaxValue)
                        throw new InvalidDataException("A conditional table-style rule is too large.");
                    WriteU16(tapx, 0xD66A);
                    tapx.WriteByte(checked((byte)(modifiers.Length + 2)));
                    WriteU16(tapx, condition);
                    tapx.Write(modifiers);
                }
            WriteUpx(output, tapx.ToArray());
        }
        if (style.Type is 1 or 3)
        {
            var properties = (style.ParagraphFormatting ?? DocParagraphFormatting.Empty).Encode(forStyle: true);
            using var conditional = new MemoryStream();
            if (style.Type == 3 && style.ConditionalParagraphFormatting != null)
                foreach (var entry in style.ConditionalParagraphFormatting
                    .OrderBy(x => x.Key))
                {
                    var condition = entry.Key;
                    var formatting = entry.Value;
                    var modifiers = formatting.Encode(forStyle: true);
                    if (modifiers.Length == 0) continue;
                    if (modifiers.Length + 2 > byte.MaxValue)
                        throw new InvalidDataException("A conditional table-style paragraph rule is too large.");
                    WriteU16(conditional, 0xC666);
                    conditional.WriteByte(checked((byte)(modifiers.Length + 2)));
                    WriteU16(conditional, condition);
                    conditional.Write(modifiers);
                }
            var conditionalBytes = conditional.ToArray();
            var papx = new byte[2 + properties.Length + conditionalBytes.Length];
            BinaryPrimitives.WriteUInt16LittleEndian(papx, checked((ushort)style.Index));
            properties.CopyTo(papx, 2);
            conditionalBytes.CopyTo(papx, 2 + properties.Length);
            WriteUpx(output, papx);
        }
        using var chpx = new MemoryStream();
        chpx.Write(DocPlainTextWriter.EncodeCharacterProperties(
            style.CharacterFormatting, fontIndexes));
        if (style.Type == 3 && style.ConditionalCharacterFormatting != null)
            foreach (var entry in style.ConditionalCharacterFormatting
                .OrderBy(x => x.Key))
            {
                var condition = entry.Key;
                var formatting = entry.Value;
                var modifiers = DocPlainTextWriter.EncodeCharacterProperties(
                    formatting, fontIndexes);
                if (modifiers.Length == 0) continue;
                if (modifiers.Length + 2 > byte.MaxValue)
                    throw new InvalidDataException("A conditional table-style run rule is too large.");
                WriteU16(chpx, 0xCA85);
                chpx.WriteByte(checked((byte)(modifiers.Length + 2)));
                WriteU16(chpx, condition);
                chpx.Write(modifiers);
            }
        WriteUpx(output, chpx.ToArray());
        var bytes = output.ToArray();
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(6),
            checked((ushort)bytes.Length));
        return bytes;
    }

    private static void WriteUpx(Stream output, byte[] content)
    {
        WriteU16(output, checked((ushort)content.Length));
        output.Write(content, 0, content.Length);
        if ((content.Length & 1) != 0) output.WriteByte(0);
    }

    private static void WriteU16(Stream output, ushort value)
    {
        output.WriteByte((byte)value);
        output.WriteByte((byte)(value >> 8));
    }

    public static byte[] CreateSectionTable(int characterCount)
    {
        // PlcfSed with one section and no exception properties (fcSepx=-1).
        return new DocSectionTable(checked((uint)characterCount), -1).Write();
    }
}
