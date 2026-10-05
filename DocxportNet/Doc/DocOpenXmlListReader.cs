using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace DocxportNet.Doc;

internal static class DocOpenXmlListReader
{
    public static (DocListIndex Lists, IReadOnlyDictionary<int, short> NumberIds)
        Read(MainDocumentPart? main,
            Func<OpenXmlCompositeElement?, DocCharacterFormatting> readFormatting,
            Func<OpenXmlCompositeElement?, DocParagraphFormatting> readParagraphFormatting)
    {
        var numbering = main?.NumberingDefinitionsPart?.Numbering;
        if (numbering == null)
            return (new DocListIndex([], []), new Dictionary<int, short>());
        var usedNumberIds = new HashSet<int>();
        void Collect(IEnumerable<NumberingProperties> properties)
        {
            foreach (var item in properties)
                if (item.NumberingId?.Val?.Value is int id && id != 0)
                    usedNumberIds.Add(id);
        }
        if (main?.Document != null)
            Collect(main.Document.Descendants<NumberingProperties>());
        if (main?.StyleDefinitionsPart?.Styles != null)
            Collect(main.StyleDefinitionsPart.Styles.Descendants<NumberingProperties>());
        foreach (var header in main?.HeaderParts ?? [])
            if (header.Header != null)
                Collect(header.Header.Descendants<NumberingProperties>());
        foreach (var footer in main?.FooterParts ?? [])
            if (footer.Footer != null)
                Collect(footer.Footer.Descendants<NumberingProperties>());
        if (usedNumberIds.Count == 0)
            return (new DocListIndex([], []), new Dictionary<int, short>());
        var usedAbstractIds = new HashSet<int>(numbering.Elements<NumberingInstance>()
            .Where(x => x.NumberID?.Value is int id && usedNumberIds.Contains(id))
            .Select(x => x.AbstractNumId?.Val?.Value)
            .Where(x => x != null).Select(x => x!.Value));
        var definitions = new List<DocListDefinition>();
        var listIds = new Dictionary<int, int>();
        foreach (var abstractNum in numbering.Elements<AbstractNum>().Where(x =>
            x.AbstractNumberId?.Value is int id && usedAbstractIds.Contains(id)))
        {
            if (abstractNum.AbstractNumberId?.Value is not int sourceId)
                throw new InvalidDataException("A DOCX abstract list has no ID.");
            var listId = checked(definitions.Count + 1);
            if (listIds.ContainsKey(sourceId))
                throw new InvalidDataException("A DOCX abstract list ID is duplicated.");
            listIds.Add(sourceId, listId);
            var levels = abstractNum.Elements<Level>().Select(x => ReadLevel(x, readFormatting,
                readParagraphFormatting))
                .OrderBy(x => x.Index).ToArray();
            if (levels.Length == 0 || levels.Length > 9 ||
                levels.Select(x => x.Index).Distinct().Count() != levels.Length ||
                levels.Where((x, i) => x.Index != i).Any())
                throw new NotSupportedException("DOC list levels must be contiguous from zero.");
            definitions.Add(new DocListDefinition(listId, levels));
        }
        var instances = new List<DocListInstance>();
        var numberIds = new Dictionary<int, short>();
        foreach (var number in numbering.Elements<NumberingInstance>().Where(x =>
            x.NumberID?.Value is int id && usedNumberIds.Contains(id)))
        {
            if (number.NumberID?.Value is not int sourceId ||
                number.AbstractNumId?.Val?.Value is not int abstractId ||
                !listIds.TryGetValue(abstractId, out var listId))
                throw new InvalidDataException("A DOCX numbered list has no abstract list.");
            var overrideIndex = checked((short)(instances.Count + 1));
            if (numberIds.ContainsKey(sourceId))
                throw new InvalidDataException("A DOCX numbered list ID is duplicated.");
            numberIds.Add(sourceId, overrideIndex);
            var starts = new Dictionary<int, int>();
            var formats = new Dictionary<int, DocListLevel>();
            foreach (var level in number.Elements<LevelOverride>())
            {
                if (level.LevelIndex?.Value is not int levelIndex || levelIndex is < 0 or > 8)
                    throw new InvalidDataException("A DOCX list override has an invalid level.");
                if (level.StartOverrideNumberingValue?.Val?.Value is int start)
                    starts[levelIndex] = start;
                if (level.Level != null)
                    formats[levelIndex] = ReadLevel(level.Level, readFormatting,
                        readParagraphFormatting);
            }
            instances.Add(new DocListInstance(overrideIndex, listId, starts, formats));
        }
        return (new DocListIndex(definitions, instances), numberIds);
    }

    private static DocListLevel ReadLevel(Level level,
        Func<OpenXmlCompositeElement?, DocCharacterFormatting> readFormatting,
        Func<OpenXmlCompositeElement?, DocParagraphFormatting> readParagraphFormatting)
    {
        if (level.LevelIndex?.Value is not int index || index is < 0 or > 8)
            throw new InvalidDataException("A DOCX list level has an invalid index.");
        var format = level.NumberingFormat?.Val?.Value;
        byte code = format == null || format == NumberFormatValues.Decimal ? (byte)0 :
            format == NumberFormatValues.UpperRoman ? (byte)1 :
            format == NumberFormatValues.LowerRoman ? (byte)2 :
            format == NumberFormatValues.UpperLetter ? (byte)3 :
            format == NumberFormatValues.LowerLetter ? (byte)4 :
            format == NumberFormatValues.Ordinal ? (byte)5 :
            format == NumberFormatValues.CardinalText ? (byte)6 :
            format == NumberFormatValues.OrdinalText ? (byte)7 :
            format == NumberFormatValues.Hex ? (byte)8 :
            format == NumberFormatValues.IdeographDigital ? (byte)0x0A :
            format == NumberFormatValues.DecimalFullWidth ? (byte)0x0E :
            format == NumberFormatValues.DecimalHalfWidth ? (byte)0x0F :
            format == NumberFormatValues.DecimalEnclosedCircle ? (byte)0x12 :
            format == NumberFormatValues.DecimalZero ? (byte)0x16 :
            format == NumberFormatValues.Bullet ? (byte)0x17 :
            format == NumberFormatValues.NumberInDash ? (byte)0x39 :
            format == NumberFormatValues.RussianLower ? (byte)0x3A :
            format == NumberFormatValues.RussianUpper ? (byte)0x3B :
            format == NumberFormatValues.None ? (byte)0xFF :
            throw new NotSupportedException($"DOCX list format '{format}' is unsupported.");
        var suffix = level.LevelSuffix?.Val?.Value;
        byte follow = suffix == null || suffix == LevelSuffixValues.Tab ? (byte)0 :
            suffix == LevelSuffixValues.Space ? (byte)1 :
            suffix == LevelSuffixValues.Nothing ? (byte)2 :
            throw new NotSupportedException("The DOCX list label suffix is unsupported.");
        var alignment = level.LevelJustification?.Val?.Value;
        byte justification = alignment == null || alignment == LevelJustificationValues.Left
            ? (byte)0 : alignment == LevelJustificationValues.Center ? (byte)1 :
            alignment == LevelJustificationValues.Right ? (byte)2 :
            throw new NotSupportedException("The DOCX list label alignment is unsupported.");
        var paragraphFormatting = readParagraphFormatting(level.PreviousParagraphProperties);
        // A DOC list's tab suffix follows the paragraph's tab stops. Word's
        // Header and Footer styles contain center/right tabs, so give the
        // label an explicit stop at its text indent when OOXML relies on the
        // implicit numbering tab.
        if (follow == 0 && paragraphFormatting.LeftTwips is short textIndent &&
            (paragraphFormatting.TabStops == null ||
             paragraphFormatting.TabStops.Count == 0))
            paragraphFormatting = paragraphFormatting with
            {
                TabStops = new[] { new DocTabStop(textIndent, 6, 0) }
            };
        var restartValue = level.GetFirstChild<LevelRestart>()?.Val?.Value;
        if (level.GetFirstChild<LevelRestart>() != null &&
            restartValue is not (>= 0 and <= 7))
            throw new NotSupportedException("A list restart level must be between 0 and 7.");
        return new DocListLevel(index,
            level.StartNumberingValue?.Val?.Value ?? 1, code,
            level.LevelText?.Val?.Value ?? string.Empty, follow,
            readFormatting(level.NumberingSymbolRunProperties),
            paragraphFormatting, justification,
            restartValue is int restart ? checked((byte)restart) : null);
    }
}
