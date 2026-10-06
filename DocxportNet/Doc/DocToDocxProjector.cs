using System.Buffers.Binary;
using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxportNet.Doc;

/// <summary>Reports which indexed content was represented by a basic DOCX projection.</summary>
public sealed record DocxProjectionCoverage(
    DocPartRange? ProjectedMainPart,
    IReadOnlyList<DocPartRange> DeferredParts,
    IReadOnlyList<DocLocation> DeferredLocations,
    IReadOnlyList<string> DeferredStreams,
    IReadOnlyDictionary<ushort, int> OmittedCharacters,
    IReadOnlyDictionary<ushort, int> ApproximateCharacters);

public sealed record DocxProjectionResult(byte[] DocxBytes, DocxProjectionCoverage Coverage);

/// <summary>Projects indexed story text into a minimal WordprocessingML package.</summary>
public sealed class DocToDocxProjector
{
    public DocxProjectionResult Project(DocTextIndex index)
    {
        if (index == null) throw new ArgumentNullException(nameof(index));
        var main = index.Parts.FirstOrDefault(x => x.Name == "Main");
        var indexedMain = main == null ? null : index.ReadStory("Main");
        var story = indexedMain?.Text;

        var omitted = new SortedDictionary<ushort, int>();
        var approximate = new SortedDictionary<ushort, int>();
        var styles = index.StyleDefinitions;
        var paragraphStyles = index.ParagraphStyles;
        var usedLists = paragraphStyles.Where(x => x.Formatting?.ListOverrideIndex > 0)
            .Select(x => x.Formatting!.ListOverrideIndex!.Value)
            .Concat(styles.Where(x => x.Type == 1 &&
                x.ParagraphFormatting?.ListOverrideIndex > 0)
                .Select(x => x.ParagraphFormatting!.ListOverrideIndex!.Value))
            .Distinct().ToArray();
        var styleByIndex = styles.ToDictionary(x => x.Index);
        var pictures = new Dictionary<int, DocInlinePicture>();
        var projectedHeaders = false;
        var needsModernFloatingStoryCompatibility = false;
        using var output = new MemoryStream();
        using (var document = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document, true))
        {
            if (DocSummaryInformation.ReadTitle(index.Structure) is { } title)
                document.PackageProperties.Title = title;
            if (DocSummaryInformation.ReadSubject(index.Structure) is { } subject)
                document.PackageProperties.Subject = subject;
            if (DocSummaryInformation.ReadAuthor(index.Structure) is { } author)
                document.PackageProperties.Creator = author;
            if (DocSummaryInformation.ReadKeywords(index.Structure) is { } keywords)
                document.PackageProperties.Keywords = keywords;
            if (DocSummaryInformation.ReadComments(index.Structure) is { } comments)
                document.PackageProperties.Description = comments;
            if (DocSummaryInformation.ReadLastAuthor(index.Structure) is { } lastAuthor)
                document.PackageProperties.LastModifiedBy = lastAuthor;
            if (DocSummaryInformation.ReadRevisionNumber(index.Structure) is { } revision)
                document.PackageProperties.Revision = revision;
            var pages = DocSummaryInformation.ReadPageCount(index.Structure);
            var words = DocSummaryInformation.ReadWordCount(index.Structure);
            var characters = DocSummaryInformation.ReadCharacterCount(index.Structure);
            if (pages != null || words != null || characters != null)
            {
                var extended = document.AddNewPart<ExtendedFilePropertiesPart>();
                var properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                if (pages is int pageCount)
                    properties.Append(new DocumentFormat.OpenXml.ExtendedProperties.Pages
                    { Text = pageCount.ToString(CultureInfo.InvariantCulture) });
                if (words is int wordCount)
                    properties.Append(new DocumentFormat.OpenXml.ExtendedProperties.Words
                    { Text = wordCount.ToString(CultureInfo.InvariantCulture) });
                if (characters is int characterCount)
                    properties.Append(new DocumentFormat.OpenXml.ExtendedProperties.Characters
                    { Text = characterCount.ToString(CultureInfo.InvariantCulture) });
                extended.Properties = properties;
                properties.Save();
            }
            var part = document.AddMainDocumentPart();
            if (styles.Count != 0 || !index.DefaultCharacterFormatting.IsEmpty)
                WriteStyles(part, styles, index.DefaultCharacterFormatting);
            IReadOnlyList<DocFontDefinition> fonts;
            try { fonts = index.Fonts; }
            catch (InvalidDataException) { fonts = []; }
            if (fonts.Count != 0)
                WriteFontTable(part, fonts);
            if (usedLists.Length != 0)
                WriteNumbering(part, index.Lists, usedLists, styleByIndex);
            var body = new Body();
            var mainParagraphs = RenderParagraphs(story, indexedMain?.CharacterFormatting ?? [],
                indexedMain?.ParagraphStyles ?? [], styleByIndex,
                indexedMain?.Bookmarks ?? [], indexedMain?.InlinePictures ??
                    new Dictionary<int, DocStoryInlinePicture>(),
                index, part, pictures, omitted, approximate,
                indexedMain?.FloatingPictures);
            SectionProperties? finalSectionProperties = null;

            var headerPart = index.Parts.FirstOrDefault(x => x.Name == "Headers");
            var sections = index.StorySections;
            if (sections.Count > 0)
            {
                var headerStories = headerPart == null ? null : index.HeaderStories;
                var hasHeaderStories = headerStories?.Count == 6 + sections.Count * 6;
                if (headerPart != null && !hasHeaderStories)
                    throw new InvalidDataException(
                        "The DOC header story slots do not match its sections.");
                if (headerPart == null || hasHeaderStories)
                {
                    for (var sectionIndex = 0; sectionIndex < sections.Count; sectionIndex++)
                    {
                        var sectionProperties = new SectionProperties();
                        for (var slotInSection = 0; hasHeaderStories && slotInSection < 6; slotInSection++)
                        {
                            var headerStory = headerStories![6 + sectionIndex * 6 + slotInSection];
                            if (!sections[sectionIndex].Slots[slotInSection].IsPresent) continue;
                            // Empty slots inherit or remain absent.
                            var indexedHeader = headerStory.ReadIndexedContent();
                            // DOC compatibility layout reflows header/footer text
                            // around anchored pictures differently from modern DOCX.
                            if (indexedHeader.FloatingPictures.Values.Any(group =>
                                group.Any(picture => picture.WrapCode != 3)))
                                needsModernFloatingStoryCompatibility = true;
                            var headerContent = indexedHeader.Text;
                            OpenXmlPart contentPart = slotInSection is 0 or 1 or 4
                                ? part.AddNewPart<HeaderPart>()
                                : part.AddNewPart<FooterPart>();
                            var paragraphs = RenderParagraphs(headerContent,
                                indexedHeader.CharacterFormatting, indexedHeader.ParagraphStyles, styleByIndex,
                                indexedHeader.Bookmarks, indexedHeader.InlinePictures,
                                index, contentPart, pictures, omitted, approximate,
                                indexedHeader.FloatingPictures);
                            ApplyStyleRunFormatting(paragraphs, styleByIndex);
                            var blocks = RenderBlocks(headerContent, paragraphs,
                                indexedHeader.ParagraphStyles, indexedHeader.TableRows,
                                styleByIndex, index.GrowAutofit == true);
                            if (slotInSection is 0 or 1 or 4)
                            {
                                var header = (HeaderPart)contentPart;
                                header.Header = new Header(blocks);
                                header.Header.Save();
                                sectionProperties.AppendChild(new HeaderReference
                                {
                                    Type = ReferenceType(slotInSection), Id = part.GetIdOfPart(header)
                                });
                            }
                            else
                            {
                                var footer = (FooterPart)contentPart;
                                footer.Footer = new Footer(blocks);
                                footer.Footer.Save();
                                sectionProperties.AppendChild(new FooterReference
                                {
                                    Type = ReferenceType(slotInSection), Id = part.GetIdOfPart(footer)
                                });
                            }
                        }
                        sections[sectionIndex].Formatting.ApplyTo(sectionProperties);
                        if (sectionIndex == sections.Count - 1)
                            finalSectionProperties = sectionProperties;
                        else
                        {
                            var sectionEnd = checked((uint)sections[sectionIndex].EndCp);
                            var paragraphIndex = story!.Paragraphs.ToList()
                                .FindIndex(x => x.CpEnd == sectionEnd);
                            if (paragraphIndex < 0)
                                throw new InvalidDataException("A DOC section boundary has no paragraph mark.");
                            var paragraphProperties = mainParagraphs[paragraphIndex].ParagraphProperties;
                            if (paragraphProperties == null)
                                mainParagraphs[paragraphIndex].ParagraphProperties =
                                    new ParagraphProperties(sectionProperties);
                            else
                                paragraphProperties.AppendChild(sectionProperties);
                        }
                    }
                    if (hasHeaderStories && (index.EvenAndOddHeaders ?? headerStories!.Any(x => !x.IsEmpty && x.Kind is
                        DocHeaderStoryKind.EvenHeader or DocHeaderStoryKind.EvenFooter)))
                    {
                        var settingsPart = part.AddNewPart<DocumentSettingsPart>();
                        settingsPart.Settings = new Settings(new EvenAndOddHeaders());
                        settingsPart.Settings.Save();
                    }
                    projectedHeaders = hasHeaderStories;
                }
            }
            ApplyStyleRunFormatting(mainParagraphs, styleByIndex);
            foreach (var block in RenderBlocks(story, mainParagraphs,
                indexedMain?.ParagraphStyles ?? [], indexedMain?.TableRows ?? [],
                styleByIndex, index.GrowAutofit == true))
                body.AppendChild(block);
            if (mainParagraphs.Count == 0) body.AppendChild(new Paragraph());
            if (finalSectionProperties != null) body.AppendChild(finalSectionProperties);
            part.Document = new Document(body);
            part.Document.Save();
            var autoWidthRows = index.ParagraphStyles.Where(x =>
                x.Formatting?.TableAutoFit == true &&
                x.Formatting.TableCellPreferredWidths?.Any(w => w?.Unit == 1) == true)
                .ToArray();
            // Spaced automatic cells use Word's DOC layout when projected.
            // Mode 15 expands their native fitted edges, while mode 11
            // preserves both the native and writer-produced appearance.
            var needsLegacySpacedAutoFitCompatibility = index.GrowAutofit != true &&
                autoWidthRows.Length > 0 && autoWidthRows.All(x =>
                    x.Formatting?.TableCellSpacingTwips is > 0);
            var needsModernAutoFitCompatibility = index.GrowAutofit != true &&
                autoWidthRows.Length > 0 && !needsLegacySpacedAutoFitCompatibility;
            // DOCX opens without a settings part in an older layout mode. An
            // applied conditional border style needs modern table placement;
            // otherwise Word shifts its table edges despite matching cell text.
            bool HasConditionalTableBorders(int styleIndex)
            {
                var visited = new HashSet<int>();
                while (visited.Add(styleIndex) &&
                    styleByIndex.TryGetValue(styleIndex, out var style))
                {
                    if (style.ConditionalTableBorders?.Count > 0) return true;
                    if (style.BasedOnIndex is not int parent) break;
                    styleIndex = parent;
                }
                return false;
            }
            // Word's DOC save can flatten the applied table style into row
            // borders while retaining its conditional style definitions.
            // Those rows still need the modern placement mode on projection.
            var tableRows = index.ParagraphStyles.Where(x =>
                x.Formatting?.TableTerminator == true).ToArray();
            // A native DOC may also flatten the style definition itself. An
            // alternating A/B/A run of colored top edges retains its visible
            // horizontal-band pattern but needs the same modern placement.
            bool HasFlattenedHorizontalBandBorders()
            {
                for (var i = 2; i < tableRows.Length; i++)
                {
                    var first = tableRows[i - 2].Formatting?.TableCellBorders?
                        .FirstOrDefault()?.Top;
                    var middle = tableRows[i - 1].Formatting?.TableCellBorders?
                        .FirstOrDefault()?.Top;
                    var last = tableRows[i].Formatting?.TableCellBorders?
                        .FirstOrDefault()?.Top;
                    if (first?.ColorRgb is uint color &&
                        middle?.ColorRgb is uint alternate &&
                        last?.ColorRgb == color && alternate != color &&
                        first.WidthEighthPoints > 0 && middle.WidthEighthPoints > 0 &&
                        last.WidthEighthPoints > 0)
                        return true;
                }
                return false;
            }
            var needsModernConditionalTableStyleCompatibility = tableRows.Length > 0 &&
                (tableRows.Any(x => x.Formatting?.TableStyleIndex is
                    ushort styleIndex && HasConditionalTableBorders(styleIndex)) ||
                 styles.Any(x => x.ConditionalTableBorders?.Count > 0) ||
                 HasFlattenedHorizontalBandBorders());
            if (index.MirrorMargins == true || index.GutterAtTop == true ||
                index.AutoHyphenation == true ||
                index.HyphenateCaps == false ||
                index.HyphenationZoneTwips is > 0 ||
                index.ConsecutiveHyphenLimit is > 0 ||
                index.BalanceSingleByteDoubleByteWidth == true ||
                index.GrowAutofit == true || index.ApplyBreakingRules == true || needsModernAutoFitCompatibility ||
                needsLegacySpacedAutoFitCompatibility || needsModernFloatingStoryCompatibility ||
                needsModernConditionalTableStyleCompatibility ||
                index.DefaultTabStopTwips is short tabInterval && tabInterval != 720)
            {
                var settingsPart = part.DocumentSettingsPart ??
                    part.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings ??= new Settings();
                if (index.MirrorMargins == true)
                    settingsPart.Settings.AddChild(new MirrorMargins(), true);
                if (index.AutoHyphenation == true)
                    settingsPart.Settings.AddChild(new AutoHyphenation(), true);
                if (index.HyphenateCaps == false)
                    settingsPart.Settings.AddChild(new DoNotHyphenateCaps(), true);
                if (index.HyphenationZoneTwips is > 0)
                    settingsPart.Settings.AddChild(new HyphenationZone
                        { Val = index.HyphenationZoneTwips.Value.ToString(
                            System.Globalization.CultureInfo.InvariantCulture) }, true);
                if (index.ConsecutiveHyphenLimit is > 0)
                    settingsPart.Settings.AddChild(new ConsecutiveHyphenLimit
                        { Val = checked((ushort)index.ConsecutiveHyphenLimit.Value) }, true);
                if (index.GutterAtTop == true)
                    settingsPart.Settings.AddChild(new GutterAtTop(), true);
                if (index.BalanceSingleByteDoubleByteWidth == true ||
                    index.GrowAutofit == true || index.ApplyBreakingRules == true ||
                    needsModernAutoFitCompatibility || needsLegacySpacedAutoFitCompatibility ||
                    needsModernFloatingStoryCompatibility ||
                    needsModernConditionalTableStyleCompatibility)
                {
                    var compatibility = new Compatibility();
                    if (index.BalanceSingleByteDoubleByteWidth == true)
                        compatibility.AddChild(new BalanceSingleByteDoubleByteWidth(), true);
                    if (index.GrowAutofit == true)
                        compatibility.AddChild(new GrowAutofit(), true);
                    if (index.ApplyBreakingRules == true)
                        compatibility.AddChild(new ApplyBreakingRules(), true);
                    if (needsModernAutoFitCompatibility || needsLegacySpacedAutoFitCompatibility ||
                        needsModernFloatingStoryCompatibility ||
                        needsModernConditionalTableStyleCompatibility)
                        compatibility.AddChild(new CompatibilitySetting
                        {
                            Name = CompatSettingNameValues.CompatibilityMode,
                            Uri = "http://schemas.microsoft.com/office/word",
                            Val = needsModernAutoFitCompatibility ||
                                needsModernFloatingStoryCompatibility ||
                                needsModernConditionalTableStyleCompatibility ? "15" : "11"
                        }, true);
                    settingsPart.Settings.AddChild(compatibility, true);
                }
                if (index.DefaultTabStopTwips is short interval && interval != 720)
                    settingsPart.Settings.AddChild(new DefaultTabStop { Val = interval }, true);
                settingsPart.Settings.Save();
            }
        }

        var coverage = new DocxProjectionCoverage(main,
            index.Parts.Where(x => x.Name != "Main" &&
                !(projectedHeaders && x.Name == "Headers")).ToArray(),
            index.Structure.Locations.Where(x => x.IsPresent && x.FibIndex != 33 &&
                !(projectedHeaders && x.Name == "HeadersAndFooters")).ToArray(),
            EnumerateStreams(index.Structure.Root)
                .Where(x => x != "WordDocument" && x != index.Structure.FibBase.TableStreamName)
                .ToArray(),
            omitted, approximate);
        return new DocxProjectionResult(output.ToArray(), coverage);
    }

    private static void WriteNumbering(MainDocumentPart main, DocListIndex lists,
        IReadOnlyList<short> usedIndexes,
        IReadOnlyDictionary<int, DocStyleDefinition> styles)
    {
        var instances = lists.Instances.ToDictionary(x => x.OverrideIndex);
        var definitions = lists.Definitions.ToDictionary(x => x.ListId);
        var numbering = new Numbering();
        var selected = new List<(DocListInstance Instance, DocListDefinition Definition,
            int AbstractId)>();
        foreach (var index in usedIndexes.OrderBy(x => x))
        {
            if (!instances.TryGetValue(index, out var instance) ||
                !definitions.TryGetValue(instance.ListId, out var definition))
                throw new InvalidDataException($"DOC list override {index} has no definition.");
            selected.Add((instance, definition, selected.Count));
        }
        foreach (var (instance, definition, abstractId) in selected)
        {
            var abstractNum = new AbstractNum { AbstractNumberId = abstractId };
            foreach (var baseLevel in definition.Levels)
            {
                var level = instance.FormattingOverrides.TryGetValue(baseLevel.Index,
                    out var replacement) ? replacement : baseLevel;
                var format = level.NumberFormat switch
                {
                    0 => NumberFormatValues.Decimal,
                    1 => NumberFormatValues.UpperRoman,
                    2 => NumberFormatValues.LowerRoman,
                    3 => NumberFormatValues.UpperLetter,
                    4 => NumberFormatValues.LowerLetter,
                    5 => NumberFormatValues.Ordinal,
                    6 => NumberFormatValues.CardinalText,
                    7 => NumberFormatValues.OrdinalText,
                    8 => NumberFormatValues.Hex,
                    0x0A => NumberFormatValues.IdeographDigital,
                    0x09 => NumberFormatValues.Chicago,
                    0x0E => NumberFormatValues.DecimalFullWidth,
                    0x0F => NumberFormatValues.DecimalHalfWidth,
                    0x12 => NumberFormatValues.DecimalEnclosedCircle,
                    0x13 => NumberFormatValues.DecimalFullWidth2,
                    0x16 => NumberFormatValues.DecimalZero,
                    0x17 => NumberFormatValues.Bullet,
                    0x39 => NumberFormatValues.NumberInDash,
                    0x3A => NumberFormatValues.RussianLower,
                    0x3B => NumberFormatValues.RussianUpper,
                    0xFF => NumberFormatValues.None,
                    _ => throw new NotSupportedException(
                        $"DOC list number format {level.NumberFormat} is unsupported.")
                };
                if (level.NumberText.Any(x => x < 0x20 && x is not ('\t' or '\n' or '\r')))
                    throw new NotSupportedException($"DOC list {instance.OverrideIndex} level {level.Index} format {level.NumberFormat} has unsupported number text {string.Join(",", level.NumberText.Select(x => $"U+{(int)x:X4}"))}.");
                var suffix = level.FollowCharacter switch
                {
                    0 => LevelSuffixValues.Tab,
                    1 => LevelSuffixValues.Space,
                    2 => LevelSuffixValues.Nothing,
                    _ => throw new InvalidDataException("A DOC list has an invalid label suffix.")
                };
                var outputLevel = new Level(
                    new StartNumberingValue { Val = level.StartAt },
                    new NumberingFormat { Val = format },
                    new LevelSuffix { Val = suffix },
                    new LevelText { Val = level.NumberText },
                    new LevelJustification { Val = level.LabelJustification switch
                    {
                        0 => LevelJustificationValues.Left,
                        1 => LevelJustificationValues.Center,
                        2 => LevelJustificationValues.Right,
                        _ => throw new InvalidDataException("A DOC list label has invalid alignment.")
                    } })
                    { LevelIndex = level.Index };
                if (level.RestartLimit is byte restartLimit)
                    outputLevel.InsertAfter(new LevelRestart { Val = restartLimit },
                        outputLevel.GetFirstChild<NumberingFormat>());
                if (level.ParagraphFormatting is { IsEmpty: false } paragraphFormatting)
                {
                    var properties = new PreviousParagraphProperties();
                    AppendParagraphFormatting(properties, paragraphFormatting);
                    if (properties.HasChildren) outputLevel.AppendChild(properties);
                }
                if (level.LabelFormatting is { IsEmpty: false } labelFormatting)
                {
                    var properties = new NumberingSymbolRunProperties();
                    AppendCharacterFormatting(properties, labelFormatting, styles);
                    if (properties.HasChildren) outputLevel.AppendChild(properties);
                }
                abstractNum.AppendChild(outputLevel);
            }
            numbering.AppendChild(abstractNum);
        }
        foreach (var (instance, _, abstractId) in selected)
        {
            var num = new NumberingInstance(new AbstractNumId { Val = abstractId })
                { NumberID = instance.OverrideIndex };
            foreach (var start in instance.StartOverrides.OrderBy(x => x.Key))
                num.AppendChild(new LevelOverride(
                    new StartOverrideNumberingValue { Val = start.Value })
                    { LevelIndex = start.Key });
            numbering.AppendChild(num);
        }
        var part = main.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering = numbering;
        part.Numbering.Save();
    }

    private static void WriteFontTable(MainDocumentPart main,
        IReadOnlyList<DocFontDefinition> fonts)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var root = new XElement(w + "fonts", new XAttribute(XNamespace.Xmlns + "w", w));
        foreach (var font in fonts)
        {
            var element = new XElement(w + "font", new XAttribute(w + "name", font.Name));
            if (font.AlternateName != null)
                element.Add(new XElement(w + "altName", new XAttribute(w + "val", font.AlternateName)));
            if (font.Panose is { Length: 10 } panose)
                element.Add(new XElement(w + "panose1",
                    new XAttribute(w + "val", DocBinaryCompat.ToHexString(panose))));
            element.Add(new XElement(w + "charset",
                new XAttribute(w + "val", font.Charset.ToString("X2", CultureInfo.InvariantCulture))));
            var family = (font.FamilyPitch >> 4) & 7;
            element.Add(new XElement(w + "family", new XAttribute(w + "val", family switch
            {
                1 => "roman", 2 => "swiss", 3 => "modern", 4 => "script",
                5 => "decorative", _ => "auto"
            })));
            if ((font.FamilyPitch & 4) == 0)
                element.Add(new XElement(w + "notTrueType"));
            element.Add(new XElement(w + "pitch", new XAttribute(w + "val",
                (font.FamilyPitch & 3) switch
                {
                    1 => "fixed", 2 => "variable", _ => "default"
                })));
            if (font.Signature is { Length: 24 } signature)
            {
                var fields = new[] { "usb0", "usb1", "usb2", "usb3", "csb0", "csb1" };
                var sig = new XElement(w + "sig");
                for (var i = 0; i < fields.Length; i++)
                    sig.Add(new XAttribute(w + fields[i],
                        BinaryPrimitives.ReadUInt32LittleEndian(signature.AsSpan(i * 4))
                            .ToString("X8", CultureInfo.InvariantCulture)));
                element.Add(sig);
            }
            root.Add(element);
        }
        var part = main.AddNewPart<FontTablePart>();
        using var stream = new MemoryStream();
        new XDocument(root).Save(stream);
        stream.Position = 0;
        part.FeedData(stream);
    }

    private static void WriteStyles(MainDocumentPart main,
        IReadOnlyList<DocStyleDefinition> definitions,
        DocCharacterFormatting defaults)
    {
        var part = main.AddNewPart<StyleDefinitionsPart>();
        var root = new Styles();
        var normal = definitions.FirstOrDefault(x => x.Index == 0);
        var preservedDefaults = definitions.FirstOrDefault(x => x.Type == 2 &&
            x.Name == DocDefaultStructures.PreservedRunDefaultsStyleName);
        var normalParagraph = normal?.ParagraphFormatting;
        var rootFormatting = preservedDefaults?.CharacterFormatting ??
            normal?.CharacterFormatting ?? defaults;
        rootFormatting = rootFormatting with
        {
            AsciiFontName = rootFormatting.AsciiFontName ?? defaults.AsciiFontName,
            EastAsiaFontName = rootFormatting.EastAsiaFontName ?? defaults.EastAsiaFontName,
            HighAnsiFontName = rootFormatting.HighAnsiFontName ?? defaults.HighAnsiFontName,
            ComplexScriptFontName = rootFormatting.ComplexScriptFontName ??
                defaults.ComplexScriptFontName
        };
        if (preservedDefaults == null && normal?.CharacterFormatting.SizeHalfPoints is
                ushort normalSize && normalSize != 20 &&
            definitions.Any(x => x.Type == 1 && x.Index != 0 &&
                x.BasedOnIndex == null && x.CharacterFormatting.SizeHalfPoints == null))
            // An unbased DOC paragraph style starts from DOC's 10-point
            // character default, even when Normal specifies another size.
            rootFormatting = rootFormatting with { SizeHalfPoints = 20 };
        var documentDefaults = new DocDefaults();
        if (!rootFormatting.IsEmpty)
        {
            var runProperties = new RunPropertiesBaseStyle();
            // w:highlight is a direct run property, not a valid docDefaults
            // rPr child. Its effective DOC value is emitted on visible runs.
            AppendCharacterFormatting(runProperties, rootFormatting with
            {
                HighlightCode = null
            }, new Dictionary<int, DocStyleDefinition>());
            documentDefaults.AppendChild(new RunPropertiesDefault(runProperties));
        }
        if (normalParagraph is { IsEmpty: false })
        {
            var paragraphProperties = new ParagraphPropertiesBaseStyle();
            AppendParagraphFormatting(paragraphProperties, normalParagraph);
            documentDefaults.AppendChild(new ParagraphPropertiesDefault(paragraphProperties));
        }
        if (documentDefaults.HasChildren) root.AppendChild(documentDefaults);
        var byIndex = definitions.ToDictionary(x => x.Index);
        foreach (var styleSource in definitions.Where(x => (x.Type is 1 or 2 or 3) &&
            !(x.Type == 2 && x.Name ==
                DocDefaultStructures.PreservedRunDefaultsStyleName)))
        {
            // Keep inherited values in the DOC index for effective formatting,
            // but write only a style's own UPX properties to styles.xml.
            var source = styleSource with
            {
                CharacterFormatting = styleSource.DirectCharacterFormatting ??
                    styleSource.CharacterFormatting,
                ParagraphFormatting = styleSource.DirectParagraphFormatting ??
                    styleSource.ParagraphFormatting
            };
            var style = new Style
            {
                Type = source.Type switch
                {
                    1 => StyleValues.Paragraph,
                    2 => StyleValues.Character,
                    _ => StyleValues.Table
                },
                StyleId = source.StyleId,
                CustomStyle = source.InvariantStyleId == 0x0FFE
            };
            var aliasSeparator = source.Name.IndexOf(',');
            style.AppendChild(new StyleName
            {
                Val = aliasSeparator < 0 ? source.Name : source.Name.Substring(0, aliasSeparator)
            });
            if (aliasSeparator >= 0)
                style.AppendChild(new Aliases
                {
                    Val = source.Name.Substring(aliasSeparator + 1)
                });
            if (source.BasedOnIndex is int basedOn && byIndex.TryGetValue(basedOn, out var parent))
                style.AppendChild(new BasedOn { Val = parent.StyleId });
            if (source.Type == 1 && source.NextIndex is int next &&
                byIndex.TryGetValue(next, out var nextStyle))
                style.AppendChild(new NextParagraphStyle { Val = nextStyle.StyleId });
            if (source.LinkedStyleIndex is int linked &&
                byIndex.TryGetValue(linked, out var linkedStyle) &&
                linkedStyle.Name != DocDefaultStructures.PreservedRunDefaultsStyleName)
                style.AppendChild(new LinkedStyle { Val = linkedStyle.StyleId });
            if (source.Type == 1 && source.ParagraphFormatting is { IsEmpty: false } paragraphFormatting)
            {
                var paragraphProperties = new StyleParagraphProperties();
                AppendParagraphFormatting(paragraphProperties, paragraphFormatting);
                style.AppendChild(paragraphProperties);
            }
            if (!source.CharacterFormatting.IsEmpty)
            {
                var properties = new StyleRunProperties();
                AppendRunFonts(properties, source.CharacterFormatting);
                if (source.CharacterFormatting.Bold is bool bold)
                    properties.AppendChild(new Bold { Val = bold });
                if (source.CharacterFormatting.ComplexScriptBold is bool complexBold)
                    properties.AppendChild(new BoldComplexScript { Val = complexBold });
                if (source.CharacterFormatting.Italic is bool italic)
                    properties.AppendChild(new Italic { Val = italic });
                if (source.CharacterFormatting.ComplexScriptItalic is bool complexItalic)
                    properties.AppendChild(new ItalicComplexScript { Val = complexItalic });
                if (source.CharacterFormatting.Caps is bool caps)
                    properties.AppendChild(new Caps { Val = caps });
                if (source.CharacterFormatting.SmallCaps is bool smallCaps)
                    properties.AppendChild(new SmallCaps { Val = smallCaps });
                if (source.CharacterFormatting.Strike is bool strike)
                    properties.AppendChild(new Strike { Val = strike });
                if (source.CharacterFormatting.DoubleStrike is bool doubleStrike)
                    properties.AppendChild(new DoubleStrike { Val = doubleStrike });
                if (source.CharacterFormatting.Outline is bool outline)
                    properties.AppendChild(new Outline { Val = outline });
                if (source.CharacterFormatting.Shadow is bool shadow)
                    properties.AppendChild(new Shadow { Val = shadow });
                if (source.CharacterFormatting.Emboss is bool emboss)
                    properties.AppendChild(new Emboss { Val = emboss });
                if (source.CharacterFormatting.Imprint is bool imprint)
                    properties.AppendChild(new Imprint { Val = imprint });
                if (source.CharacterFormatting.Hidden is bool hidden)
                    properties.AppendChild(new Vanish { Val = hidden });
                if (source.CharacterFormatting.SnapToGrid is bool snap)
                    properties.AppendChild(new SnapToGrid { Val = snap });
                if (ColorValue(source.CharacterFormatting.ColorRef) is string color)
                    properties.AppendChild(new Color { Val = color });
                AppendCharacterSpacing(properties, source.CharacterFormatting);
                if (source.CharacterFormatting.CharacterScalePercent is ushort scale)
                    properties.AppendChild(new CharacterScale { Val = scale });
                AppendKerning(properties, source.CharacterFormatting);
                if (source.CharacterFormatting.BaselineOffsetHalfPoints is short offset)
                    properties.AppendChild(new Position { Val = offset.ToString(
                        System.Globalization.CultureInfo.InvariantCulture) });
                if (source.CharacterFormatting.SizeHalfPoints is ushort size && size > 0)
                    properties.AppendChild(new FontSize
                    {
                        Val = size.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    });
                if (source.CharacterFormatting.ComplexScriptSizeHalfPoints is ushort complexSize &&
                    complexSize > 0)
                    properties.AppendChild(new FontSizeComplexScript
                    {
                        Val = complexSize.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    });
                AppendUnderline(properties, source.CharacterFormatting);
                AppendRunBorder(properties, source.CharacterFormatting);
                AppendRunShading(properties, source.CharacterFormatting);
                AppendFitText(properties, source.CharacterFormatting);
                if (ScriptValue(source.CharacterFormatting.ScriptCode) is VerticalPositionValues script)
                    properties.AppendChild(new VerticalTextAlignment { Val = script });
                // w:rtl is not valid in style rPr. Applied character styles
                // carry it onto their runs below.
                if (source.CharacterFormatting.ForceComplexScript is bool forceComplexScript)
                    properties.AppendChild(new ComplexScript { Val = forceComplexScript });
                if (EmphasisValue(source.CharacterFormatting.EmphasisMarkCode)
                    is EmphasisMarkValues emphasis)
                    properties.AppendChild(new Emphasis { Val = emphasis });
                AppendLanguages(properties, source.CharacterFormatting);
                style.AppendChild(properties);
            }
            if (source.Type == 3 && source.TableFormatting is { } tableFormatting)
            {
                var properties = new StyleTableProperties();
                if (tableFormatting.TableHorizontalBandSize is byte rowBandSize)
                    properties.AppendChild(new TableStyleRowBandSize { Val = rowBandSize });
                if (tableFormatting.TableVerticalBandSize is byte columnBandSize)
                    properties.AppendChild(new TableStyleColumnBandSize { Val = columnBandSize });
                if (tableFormatting.TableJustification is byte tableAlignment)
                    properties.AppendChild(new TableJustification
                    {
                        Val = tableAlignment switch
                        {
                            0 => TableRowAlignmentValues.Left,
                            1 => TableRowAlignmentValues.Center,
                            2 => TableRowAlignmentValues.Right,
                            _ => throw new InvalidDataException("A DOC table style has invalid justification.")
                        }
                    });
                if (tableFormatting.TableIndentTwips is short tableIndent)
                    properties.AppendChild(new TableIndentation
                    {
                        Width = tableIndent,
                        Type = TableWidthUnitValues.Dxa
                    });
                if (tableFormatting.TableCellSpacingTwips is ushort spacing)
                    properties.AppendChild(new TableCellSpacing
                    {
                        Width = spacing.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        Type = TableWidthUnitValues.Dxa
                    });
                if (tableFormatting.TableBorders is { } tableBorders)
                {
                    var borders = new TableBorders();
                    // DOC encodes absent sides as zero-valued slots. In an
                    // unbased style, writing those slots as explicit none
                    // changes Word's conditional border extent.
                    var preserveNone = source.BasedOnIndex != null;
                    if (tableBorders.Top is { } top && (preserveNone || top.Type != 0))
                    { var edge = new TopBorder(); top.ApplyTo(edge); borders.AppendChild(edge); }
                    if (tableBorders.Left is { } left && (preserveNone || left.Type != 0))
                    { var edge = new LeftBorder(); left.ApplyTo(edge); borders.AppendChild(edge); }
                    if (tableBorders.Bottom is { } bottom && (preserveNone || bottom.Type != 0))
                    { var edge = new BottomBorder(); bottom.ApplyTo(edge); borders.AppendChild(edge); }
                    if (tableBorders.Right is { } right && (preserveNone || right.Type != 0))
                    { var edge = new RightBorder(); right.ApplyTo(edge); borders.AppendChild(edge); }
                    if (tableBorders.InsideHorizontal is { } insideH &&
                        (preserveNone || insideH.Type != 0))
                    { var edge = new InsideHorizontalBorder(); insideH.ApplyTo(edge);
                        borders.AppendChild(edge); }
                    if (tableBorders.InsideVertical is { } insideV &&
                        (preserveNone || insideV.Type != 0))
                    { var edge = new InsideVerticalBorder(); insideV.ApplyTo(edge);
                        borders.AppendChild(edge); }
                    if (borders.HasChildren) properties.AppendChild(borders);
                }
                if (tableFormatting.TableBackgroundShading is { } shade &&
                    DocShadingPatterns.ToOpenXml(shade.Pattern) is { } tablePattern)
                {
                    var shading = new Shading { Val = tablePattern };
                    if (ColorValue(shade.FillRgb) is string fill) shading.Fill = fill;
                    if (ColorValue(shade.ForegroundRgb) is string foreground)
                        shading.Color = foreground;
                    properties.AppendChild(shading);
                }
                if (tableFormatting.TableDefaultCellMargins is { } margins)
                    properties.AppendChild(CreateDefaultCellMargins(margins));
                if (properties.HasChildren) style.AppendChild(properties);
            }
            if (source.Type == 3)
            {
                var cellProperties = new StyleTableCellProperties();
                if (source.TableFormatting?.TableStyleShading is { } shade &&
                    DocShadingPatterns.ToOpenXml(shade.Pattern) is { } tablePattern)
                {
                    var shading = new Shading { Val = tablePattern };
                    if (ColorValue(shade.FillRgb) is string fill) shading.Fill = fill;
                    if (ColorValue(shade.ForegroundRgb) is string foreground)
                        shading.Color = foreground;
                    cellProperties.AppendChild(shading);
                }
                if (source.TableFormatting?.TableStyleNoWrap is bool noWrap)
                    cellProperties.AppendChild(new NoWrap
                        { Val = noWrap ? OnOffOnlyValues.On : OnOffOnlyValues.Off });
                if (source.TableFormatting?.TableStyleVerticalAlignment is
                    byte alignment)
                    cellProperties.AppendChild(new TableCellVerticalAlignment
                    {
                        Val = alignment switch
                        {
                            1 => TableVerticalAlignmentValues.Center,
                            2 => TableVerticalAlignmentValues.Bottom,
                            _ => TableVerticalAlignmentValues.Top
                        }
                    });
                if (cellProperties.HasChildren) style.AppendChild(cellProperties);
            }
            if (source.Type == 3)
                foreach (var condition in (source.ConditionalParagraphFormatting?.Keys ??
                    Enumerable.Empty<ushort>()).Concat(
                    source.ConditionalCharacterFormatting?.Keys ?? Enumerable.Empty<ushort>())
                    .Concat(source.ConditionalTableShading?.Keys ??
                        Enumerable.Empty<ushort>())
                    .Concat(source.ConditionalTableBorders?.Keys ??
                        Enumerable.Empty<ushort>())
                    .Concat(source.ConditionalTableVerticalAlignment?.Keys ??
                        Enumerable.Empty<ushort>())
                    .Concat(source.ConditionalTableNoWrap?.Keys ??
                        Enumerable.Empty<ushort>())
                    .Distinct().OrderBy(x => x))
                {
                    var kind = DocTableStyleCondition.ToOpenXml(condition);
                    var conditional = new TableStyleProperties { Type = kind };
                    if (source.ConditionalParagraphFormatting?.TryGetValue(condition,
                        out var conditionalParagraphFormatting) == true)
                    {
                        var properties = new StyleParagraphProperties();
                        AppendParagraphFormatting(properties, conditionalParagraphFormatting);
                        if (properties.HasChildren) conditional.AppendChild(properties);
                    }
                    if (source.ConditionalCharacterFormatting?.TryGetValue(condition,
                        out var characterFormatting) == true)
                    {
                        var properties = new RunPropertiesBaseStyle();
                        AppendCharacterFormatting(properties, characterFormatting, byIndex);
                        if (properties.HasChildren) conditional.AppendChild(properties);
                    }
                    var cellProperties = new StyleTableCellProperties();
                    if (source.ConditionalTableBorders?.TryGetValue(condition,
                        out var conditionalBorders) == true)
                        cellProperties.AppendChild(conditionalBorders.ToOpenXml());
                    if (source.ConditionalTableNoWrap?.TryGetValue(condition,
                        out var noWrap) == true)
                        cellProperties.AppendChild(new NoWrap
                            { Val = noWrap ? OnOffOnlyValues.On : OnOffOnlyValues.Off });
                    if (source.ConditionalTableVerticalAlignment?.TryGetValue(
                        condition, out var verticalAlignment) == true)
                        cellProperties.AppendChild(new TableCellVerticalAlignment
                        {
                            Val = verticalAlignment switch
                            {
                                1 => TableVerticalAlignmentValues.Center,
                                2 => TableVerticalAlignmentValues.Bottom,
                                _ => TableVerticalAlignmentValues.Top
                            }
                        });
                    if (source.ConditionalTableShading?.TryGetValue(condition,
                        out var conditionalShade) == true &&
                        DocShadingPatterns.ToOpenXml(conditionalShade.Pattern) is { } pattern)
                    {
                        var shading = new Shading { Val = pattern };
                        if (ColorValue(conditionalShade.FillRgb) is string fill)
                            shading.Fill = fill;
                        if (ColorValue(conditionalShade.ForegroundRgb) is string foreground)
                            shading.Color = foreground;
                        cellProperties.AppendChild(shading);
                    }
                    if (cellProperties.HasChildren) conditional.AppendChild(cellProperties);
                    if (conditional.HasChildren) style.AppendChild(conditional);
                }
            root.AppendChild(style);
        }
        part.Styles = root;
        part.Styles.Save();
    }

    private static List<Paragraph> RenderParagraphs(DocStoryText? story,
        IReadOnlyList<DocCharacterFormattingRange> formatting,
        IReadOnlyList<DocParagraphStyleRange> paragraphStyles,
        IReadOnlyDictionary<int, DocStyleDefinition> styles,
        IReadOnlyList<DocStoryBookmark> bookmarks,
        IReadOnlyDictionary<int, DocStoryInlinePicture> inlinePictureReferences,
        DocTextIndex index, OpenXmlPart contentPart,
        IDictionary<int, DocInlinePicture> pictures,
        IDictionary<ushort, int> omitted, IDictionary<ushort, int> approximate,
        IReadOnlyDictionary<int, DocStoryFloatingPicture[]>? floatingPictures = null)
    {
        var result = new List<Paragraph>();
        var fieldInstructions = new Stack<bool>();
        var bookmarkEvents = new Dictionary<uint, List<(int Id, DocBookmark Bookmark, bool Start)>>();
        for (var i = 0; i < bookmarks.Count; i++)
        {
            if (story == null) continue;
            var reference = bookmarks[i];
            var id = reference.Id ?? throw new InvalidDataException(
                "An indexed DOC bookmark has no document identifier.");
            var bookmark = new DocBookmark(reference.Name,
                checked(story.CpStart + (uint)reference.Start),
                checked(story.CpStart + (uint)reference.End));
            void Add(uint cp, bool start)
            {
                if (!bookmarkEvents.TryGetValue(cp, out var events))
                    bookmarkEvents[cp] = events = new();
                events.Add((id, bookmark, start));
            }
            Add(bookmark.CpStart, true);
            Add(bookmark.CpEnd, false);
        }
        foreach (var sourceParagraph in story?.Paragraphs ?? [])
        {
            var paragraph = new Paragraph();
            var emittedBookmarks = new HashSet<uint>();
            void EmitBookmarks(uint cp)
            {
                if (!emittedBookmarks.Add(cp) ||
                    !bookmarkEvents.TryGetValue(cp, out var events)) return;
                foreach (var item in events.OrderBy(x =>
                    x.Bookmark.CpStart == x.Bookmark.CpEnd ?
                        (x.Start ? 0 : 1) : (x.Start ? 1 : 0)))
                {
                    var id = item.Id.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    if (item.Start)
                        paragraph.AppendChild(new BookmarkStart
                            { Id = id, Name = item.Bookmark.Name });
                    else
                        paragraph.AppendChild(new BookmarkEnd { Id = id });
                }
            }
            var paragraphRange = sourceParagraph.CpEnd > sourceParagraph.CpStart
                ? FindParagraphRange(paragraphStyles, sourceParagraph.CpEnd - 1) : null;
            var paragraphProperties = new ParagraphProperties();
            if (paragraphRange != null &&
                styles.TryGetValue(paragraphRange.StyleIndex, out var style) && style.Type == 1)
                paragraphProperties.AppendChild(new ParagraphStyleId { Val = style.StyleId });
            if (paragraphRange?.Formatting is { IsEmpty: false } directParagraphFormatting)
                AppendParagraphFormatting(paragraphProperties, directParagraphFormatting);
            if (sourceParagraph.End != DocParagraphEnd.None &&
                FindCharacterRange(formatting, sourceParagraph.CpEnd - 1, out _) is
                    { Formatting.IsEmpty: false } markRange)
            {
                var markProperties = new ParagraphMarkRunProperties();
                if (markRange.Formatting.DeletedRevision == true)
                    markProperties.AppendChild(new Deleted
                    {
                        Id = (sourceParagraph.CpEnd - 1).ToString(System.Globalization.CultureInfo.InvariantCulture),
                        Author = markRange.Formatting.DeletedRevisionAuthor ?? "Unknown",
                        Date = markRange.Formatting.DeletedRevisionAt
                    });
                if (markRange.Formatting.InsertedRevision == true)
                    markProperties.AppendChild(new Inserted
                    {
                        Id = (sourceParagraph.CpEnd - 1).ToString(System.Globalization.CultureInfo.InvariantCulture),
                        Author = markRange.Formatting.InsertedRevisionAuthor ?? "Unknown",
                        Date = markRange.Formatting.InsertedRevisionAt
                    });
                AppendCharacterFormatting(markProperties, markRange.Formatting, styles);
                if (markProperties.HasChildren)
                    paragraphProperties.AppendChild(markProperties);
            }
            if (paragraphProperties.HasChildren) paragraph.AppendChild(paragraphProperties);
            EmitBookmarks(sourceParagraph.CpStart);
            var sourceAtoms = sourceParagraph.Atoms.ToArray();
            var suppressedShapeAtoms = new HashSet<uint>();
            var inlineShapePictures = new Dictionary<uint, DocInlinePicture>();
            for (var i = 0; i + 5 < sourceAtoms.Length; i++)
            {
                if (sourceAtoms[i].Kind != DocStoryAtomKind.FieldBegin ||
                    sourceAtoms[i + 1].Kind != DocStoryAtomKind.Text ||
                    !sourceAtoms[i + 1].Text.TrimStart().StartsWith("SHAPE ",
                        StringComparison.OrdinalIgnoreCase) ||
                    sourceAtoms[i + 2].Kind != DocStoryAtomKind.FieldSeparator ||
                    sourceAtoms[i + 3].Kind != DocStoryAtomKind.FloatingShapeAnchor ||
                    sourceAtoms[i + 4].Kind != DocStoryAtomKind.InlinePicture ||
                    sourceAtoms[i + 5].Kind != DocStoryAtomKind.FieldEnd ||
                    sourceAtoms[i + 3].CpEnd != sourceAtoms[i + 4].CpStart ||
                    floatingPictures?.TryGetValue(
                        checked((int)(sourceAtoms[i + 3].CpStart - story!.CpStart)),
                        out var anchored) != true || anchored.Length != 1)
                    continue;
                if (!inlinePictureReferences.TryGetValue(
                        checked((int)(sourceAtoms[i + 4].CpStart - story!.CpStart)),
                        out var placeholder) || placeholder.DataOffset is not int offset ||
                    DocInlinePictureReader.TryRead(index.Structure, offset) != null)
                    continue;
                var floating = anchored[0];
                var width = floating.WidthEmu;
                var height = floating.HeightEmu;
                if (width <= 0 || height <= 0) continue;
                inlineShapePictures[sourceAtoms[i + 3].CpStart] = new DocInlinePicture(
                    floating.Bytes, floating.ContentType, width, height, floating.Crop,
                    floating.FlipHorizontal, floating.FlipVertical,
                    floating.RotationDegrees);
                foreach (var at in new[] { i, i + 1, i + 2, i + 4, i + 5 })
                    suppressedShapeAtoms.Add(sourceAtoms[at].CpStart);
            }
            foreach (var atom in sourceParagraph.Atoms)
            {
                EmitBookmarks(atom.CpStart);
                if (suppressedShapeAtoms.Contains(atom.CpStart)) continue;
                switch (atom.Kind)
                {
                    case DocStoryAtomKind.Text:
                        {
                            var fieldCode = fieldInstructions.Count != 0 &&
                                fieldInstructions.Peek();
                            var start = atom.CpStart;
                            foreach (var boundary in bookmarkEvents.Keys.Where(x =>
                                x > atom.CpStart && x < atom.CpEnd).OrderBy(x => x))
                            {
                                var length = checked((int)(boundary - start));
                                AppendTextRuns(paragraph, new DocStoryAtom(start, boundary,
                                    DocStoryAtomKind.Text,
                                    atom.Text.Substring(checked((int)(start - atom.CpStart)), length)),
                                    formatting, styles, fieldCode);
                                EmitBookmarks(boundary);
                                start = boundary;
                            }
                            AppendTextRuns(paragraph, new DocStoryAtom(start, atom.CpEnd,
                                DocStoryAtomKind.Text,
                                atom.Text.Substring(checked((int)(start - atom.CpStart)))),
                                formatting, styles, fieldCode);
                        }
                        break;
                    case DocStoryAtomKind.FieldBegin:
                        fieldInstructions.Push(true);
                        AppendControlRun(paragraph, new FieldChar
                            { FieldCharType = FieldCharValues.Begin }, atom.CpStart,
                            formatting, styles);
                        break;
                    case DocStoryAtomKind.FieldSeparator:
                        if (fieldInstructions.Count == 0 || !fieldInstructions.Peek())
                            throw new InvalidDataException("A DOC field separator is unmatched.");
                        fieldInstructions.Pop();
                        fieldInstructions.Push(false);
                        AppendControlRun(paragraph, new FieldChar
                            { FieldCharType = FieldCharValues.Separate }, atom.CpStart,
                            formatting, styles);
                        break;
                    case DocStoryAtomKind.FieldEnd:
                        if (fieldInstructions.Count == 0)
                            throw new InvalidDataException("A DOC field end is unmatched.");
                        fieldInstructions.Pop();
                        AppendControlRun(paragraph, new FieldChar
                            { FieldCharType = FieldCharValues.End }, atom.CpStart,
                            formatting, styles);
                        break;
                    case DocStoryAtomKind.FieldData:
                        // The binary HFD payload accompanies the editable field instruction.
                        break;
                    case DocStoryAtomKind.Tab:
                        AppendControlRun(paragraph, new TabChar(), atom.CpStart,
                            formatting, styles);
                        break;
                    case DocStoryAtomKind.LineBreak:
                        AppendControlRun(paragraph, new Break(), atom.CpStart,
                            formatting, styles);
                        break;
                    case DocStoryAtomKind.NoBreakHyphen:
                        AppendControlRun(paragraph, new NoBreakHyphen(), atom.CpStart,
                            formatting, styles);
                        break;
                    case DocStoryAtomKind.SoftHyphen:
                        AppendControlRun(paragraph, new SoftHyphen(), atom.CpStart,
                            formatting, styles);
                        break;
                    case DocStoryAtomKind.LegacyPageNumberBlock:
                        AppendControlRun(paragraph, new PageNumber(), atom.CpStart,
                            formatting, styles);
                        break;
                    case DocStoryAtomKind.LegacyDateBlock:
                        AppendControlRun(paragraph, atom.Text[0] switch
                        {
                            '\u0010' => new DayShort(),
                            '\u000F' => new DayLong(),
                            '!' => new MonthShort(),
                            '%' => new MonthLong(),
                            '#' => new YearShort(),
                            '"' => new YearLong(),
                            _ => throw new InvalidDataException("Unknown DOC date block.")
                        }, atom.CpStart, formatting, styles);
                        break;
                    case DocStoryAtomKind.CellMark:
                        paragraph.AppendChild(new Run(new TabChar()));
                        Count(approximate, atom.Text[0]);
                        break;
                    case DocStoryAtomKind.PageBreak:
                        AppendControlRun(paragraph, new Break { Type = BreakValues.Page },
                            atom.CpStart, formatting, styles);
                        break;
                    case DocStoryAtomKind.ColumnBreak:
                        AppendControlRun(paragraph, new Break { Type = BreakValues.Column },
                            atom.CpStart, formatting, styles);
                        break;
                    case DocStoryAtomKind.InlinePicture:
                        if (!inlinePictureReferences.TryGetValue(
                                checked((int)(atom.CpStart - story!.CpStart)),
                                out var pictureReference) ||
                            pictureReference.DataOffset is not int pictureOffset)
                            throw new InvalidDataException("A DOC picture marker has no Data-stream offset.");
                        if (!pictures.TryGetValue(pictureOffset, out var picture))
                        {
                            picture = DocInlinePictureReader.TryRead(index.Structure,
                                pictureOffset);
                            if (picture == null)
                                throw new NotSupportedException(
                                    "The DOC inline picture has no supported embedded blip or linked image.");
                            pictures[pictureOffset] = picture;
                        }
                        paragraph.AppendChild(new Run(CreatePictureDrawing(contentPart, picture,
                            atom.CpStart)));
                        break;
                    case DocStoryAtomKind.FloatingShapeAnchor:
                        if (inlineShapePictures.TryGetValue(atom.CpStart,
                            out var inlineShape))
                        {
                            paragraph.AppendChild(new Run(CreatePictureDrawing(
                                contentPart, inlineShape, atom.CpStart)));
                            break;
                        }
                        if (floatingPictures?.TryGetValue(
                                checked((int)(atom.CpStart - story!.CpStart)),
                                out var anchored) == true)
                            foreach (var floating in anchored)
                                paragraph.AppendChild(new Run(CreateFloatingPictureDrawing(
                                    contentPart, floating)));
                        else
                            Count(omitted, atom.Text[0]);
                        break;
                    case DocStoryAtomKind.UnsupportedControl:
                        Count(omitted, atom.Text[0]);
                        break;
                }
            }
            if (sourceParagraph.End != DocParagraphEnd.None)
                EmitBookmarks(sourceParagraph.CpEnd - 1);
            if (story != null && sourceParagraph.CpEnd == story.CpEnd)
                EmitBookmarks(sourceParagraph.CpEnd);
            result.Add(paragraph);
        }
        if (fieldInstructions.Count != 0)
            throw new InvalidDataException("A DOC field has no end.");
        return result;
    }

    private static Drawing CreatePictureDrawing(OpenXmlPart target,
        DocInlinePicture picture, uint cp)
    {
        string relationship;
        var blip = new A.Blip();
        if (picture.LinkedImage is { } linkedImage)
        {
            relationship = target.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                linkedImage).Id;
            blip.Link = relationship;
        }
        else
        {
            var image = target.AddNewPart<ImagePart>(picture.ContentType);
            using (var data = new MemoryStream(picture.Bytes, writable: false))
                image.FeedData(data);
            relationship = target.GetIdOfPart(image);
            blip.Embed = relationship;
        }
        var pictureId = checked(cp + 1);
        var fill = new PIC.BlipFill(blip);
        if (picture.Crop is { IsEmpty: false } crop)
            fill.Append(new A.SourceRectangle
            {
                Left = checked((int)Math.Round(crop.Left * 100000)),
                Top = checked((int)Math.Round(crop.Top * 100000)),
                Right = checked((int)Math.Round(crop.Right * 100000)),
                Bottom = checked((int)Math.Round(crop.Bottom * 100000))
            });
        fill.Append(new A.Stretch(new A.FillRectangle()));
        var shape = new PIC.Picture(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties { Id = pictureId,
                    Name = $"Picture {pictureId}" },
                new PIC.NonVisualPictureDrawingProperties()),
            fill,
            new PIC.ShapeProperties(
                new A.Transform2D(new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = picture.WidthEmu, Cy = picture.HeightEmu })
                {
                    HorizontalFlip = picture.FlipHorizontal,
                    VerticalFlip = picture.FlipVertical,
                    Rotation = checked((int)Math.Round(picture.RotationDegrees * 60000))
                },
                new A.PresetGeometry(new A.AdjustValueList())
                    { Preset = A.ShapeTypeValues.Rectangle }));
        return new Drawing(new WP.Inline(
            new WP.Extent { Cx = picture.WidthEmu, Cy = picture.HeightEmu },
            new WP.DocProperties { Id = pictureId, Name = $"Picture {pictureId}" },
            new WP.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(new A.GraphicData(shape)
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        {
            DistanceFromTop = 0U, DistanceFromBottom = 0U,
            DistanceFromLeft = 0U, DistanceFromRight = 0U
        });
    }

    private static Drawing CreateFloatingPictureDrawing(OpenXmlPart target,
        DocStoryFloatingPicture picture)
    {
        var shapeId = picture.ShapeId ?? throw new InvalidDataException(
            "An indexed DOC floating picture has no shape identifier.");
        var image = target.AddNewPart<ImagePart>(picture.ContentType);
        using (var data = new MemoryStream(picture.Bytes, writable: false))
            image.FeedData(data);
        var relationship = target.GetIdOfPart(image);
        var width = picture.WidthEmu;
        var height = picture.HeightEmu;
        if (width <= 0 || height <= 0)
            throw new InvalidDataException("A DOC floating picture has invalid bounds.");
        var fill = new PIC.BlipFill(new A.Blip { Embed = relationship });
        if (picture.Crop is { IsEmpty: false } crop)
            fill.Append(new A.SourceRectangle
            {
                Left = checked((int)Math.Round(crop.Left * 100000)),
                Top = checked((int)Math.Round(crop.Top * 100000)),
                Right = checked((int)Math.Round(crop.Right * 100000)),
                Bottom = checked((int)Math.Round(crop.Bottom * 100000))
            });
        fill.Append(new A.Stretch(new A.FillRectangle()));
        var shape = new PIC.Picture(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties { Id = shapeId,
                    Name = $"Picture {shapeId}" },
                new PIC.NonVisualPictureDrawingProperties()),
            fill,
            new PIC.ShapeProperties(
                new A.Transform2D(new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = width, Cy = height })
                {
                    HorizontalFlip = picture.FlipHorizontal,
                    VerticalFlip = picture.FlipVertical,
                    Rotation = checked((int)Math.Round(picture.RotationDegrees * 60000))
                },
                new A.PresetGeometry(new A.AdjustValueList())
                    { Preset = A.ShapeTypeValues.Rectangle }));
        OpenXmlElement wrap = picture.WrapCode switch
        {
            0 or 2 => new WP.WrapSquare
            {
                WrapText = picture.WrapSide switch
                {
                    1 => WP.WrapTextValues.Left,
                    2 => WP.WrapTextValues.Right,
                    3 => WP.WrapTextValues.Largest,
                    0 => WP.WrapTextValues.BothSides,
                    _ => throw new NotSupportedException(
                        $"DOC floating picture wrap side {picture.WrapSide} is unsupported.")
                }
            },
            1 => new WP.WrapTopBottom(),
            3 => new WP.WrapNone(),
            4 => new WP.WrapTight(RectangularWrapPolygon())
            {
                WrapText = WrapText(picture.WrapSide)
            },
            5 => new WP.WrapThrough(RectangularWrapPolygon())
            {
                WrapText = WrapText(picture.WrapSide)
            },
            _ => throw new NotSupportedException(
                $"DOC floating picture wrap code {picture.WrapCode} is unsupported.")
        };
        static string Alignment(byte code, bool vertical) => code switch
        {
            1 => vertical ? "top" : "left",
            2 => "center",
            3 => vertical ? "bottom" : "right",
            4 => "inside",
            5 => "outside",
            _ => throw new NotSupportedException("DOC floating picture alignment is unsupported.")
        };
        return new Drawing(new WP.Anchor(
            new WP.SimplePosition { X = 0L, Y = 0L },
            new WP.HorizontalPosition(picture.HorizontalAlignment == 0
                ? (OpenXmlElement)new WP.PositionOffset(checked(picture.LeftTwips * 635L)
                    .ToString(System.Globalization.CultureInfo.InvariantCulture))
                : new WP.HorizontalAlignment { Text = Alignment(picture.HorizontalAlignment, false) })
                { RelativeFrom = picture.HorizontalOrigin switch
                    {
                        0 => WP.HorizontalRelativePositionValues.Margin,
                        1 => WP.HorizontalRelativePositionValues.Page,
                        2 => WP.HorizontalRelativePositionValues.Column,
                        _ => throw new NotSupportedException(
                            $"DOC floating picture horizontal origin {picture.HorizontalOrigin} is unsupported.")
                    } },
            new WP.VerticalPosition(picture.VerticalAlignment == 0
                ? (OpenXmlElement)new WP.PositionOffset(checked(picture.TopTwips * 635L)
                    .ToString(System.Globalization.CultureInfo.InvariantCulture))
                : new WP.VerticalAlignment { Text = Alignment(picture.VerticalAlignment, true) })
                { RelativeFrom = picture.VerticalOrigin switch
                    {
                        0 => WP.VerticalRelativePositionValues.Margin,
                        1 => WP.VerticalRelativePositionValues.Page,
                        2 => WP.VerticalRelativePositionValues.Paragraph,
                        _ => throw new NotSupportedException(
                            $"DOC floating picture vertical origin {picture.VerticalOrigin} is unsupported.")
                    } },
            new WP.Extent { Cx = width, Cy = height },
            new WP.EffectExtent { LeftEdge = 0L, TopEdge = 0L,
                RightEdge = 0L, BottomEdge = 0L },
            wrap,
            new WP.DocProperties { Id = shapeId, Name = $"Picture {shapeId}" },
            new WP.NonVisualGraphicFrameDrawingProperties(),
            new A.Graphic(new A.GraphicData(shape)
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        {
            DistanceFromTop = checked((uint)picture.DistanceTopEmu),
            DistanceFromBottom = checked((uint)picture.DistanceBottomEmu),
            DistanceFromLeft = checked((uint)picture.DistanceLeftEmu),
            DistanceFromRight = checked((uint)picture.DistanceRightEmu),
            SimplePos = false, RelativeHeight = 251658240U,
            BehindDoc = picture.WrapCode == 3 && picture.BehindText,
            Locked = false, LayoutInCell = true,
            AllowOverlap = true
        });
    }

    private static WP.WrapTextValues WrapText(byte side) => side switch
    {
        1 => WP.WrapTextValues.Left,
        2 => WP.WrapTextValues.Right,
        3 => WP.WrapTextValues.Largest,
        0 => WP.WrapTextValues.BothSides,
        _ => throw new NotSupportedException(
            $"DOC floating picture wrap side {side} is unsupported.")
    };

    private static WP.WrapPolygon RectangularWrapPolygon() => new(
        new WP.StartPoint { X = 0L, Y = 0L },
        new WP.LineTo { X = 0L, Y = 21600L },
        new WP.LineTo { X = 21600L, Y = 21600L },
        new WP.LineTo { X = 21600L, Y = 0L },
        new WP.LineTo { X = 0L, Y = 0L }) { Edited = false };

    private static List<OpenXmlElement> RenderBlocks(DocStoryText? story,
        IReadOnlyList<Paragraph> paragraphs,
        IReadOnlyList<DocParagraphStyleRange> paragraphStyles,
        IReadOnlyList<DocStoryTableRow> tableRows,
        IReadOnlyDictionary<int, DocStyleDefinition> styles, bool growAutofit)
    {
        var result = new List<OpenXmlElement>();
        if (story == null) return result;
        var sources = story.Paragraphs;
        DocParagraphFormatting? RowFormattingAt(uint endCp) =>
            tableRows.FirstOrDefault(x => x.SourceCpEnd == endCp)?.Formatting ??
            FindParagraphRange(paragraphStyles, endCp - 1)?.Formatting;
        if (sources.Count != paragraphs.Count)
            throw new InvalidDataException("A DOC story has mismatched paragraph counts.");
        for (var position = 0; position < sources.Count;)
        {
            var source = sources[position];
            var style = source.CpEnd > source.CpStart
                ? FindParagraphRange(paragraphStyles, source.CpEnd - 1) : null;
            if (style?.Formatting?.InTable != true)
            {
                result.Add(paragraphs[position++]);
                continue;
            }
            var tableDepth = style.Formatting.TableDepth ?? 1;

            var table = new Table();
            var rowEdges = new List<(TableRow Row, short[] Edges,
                int[] CellBoundaryIndexes, bool HasBeforeGap)>();
            var rowCellSpacings = new List<ushort?>();
            var rowDefaultMargins = new List<DocCellMargins?>();
            var tableAutoFit = false;
            var tableRightToLeft = false;
            ushort? tableStyleIndex = null;
            ushort? tableLookMask = null;
            uint? tableGroupId = null;
            ushort? tableCellSpacingTwips = null;
            DocTablePreferredWidth? tablePreferredWidth = null;
            short? tableIndentTwips = null;
            byte? tableJustification = null;
            short? firstRowOriginTwips = null;
            DocTableBorders? firstRowBorders = null;
            DocCellShading? tableBackgroundShading = null;
            DocCellMargins? tableDefaultMargins = null;
            while (position < sources.Count)
            {
                source = sources[position];
                style = source.CpEnd > source.CpStart
                    ? FindParagraphRange(paragraphStyles, source.CpEnd - 1) : null;
                if (style?.Formatting?.InTable != true) break;
                if (rowEdges.Count != 0)
                {
                    var rowEndPosition = position;
                    while (rowEndPosition < sources.Count &&
                        sources[rowEndPosition].End != DocParagraphEnd.RowMark)
                        rowEndPosition++;
                    if (rowEndPosition < sources.Count)
                    {
                        var nextRow = RowFormattingAt(sources[rowEndPosition].CpEnd);
                        if ((nextRow?.TableRightToLeft == true) != tableRightToLeft ||
                            (nextRow?.TableStyleIndex ?? 11) != (tableStyleIndex ?? 11) ||
                            nextRow?.TableGroupId != tableGroupId)
                            break;
                    }
                }
                var row = new TableRow();
                var cellBlocks = new List<OpenXmlElement>();
                for (; position < sources.Count; position++)
                {
                    source = sources[position];
                    var currentDepth = source.CpEnd > source.CpStart
                        ? FindParagraphRange(paragraphStyles, source.CpEnd - 1)?
                            .Formatting?.TableDepth ?? 1 : 1;
                    if (source.End == DocParagraphEnd.RowMark &&
                        currentDepth == tableDepth) break;
                    if (currentDepth > tableDepth)
                    {
                        var nestedStart = position;
                        do
                        {
                            position++;
                            if (position >= sources.Count) break;
                            source = sources[position];
                            currentDepth = source.CpEnd > source.CpStart
                                ? FindParagraphRange(paragraphStyles, source.CpEnd - 1)?
                                    .Formatting?.TableDepth ?? 1 : 1;
                        } while (currentDepth > tableDepth);
                        var nestedSources = sources.Skip(nestedStart)
                            .Take(position - nestedStart).ToArray();
                        cellBlocks.AddRange(RenderBlocks(new DocStoryText(story.Name,
                                nestedSources[0].CpStart,
                                nestedSources[nestedSources.Length - 1].CpEnd,
                                nestedSources),
                            paragraphs.Skip(nestedStart).Take(position - nestedStart)
                                .ToArray(), paragraphStyles, tableRows, styles,
                            growAutofit));
                        position--;
                        continue;
                    }
                    cellBlocks.Add(paragraphs[position]);
                    if (source.End != DocParagraphEnd.CellMark) continue;
                    row.AppendChild(new TableCell(cellBlocks));
                    cellBlocks.Clear();
                }
                if (position >= sources.Count || cellBlocks.Count != 0)
                    throw new InvalidDataException("A DOC table row has no terminating row mark.");
                var rowMark = sources[position];
                var rowFormatting = RowFormattingAt(rowMark.CpEnd);
                rowCellSpacings.Add(rowFormatting?.TableCellSpacingTwips);
                rowDefaultMargins.Add(rowFormatting?.TableDefaultCellMargins);
                if (rowEdges.Count == 0)
                {
                    tableAutoFit = rowFormatting?.TableAutoFit == true;
                    tableRightToLeft = rowFormatting?.TableRightToLeft == true;
                    tableStyleIndex = rowFormatting?.TableStyleIndex;
                    tableLookMask = rowFormatting?.TableLookMask;
                    tableGroupId = rowFormatting?.TableGroupId;
                    tableCellSpacingTwips = rowFormatting?.TableCellSpacingTwips;
                    tablePreferredWidth = rowFormatting?.TablePreferredWidth;
                    tableBackgroundShading = rowFormatting?.TableBackgroundShading;
                    tableIndentTwips = rowFormatting?.TableIndentTwips;
                    tableJustification = rowFormatting?.TableJustification;
                    tableDefaultMargins = rowFormatting?.TableDefaultCellMargins;
                    firstRowOriginTwips = rowFormatting?.TableRowOriginTwips;
                    firstRowBorders = rowFormatting?.TableBorders;
                }
                var nextRowPosition = position + 1;
                var nextRowFormatting = nextRowPosition < sources.Count
                    ? FindParagraphRange(paragraphStyles,
                        sources[nextRowPosition].CpEnd - 1)?.Formatting : null;
                var isLastRow = nextRowFormatting?.InTable != true;
                if (!isLastRow)
                {
                    var nextRowEnd = nextRowPosition;
                    while (nextRowEnd < sources.Count &&
                        sources[nextRowEnd].End != DocParagraphEnd.RowMark)
                        nextRowEnd++;
                    if (nextRowEnd < sources.Count)
                    {
                        var nextRow = FindParagraphRange(paragraphStyles,
                            sources[nextRowEnd].CpEnd - 1)?.Formatting;
                        isLastRow = (nextRow?.TableRightToLeft == true) != tableRightToLeft ||
                            (nextRow?.TableStyleIndex ?? 11) != (tableStyleIndex ?? 11) ||
                            nextRow?.TableGroupId != tableGroupId;
                    }
                }
                var edges = rowFormatting?.TableCellEdges;
                if (edges == null || edges.Count != row.Elements<TableCell>().Count() + 1)
                    throw new InvalidDataException("A DOC table row has inconsistent cell edges.");
                var originGap = rowFormatting?.TableWidthBefore is { Unit: 3, Value: 0 } &&
                    firstRowOriginTwips is short firstOrigin &&
                    rowFormatting.TableRowOriginTwips is short rowOrigin &&
                    rowOrigin > firstOrigin && rowEdges.Count > 0 &&
                    edges.Count < rowEdges[0].Edges.Length
                    ? rowOrigin - firstOrigin : 0;
                var omittedBeforeWidth = rowFormatting?.TableWidthBefore is
                    { Unit: 3, Value: 0 } && (originGap > 0 || edges[0] > 0);
                var rowProperties = new TableRowProperties();
                if (rowFormatting?.TableWidthBefore is { } widthBefore &&
                    !omittedBeforeWidth)
                    rowProperties.AppendChild(new WidthBeforeTableRow
                    {
                        Type = widthBefore.Unit == 3 ? TableWidthUnitValues.Dxa :
                            widthBefore.Unit == 2 ? TableWidthUnitValues.Pct :
                                TableWidthUnitValues.Auto,
                        Width = widthBefore.Value.ToString(
                            System.Globalization.CultureInfo.InvariantCulture)
                    });
                if (rowFormatting?.TableWidthAfter is { } widthAfter)
                    rowProperties.AppendChild(new WidthAfterTableRow
                    {
                        Type = widthAfter.Unit == 3 ? TableWidthUnitValues.Dxa :
                            widthAfter.Unit == 2 ? TableWidthUnitValues.Pct :
                                TableWidthUnitValues.Auto,
                        Width = widthAfter.Value.ToString(
                            System.Globalization.CultureInfo.InvariantCulture)
                    });
                if (rowFormatting?.TableCantSplit is bool cannotSplit)
                    rowProperties.AppendChild(new CantSplit { Val = cannotSplit ?
                        OnOffOnlyValues.On : OnOffOnlyValues.Off });
                if (rowFormatting?.TableRowHeightTwips is short rowHeight)
                    rowProperties.AppendChild(new TableRowHeight
                    {
                        Val = (uint)Math.Abs((int)rowHeight),
                        HeightType = rowHeight < 0 ? HeightRuleValues.Exact :
                            HeightRuleValues.AtLeast
                    });
                if (rowFormatting?.TableHeader is bool isHeader)
                    rowProperties.AppendChild(new TableHeader { Val = isHeader ?
                        OnOffOnlyValues.On : OnOffOnlyValues.Off });
                if (rowProperties.HasChildren || omittedBeforeWidth)
                    row.PrependChild(rowProperties);
                var cells = row.Elements<TableCell>().ToArray();
                var rowIndex = table.Elements<TableRow>().Count();
                for (var cellIndex = 0; cellIndex < cells.Length; cellIndex++)
                {
                    var width = edges[cellIndex + 1] - edges[cellIndex];
                    var preferredCellWidth = rowFormatting?.TableCellPreferredWidths is { } widths &&
                        cellIndex < widths.Count ? widths[cellIndex] : null;
                    var cellProperties = new TableCellProperties();
                    // Word uses the cell's conditional-style context when a
                    // direct left alignment overrides a table style. Preserve
                    // that context for a further DOC hop. The val-only form
                    // validates without Word's redundant boolean attributes.
                    if (tableStyleIndex != null && tableLookMask is ushort look &&
                        cells[cellIndex].Descendants<Paragraph>().Any(p =>
                            p.ParagraphProperties?.Justification?.Val?.Value ==
                                JustificationValues.Left))
                    {
                        var bits = "000000000000".ToCharArray();
                        var firstColumn = cellIndex == 0 && (look & 0x0080) != 0;
                        var lastColumn = cellIndex == cells.Length - 1 &&
                            (look & 0x0100) != 0;
                        var firstRow = rowIndex == 0 && (look & 0x0020) != 0;
                        var lastRow = isLastRow && (look & 0x0040) != 0;
                        if (firstColumn) bits[2] = '1';
                        if (lastColumn) bits[3] = '1';
                        if (firstRow && lastColumn) bits[8] = '1';
                        if (firstRow && firstColumn) bits[9] = '1';
                        if (lastRow && lastColumn) bits[10] = '1';
                        if (lastRow && firstColumn) bits[11] = '1';
                        if (bits.Contains('1'))
                            cellProperties.AppendChild(new ConditionalFormatStyle
                            { Val = new string(bits) });
                    }
                    // A styled fixed table can rely on tblGrid without inventing a
                    // preferred cell width. Word otherwise shifts the visible edge.
                    if (preferredCellWidth != null || (!tableAutoFit && tableStyleIndex == null))
                        cellProperties.AppendChild(new TableCellWidth
                        {
                            Type = preferredCellWidth?.Unit switch
                            {
                                1 => TableWidthUnitValues.Auto,
                                2 => TableWidthUnitValues.Pct,
                                _ => TableWidthUnitValues.Dxa
                            },
                            Width = (preferredCellWidth?.Value ?? (ushort)width).ToString(
                                System.Globalization.CultureInfo.InvariantCulture)
                        });
                    if (rowFormatting?.TableCellVerticalMerges is { } merges &&
                        cellIndex < merges.Count && merges[cellIndex] is byte merge)
                    {
                        cellProperties.AppendChild(new VerticalMerge
                        {
                            Val = merge switch
                            {
                                1 => MergedCellValues.Continue,
                                3 => MergedCellValues.Restart,
                                _ => throw new InvalidDataException("A DOC table cell has invalid vertical merge flags.")
                            }
                        });
                    }
                    if (rowFormatting?.TableCellShadings is { } shadings &&
                        cellIndex < shadings.Count && shadings[cellIndex] is { } cellShading &&
                        (cellShading.Pattern == ushort.MaxValue
                            ? ShadingPatternValues.Clear
                            : DocShadingPatterns.ToOpenXml(cellShading.Pattern)) is { } pattern)
                    {
                        static string Color(uint value) =>
                            (((value & 0xFF) << 16) | (value & 0x00FF00) | (value >> 16))
                            .ToString("X6", System.Globalization.CultureInfo.InvariantCulture);
                        var shading = new Shading { Val = pattern };
                        if (cellShading.FillRgb is uint fill) shading.Fill = Color(fill);
                        if (cellShading.ForegroundRgb is uint foreground)
                            shading.Color = Color(foreground);
                        cellProperties.AppendChild(shading);
                    }
                    if (rowFormatting?.TableCellMargins is { } margins &&
                        cellIndex < margins.Count && margins[cellIndex] is { } cellMargins)
                        cellProperties.AppendChild(CreateCellMargins(cellMargins));
                    if (rowFormatting?.TableCellVerticalAlignments is { } alignments &&
                        cellIndex < alignments.Count && alignments[cellIndex] is byte alignment)
                        cellProperties.AppendChild(new TableCellVerticalAlignment
                        {
                            Val = alignment switch
                            {
                                1 => TableVerticalAlignmentValues.Center,
                                2 => TableVerticalAlignmentValues.Bottom,
                                _ => TableVerticalAlignmentValues.Top
                            }
                        });
                    if (rowFormatting?.TableCellTextFlows is { } textFlows &&
                        cellIndex < textFlows.Count && textFlows[cellIndex] is ushort textFlow)
                        cellProperties.AppendChild(new TextDirection
                        {
                            Val = textFlow switch
                            {
                                0 => TextDirectionValues.LefToRightTopToBottom,
                                1 => TextDirectionValues.TopToBottomRightToLeft,
                                3 => TextDirectionValues.BottomToTopLeftToRight,
                                4 => TextDirectionValues.LefttoRightTopToBottomRotated,
                                5 => TextDirectionValues.TopToBottomRightToLeftRotated,
                                _ => throw new InvalidDataException("A DOC table cell text flow is invalid.")
                            }
                        });
                    if (rowFormatting?.TableCellHideMarks is { } hideMarks &&
                        cellIndex < hideMarks.Count && hideMarks[cellIndex] is bool hideMark)
                        cellProperties.AppendChild(new HideMark
                        {
                            Val = hideMark ? OnOffOnlyValues.On : OnOffOnlyValues.Off
                        });
                    if (rowFormatting?.TableCellNoWraps is { } noWraps &&
                        cellIndex < noWraps.Count && noWraps[cellIndex] is bool noWrap)
                        cellProperties.AppendChild(new NoWrap
                        {
                            Val = noWrap ? OnOffOnlyValues.On : OnOffOnlyValues.Off
                        });
                    if (rowFormatting?.TableCellFitTexts is { } fitTexts &&
                        cellIndex < fitTexts.Count && fitTexts[cellIndex] is bool fitText)
                        cellProperties.AppendChild(new TableCellFitText
                        {
                            Val = fitText ? OnOffOnlyValues.On : OnOffOnlyValues.Off
                        });
                    var explicitBorders = rowFormatting?.TableCellBorders is { } borders &&
                        cellIndex < borders.Count ? borders[cellIndex] : null;
                    var tableBorders = rowFormatting?.TableBorders;
                    var cellBorders = new DocCellBorders(
                        explicitBorders?.Top ?? (rowIndex == 0 ? tableBorders?.Top :
                            tableBorders?.InsideHorizontal),
                        explicitBorders?.Left ?? (cellIndex == 0 ? tableBorders?.Left :
                            tableBorders?.InsideVertical),
                        explicitBorders?.Bottom ?? (isLastRow ? tableBorders?.Bottom :
                            tableBorders?.InsideHorizontal),
                        explicitBorders?.Right ?? (cellIndex == cells.Length - 1 ?
                            tableBorders?.Right : tableBorders?.InsideVertical),
                        explicitBorders?.TopLeftToBottomRight,
                        explicitBorders?.TopRightToBottomLeft);
                    if (cellBorders.Top != null || cellBorders.Left != null ||
                        cellBorders.Bottom != null || cellBorders.Right != null ||
                        cellBorders.TopLeftToBottomRight != null ||
                        cellBorders.TopRightToBottomLeft != null)
                    {
                        var openXmlBorders = new TableCellBorders();
                        if (cellBorders.Top is { } top)
                        { var edge = new TopBorder(); top.ApplyTo(edge); if (top.Type == 0) edge.Val = BorderValues.Nil; openXmlBorders.AppendChild(edge); }
                        if (cellBorders.Left is { } left)
                        { var edge = new LeftBorder(); left.ApplyTo(edge); if (left.Type == 0) edge.Val = BorderValues.Nil; openXmlBorders.AppendChild(edge); }
                        if (cellBorders.Bottom is { } bottom)
                        { var edge = new BottomBorder(); bottom.ApplyTo(edge); if (bottom.Type == 0) edge.Val = BorderValues.Nil; openXmlBorders.AppendChild(edge); }
                        if (cellBorders.Right is { } right)
                        { var edge = new RightBorder(); right.ApplyTo(edge); if (right.Type == 0) edge.Val = BorderValues.Nil; openXmlBorders.AppendChild(edge); }
                        if (cellBorders.TopLeftToBottomRight is { } down)
                        { var edge = new TopLeftToBottomRightCellBorder(); down.ApplyTo(edge); if (down.Type == 0) edge.Val = BorderValues.Nil; openXmlBorders.AppendChild(edge); }
                        if (cellBorders.TopRightToBottomLeft is { } up)
                        { var edge = new TopRightToBottomLeftCellBorder(); up.ApplyTo(edge); if (up.Type == 0) edge.Val = BorderValues.Nil; openXmlBorders.AppendChild(edge); }
                        if (openXmlBorders.HasChildren)
                        {
                            var following = (OpenXmlElement?)cellProperties
                                .GetFirstChild<Shading>() ??
                                (OpenXmlElement?)cellProperties.GetFirstChild<TableCellMargin>() ??
                                cellProperties.GetFirstChild<TableCellVerticalAlignment>();
                            if (following == null) cellProperties.AppendChild(openXmlBorders);
                            else cellProperties.InsertBefore(openXmlBorders, following);
                        }
                    }
                    cells[cellIndex].PrependChild(cellProperties);
                }
                if (rowFormatting?.TableCellHorizontalMerges is { } horizontalMerges)
                {
                    if (horizontalMerges.Count != cells.Length)
                        throw new InvalidDataException("A DOC table row has inconsistent horizontal merges.");
                    var mergeStates = horizontalMerges.ToArray();
                    // Word's DOC save can place the continuation flags before
                    // the terminal restart flag. Normalize that physical run
                    // while leaving every cell's content in its original cell.
                    for (var i = 0; i < mergeStates.Length; i++)
                    {
                        if (mergeStates[i] != 1 ||
                            (i > 0 && mergeStates[i - 1] != null)) continue;
                        var end = i + 1;
                        while (end < mergeStates.Length && mergeStates[end] == 1) end++;
                        if (end >= mergeStates.Length || mergeStates[end] is not (2 or 3))
                            continue;
                        mergeStates[i] = 2;
                        for (var merged = i + 1; merged <= end; merged++)
                            mergeStates[merged] = 1;
                        i = end;
                    }
                    for (var cellIndex = 0; cellIndex < cells.Length; cellIndex++)
                    {
                        var state = mergeStates[cellIndex];
                        if (state == null) continue;
                        if (state == 1)
                            throw new InvalidDataException("A DOC table has an orphaned horizontal merge continuation.");
                        if (state is not (2 or 3))
                            throw new InvalidDataException("A DOC table has invalid horizontal merge flags.");
                        var end = cellIndex + 1;
                        while (end < cells.Length && mergeStates[end] == 1) end++;
                        if (end == cellIndex + 1)
                            throw new InvalidDataException("A DOC horizontal merge has no continuation cell.");
                        if (Enumerable.Range(cellIndex + 1, end - cellIndex - 1)
                            .Any(index => !string.IsNullOrEmpty(cells[index].InnerText)))
                        {
                            for (var merged = cellIndex; merged < end; merged++)
                            {
                                var properties = cells[merged].TableCellProperties!;
                                var marker = new HorizontalMerge
                                {
                                    Val = merged == cellIndex
                                        ? MergedCellValues.Restart : MergedCellValues.Continue
                                };
                                var after = (OpenXmlElement?)properties.GetFirstChild<GridSpan>() ??
                                    properties.GetFirstChild<TableCellWidth>();
                                if (after == null) properties.PrependChild(marker);
                                else properties.InsertAfter(marker, after);
                            }
                            cellIndex = end - 1;
                            continue;
                        }
                        var master = cells[cellIndex];
                        var masterProperties = master.TableCellProperties!;
                        var masterWidth = masterProperties.GetFirstChild<TableCellWidth>();
                        if (masterWidth != null &&
                            (rowFormatting?.TableCellPreferredWidths is not { } preferredWidths ||
                             cellIndex >= preferredWidths.Count ||
                             preferredWidths[cellIndex] == null))
                            masterWidth.Width = (edges[end] - edges[cellIndex]).ToString(
                                System.Globalization.CultureInfo.InvariantCulture);
                        for (var continuation = cellIndex + 1; continuation < end; continuation++)
                        {
                            if (!string.IsNullOrEmpty(cells[continuation].InnerText))
                                throw new NotSupportedException("A DOC merged continuation cell contains text.");
                            row.RemoveChild(cells[continuation]);
                        }
                        cellIndex = end - 1;
                    }
                }
                var positionedEdges = edges.ToArray();
                var leadingOffset = originGap > 0 ? originGap :
                    rowFormatting?.TableWidthBefore is { Unit: 3 } leading
                        ? leading.Value : 0;
                if (leadingOffset > 0)
                    for (var edgeIndex = 0; edgeIndex < positionedEdges.Length; edgeIndex++)
                        positionedEdges[edgeIndex] = checked((short)(
                            positionedEdges[edgeIndex] + leadingOffset));
                rowEdges.Add((row, positionedEdges,
                    Enumerable.Range(0, cells.Length)
                        .Where(i => cells[i].Parent == row)
                        .Append(cells.Length).ToArray(),
                    rowFormatting?.TableWidthBefore != null || originGap > 0));
                table.AppendChild(row);
                position++;
            }
            if (rowEdges.Count == 0) throw new InvalidDataException("A DOC table has no rows.");
            if (rowDefaultMargins.Any(x => x != rowDefaultMargins[0]))
            {
                tableDefaultMargins = null;
                for (var i = 0; i < rowEdges.Count; i++)
                    if (rowDefaultMargins[i] is { } margins)
                        rowEdges[i].Row.PrependChild(new TablePropertyExceptions(
                            CreateDefaultCellMargins(margins)));
            }
            var uniformCellSpacing = rowCellSpacings.All(x => x == rowCellSpacings[0]);
            if (!uniformCellSpacing)
            {
                for (var i = 0; i < rowEdges.Count; i++)
                {
                    if (rowCellSpacings[i] is not ushort rowSpacing) continue;
                    var rowProperties = rowEdges[i].Row.GetFirstChild<TableRowProperties>();
                    if (rowProperties == null)
                    {
                        rowProperties = new TableRowProperties();
                        rowEdges[i].Row.PrependChild(rowProperties);
                    }
                    rowProperties.AppendChild(new TableCellSpacing
                    {
                        Width = rowSpacing.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        Type = TableWidthUnitValues.Dxa
                    });
                }
            }
            // Word can store a one-twip internal edge in a horizontally
            // merged row. It marks the physical cells, not a useful grid
            // column. Use an unmerged row's shared outer edges to recover
            // the table grid when one is available.
            for (var i = 0; i < rowEdges.Count; i++)
            {
                var merged = rowEdges[i];
                if (!merged.Row.Descendants<HorizontalMerge>().Any() ||
                    !Enumerable.Range(1, merged.Edges.Length - 2).Any(edge =>
                        merged.Edges[edge] - merged.Edges[edge - 1] == 1))
                    continue;
                var reference = rowEdges.FirstOrDefault(other =>
                    other.Row != merged.Row &&
                    !other.Row.Descendants<HorizontalMerge>().Any() &&
                    other.Edges.Length == merged.Edges.Length &&
                    other.Edges[0] == merged.Edges[0] &&
                    other.Edges[other.Edges.Length - 1] ==
                    merged.Edges[merged.Edges.Length - 1]);
                if (reference.Edges == null) continue;
                rowEdges[i] = (merged.Row, reference.Edges.ToArray(),
                    merged.CellBoundaryIndexes, merged.HasBeforeGap);
            }
            var gridEdgeSet = new HashSet<int>(rowEdges.SelectMany(x => x.Edges)
                .Select(x => (int)x));
            // A common leading/trailing row gap has no cell edge from another
            // row to supply its grid column. Add one so gridBefore/gridAfter
            // can carry the preferred width through another DOCX walk.
            foreach (var (row, edges, _, hasBeforeGap) in rowEdges)
            {
                var properties = row.GetFirstChild<TableRowProperties>();
                if (hasBeforeGap && properties?.GetFirstChild<WidthBeforeTableRow>()
                    is { } before && before.Type?.Value != TableWidthUnitValues.Dxa &&
                    !gridEdgeSet.Any(x => x < edges[0]))
                {
                    gridEdgeSet.Add(edges[0] - 1);
                }
                if (properties?.GetFirstChild<WidthAfterTableRow>() is { } after &&
                    after.Type?.Value != TableWidthUnitValues.Dxa &&
                    !gridEdgeSet.Any(x => x > edges[edges.Length - 1]))
                {
                    gridEdgeSet.Add(edges[edges.Length - 1] + 1);
                }
            }
            var gridEdges = gridEdgeSet.OrderBy(x => x).ToArray();
            foreach (var (row, edges, cellBoundaryIndexes, hasBeforeGap) in rowEdges)
            {
                var rowProperties = row.GetFirstChild<TableRowProperties>();
                if (hasBeforeGap && rowProperties != null)
                {
                    var beforeCount = Array.BinarySearch(gridEdges, edges[0]);
                    if (beforeCount > 0)
                        rowProperties.PrependChild(new GridBefore { Val = beforeCount });
                }
                if (rowProperties?.GetFirstChild<WidthAfterTableRow>() != null)
                {
                    var afterCount = gridEdges.Length - 1 -
                        Array.BinarySearch(gridEdges, edges[edges.Length - 1]);
                    if (afterCount > 0)
                    {
                        rowProperties.PrependChild(new GridAfter { Val = afterCount });
                        var widthAfter = rowProperties.GetFirstChild<WidthAfterTableRow>();
                        if (widthAfter?.Type?.Value == TableWidthUnitValues.Dxa &&
                            widthAfter.Width?.Value == "0")
                            widthAfter.Remove();
                    }
                }
                var cells = row.Elements<TableCell>().ToArray();
                for (var cellIndex = 0; cellIndex < cells.Length; cellIndex++)
                {
                    var first = Array.BinarySearch(gridEdges,
                        edges[cellBoundaryIndexes[cellIndex]]);
                    var last = Array.BinarySearch(gridEdges,
                        edges[cellBoundaryIndexes[cellIndex + 1]]);
                    if (first < 0 || last <= first)
                        throw new InvalidDataException("A DOC table row has inconsistent grid edges.");
                    if (last - first > 1)
                    {
                        var properties = cells[cellIndex].TableCellProperties!;
                        var cellWidth = properties.GetFirstChild<TableCellWidth>();
                        if (cellWidth == null)
                            properties.PrependChild(new GridSpan { Val = last - first });
                        else properties.InsertAfter(new GridSpan { Val = last - first },
                            cellWidth);
                    }
                }
            }
            var grid = new TableGrid();
            for (var i = 0; i < gridEdges.Length - 1; i++)
                grid.AppendChild(new GridColumn
                {
                    Width = (gridEdges[i + 1] - gridEdges[i]).ToString(
                        System.Globalization.CultureInfo.InvariantCulture)
                });
            var tableProperties = new TableProperties();
            if (tableStyleIndex is ushort appliedStyle &&
                styles.TryGetValue(appliedStyle, out var tableStyle) &&
                tableStyle.Type == 3)
                tableProperties.AppendChild(new TableStyle { Val = tableStyle.StyleId });
            if (tableRightToLeft) tableProperties.AppendChild(new BiDiVisual());
            if (tablePreferredWidth is { } preferredWidth)
                tableProperties.AppendChild(new TableWidth
                {
                    Type = preferredWidth.Unit switch
                    {
                        1 => TableWidthUnitValues.Auto,
                        2 => TableWidthUnitValues.Pct,
                        3 => TableWidthUnitValues.Dxa,
                        _ => throw new InvalidDataException("A DOC table has unsupported width units.")
                    },
                    Width = preferredWidth.Value.ToString(
                        System.Globalization.CultureInfo.InvariantCulture)
                });
            if (tableJustification is byte tableAlignment)
                tableProperties.AppendChild(new TableJustification
                {
                    Val = tableAlignment switch
                    {
                        0 => TableRowAlignmentValues.Left,
                        1 => TableRowAlignmentValues.Center,
                        2 => TableRowAlignmentValues.Right,
                        _ => throw new InvalidDataException("A DOC table has invalid justification.")
                    }
                });
            if (uniformCellSpacing && tableCellSpacingTwips is ushort spacing)
                tableProperties.AppendChild(new TableCellSpacing
                {
                    Width = spacing.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    Type = TableWidthUnitValues.Dxa
                });
            // Usually DOC's zero table indent preserves Word's default
            // cell-edge offset. A fully bordered row with a matching origin,
            // or a generated row whose first edge includes that inset, needs
            // an explicit zero to keep the table at the DOCX page margin.
            var sideInset = ((tableDefaultMargins?.Left ?? 108) +
                (tableDefaultMargins?.Right ?? 108)) / 2;
            var uniformSideInsets = rowDefaultMargins.All(margins =>
                (margins?.Left ?? 108) == (tableDefaultMargins?.Left ?? 108) &&
                (margins?.Right ?? 108) == (tableDefaultMargins?.Right ?? 108));
            // A Word DOC row with an explicit zero origin/indent and no cell
            // margin or table style renders flush to the cell edges. If those
            // properties are omitted in DOCX, Word adds its intrinsic inset.
            var zeroInsetRow = tableStyleIndex == null &&
                tableDefaultMargins == null &&
                rowDefaultMargins.All(margins => margins == null) &&
                firstRowOriginTwips == 0 && tableIndentTwips == 0;
            if (zeroInsetRow)
                tableProperties.AppendChild(new TableIndentation
                {
                    Width = 0,
                    Type = TableWidthUnitValues.Dxa
                });
            // A row with a negative origin and a direct left cell border
            // retains a zero table indent in Word, including the styled
            // fixed row emitted by this writer. Omitting tblInd on DOCX
            // projection shifts that visible border inward.
            else if (firstRowOriginTwips < 0 && tableIndentTwips is (null or 0) &&
                rowEdges[0].Row.Elements<TableCell>()
                    .FirstOrDefault()?.TableCellProperties?.TableCellBorders is
                    { LeftBorder: not null } or { StartBorder: not null })
                tableProperties.AppendChild(new TableIndentation
                {
                    Width = 0, Type = TableWidthUnitValues.Dxa
                });
            else if (tableAutoFit && growAutofit &&
                tableIndentTwips is null or 0)
                tableProperties.AppendChild(new TableIndentation
                {
                    Width = 0,
                    Type = TableWidthUnitValues.Dxa
                });
            // Word's DOC save records the intrinsic cell-side inset as a
            // nominal table indent for occupied auto-fit horizontal merges.
            // Emitting it as tblInd shifts all three story tables in DOCX.
            else if (tableIndentTwips is short indentTwips &&
                !(tableAutoFit && tableStyleIndex == null &&
                  indentTwips == sideInset &&
                  rowEdges.Any(entry => entry.Row.Elements<TableCell>().Any(cell =>
                      cell.TableCellProperties?.GetFirstChild<HorizontalMerge>()?
                          .Val?.Value == MergedCellValues.Continue &&
                      !string.IsNullOrEmpty(cell.InnerText)))) &&
                (indentTwips != 0 ||
                    (uniformSideInsets &&
                        (tableDefaultMargins?.Left ?? 108) ==
                            (tableDefaultMargins?.Right ?? 108) &&
                        (firstRowBorders is
                            { Top.Type: > 0, Left.Type: > 0,
                              Bottom.Type: > 0, Right.Type: > 0,
                              InsideHorizontal.Type: > 0,
                              InsideVertical.Type: > 0 } ||
                            rowEdges[0].Edges[0] < 0) &&
                        firstRowOriginTwips > sideInset)))
                tableProperties.AppendChild(new TableIndentation
                {
                    Width = indentTwips,
                    Type = TableWidthUnitValues.Dxa
                });
            if (tableBackgroundShading is { } shade &&
                DocShadingPatterns.ToOpenXml(shade.Pattern) is { } tablePattern)
            {
                var shading = new Shading { Val = tablePattern };
                if (ColorValue(shade.FillRgb) is string fill) shading.Fill = fill;
                if (ColorValue(shade.ForegroundRgb) is string foreground)
                    shading.Color = foreground;
                tableProperties.AppendChild(shading);
            }
            tableProperties.AppendChild(new TableLayout
            {
                Type = tableAutoFit ? TableLayoutValues.Autofit : TableLayoutValues.Fixed
            });
            if (zeroInsetRow)
                tableProperties.AppendChild(CreateDefaultCellMargins(
                    new DocCellMargins(Left: 0, Right: 0)));
            else if (tableDefaultMargins is { } defaultMargins)
                tableProperties.AppendChild(CreateDefaultCellMargins(defaultMargins));
            if (tableLookMask is ushort mask)
                tableProperties.AppendChild(new TableLook
                {
                    Val = mask.ToString("X4", System.Globalization.CultureInfo.InvariantCulture)
                });
            table.InsertAt(tableProperties, 0);
            table.InsertAt(grid, 1);
            result.Add(table);
        }
        return result;
    }

    private static TableCellMargin CreateCellMargins(DocCellMargins source)
    {
        var margin = new TableCellMargin();
        if (source.Top is ushort top) margin.AppendChild(new TopMargin
        { Width = top.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Type = TableWidthUnitValues.Dxa });
        if (source.Left is ushort left) margin.AppendChild(new LeftMargin
        { Width = left.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Type = TableWidthUnitValues.Dxa });
        if (source.Bottom is ushort bottom) margin.AppendChild(new BottomMargin
        { Width = bottom.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Type = TableWidthUnitValues.Dxa });
        if (source.Right is ushort right) margin.AppendChild(new RightMargin
        { Width = right.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Type = TableWidthUnitValues.Dxa });
        return margin;
    }

    private static TableCellMarginDefault CreateDefaultCellMargins(DocCellMargins source)
    {
        var margin = new TableCellMarginDefault();
        if (source.Top is ushort top) margin.AppendChild(new TopMargin
        { Width = top.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Type = TableWidthUnitValues.Dxa });
        if (source.Left is ushort left) margin.AppendChild(new TableCellLeftMargin
        { Width = checked((short)left), Type = TableWidthValues.Dxa });
        if (source.Bottom is ushort bottom) margin.AppendChild(new BottomMargin
        { Width = bottom.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Type = TableWidthUnitValues.Dxa });
        if (source.Right is ushort right) margin.AppendChild(new TableCellRightMargin
        { Width = checked((short)right), Type = TableWidthValues.Dxa });
        return margin;
    }

    private static DocParagraphStyleRange? FindParagraphRange(
        IReadOnlyList<DocParagraphStyleRange> ranges, uint cp)
    {
        var lower = 0;
        var upper = ranges.Count;
        while (lower < upper)
        {
            var middle = (lower + upper) / 2;
            if (ranges[middle].CpStart <= cp) lower = middle + 1;
            else upper = middle;
        }
        var index = lower - 1;
        return index >= 0 && cp < ranges[index].CpEnd ? ranges[index] : null;
    }

    private static void AppendParagraphFormatting(DocumentFormat.OpenXml.OpenXmlCompositeElement properties,
        DocParagraphFormatting formatting)
    {
        if (formatting.KeepWithNext is bool keepWithNext)
            properties.AppendChild(new KeepNext { Val = keepWithNext });
        if (formatting.KeepLines is bool keepLines)
            properties.AppendChild(new KeepLines { Val = keepLines });
        if (formatting.PageBreakBefore is bool pageBreakBefore)
            properties.AppendChild(new PageBreakBefore { Val = pageBreakBefore });
        if (formatting.WidowControl is bool widowControl)
            properties.AppendChild(new WidowControl { Val = widowControl });
        if (formatting.ListOverrideIndex is > 0 and var listIndex)
            properties.AppendChild(new NumberingProperties(
                new NumberingLevelReference { Val = formatting.ListLevel ?? 0 },
                new NumberingId { Val = listIndex }));
        if (formatting.TopBorder != null || formatting.LeftBorder != null ||
            formatting.BottomBorder != null || formatting.RightBorder != null ||
            formatting.BetweenBorder != null)
        {
            var borders = new ParagraphBorders();
            if (formatting.TopBorder is { } top)
            { var border = new TopBorder(); top.ApplyTo(border); borders.AppendChild(border); }
            if (formatting.LeftBorder is { } left)
            { var border = new LeftBorder(); left.ApplyTo(border); borders.AppendChild(border); }
            if (formatting.BottomBorder is { } bottom)
            { var border = new BottomBorder(); bottom.ApplyTo(border); borders.AppendChild(border); }
            if (formatting.RightBorder is { } right)
            { var border = new RightBorder(); right.ApplyTo(border); borders.AppendChild(border); }
            if (formatting.BetweenBorder is { } between)
            { var border = new BetweenBorder(); between.ApplyTo(border); borders.AppendChild(border); }
            properties.AppendChild(borders);
        }
        if (formatting.ShadingPattern is ushort pattern &&
            (pattern == ushort.MaxValue ? ShadingPatternValues.Nil :
                DocShadingPatterns.ToOpenXml(pattern)) is { } shadingValue)
        {
            static string Color(uint rgb)
            {
                var displayRgb = ((rgb & 0xFF) << 16) | (rgb & 0x00FF00) | (rgb >> 16);
                return displayRgb.ToString("X6", System.Globalization.CultureInfo.InvariantCulture);
            }
            var shading = new Shading { Val = shadingValue };
            if (formatting.FillRgb is uint fill) shading.Fill = Color(fill);
            if (formatting.ShadingForegroundRgb is uint foreground)
                shading.Color = Color(foreground);
            properties.AppendChild(shading);
        }
        if ((formatting.ClearedTabPositions?.Count ?? 0) > 0 ||
            (formatting.TabStops?.Count ?? 0) > 0)
        {
            var tabs = new Tabs();
            foreach (var position in formatting.ClearedTabPositions ?? Array.Empty<short>())
                tabs.AppendChild(new TabStop
                {
                    Position = position,
                    Val = TabStopValues.Clear
                });
            foreach (var tab in formatting.TabStops ?? Array.Empty<DocTabStop>())
            {
                if (tab.Alignment is not (0 or 1 or 2 or 3 or 4 or 6) ||
                    tab.Leader is not (0 or 1 or 2 or 3 or 4 or 5 or 7))
                    continue;
                var alignment = tab.Alignment switch
                {
                    0 => TabStopValues.Left,
                    1 => TabStopValues.Center,
                    2 => TabStopValues.Right,
                    3 => TabStopValues.Decimal,
                    4 => TabStopValues.Bar,
                    6 => TabStopValues.Number,
                    _ => throw new InvalidDataException("Unsupported tab alignment.")
                };
                var leader = tab.Leader switch
                {
                    0 => TabStopLeaderCharValues.None,
                    1 => TabStopLeaderCharValues.Dot,
                    2 => TabStopLeaderCharValues.Hyphen,
                    3 => TabStopLeaderCharValues.Underscore,
                    4 => TabStopLeaderCharValues.Heavy,
                    5 => TabStopLeaderCharValues.MiddleDot,
                    7 => TabStopLeaderCharValues.None,
                    _ => throw new InvalidDataException("Unsupported tab leader.")
                };
                tabs.AppendChild(new TabStop { Position = tab.PositionTwips,
                    Val = alignment, Leader = leader });
            }
            properties.AppendChild(tabs);
        }
        if (formatting.SuppressAutoHyphens is bool suppressAutoHyphens)
            properties.AppendChild(new SuppressAutoHyphens { Val = suppressAutoHyphens });
        if (formatting.SuppressLineNumbers is bool suppressLineNumbers)
            properties.AppendChild(new SuppressLineNumbers { Val = suppressLineNumbers });
        if (formatting.Kinsoku is bool kinsoku)
            properties.AppendChild(new Kinsoku { Val = kinsoku });
        if (formatting.WordWrap is bool wordWrap)
            properties.AppendChild(new WordWrap { Val = wordWrap });
        if (formatting.AutoSpaceDE is bool autoSpaceDE)
            properties.AppendChild(new AutoSpaceDE { Val = autoSpaceDE });
        if (formatting.AutoSpaceDN is bool autoSpaceDN)
            properties.AppendChild(new AutoSpaceDN { Val = autoSpaceDN });
        if (formatting.SnapToGrid is bool snapToGrid)
            properties.AppendChild(new SnapToGrid { Val = snapToGrid });
        if (formatting.AdjustRightIndent is bool adjustRightIndent)
            properties.AppendChild(new AdjustRightIndent { Val = adjustRightIndent });
        if (formatting.ParagraphRightToLeft is bool paragraphRightToLeft)
            properties.AppendChild(new BiDi { Val = paragraphRightToLeft });
        if (formatting.BeforeTwips != null || formatting.AfterTwips != null ||
            formatting.LineValue != null || formatting.BeforeAutoSpacing != null ||
            formatting.AfterAutoSpacing != null || formatting.BeforeLines != null ||
            formatting.AfterLines != null)
        {
            var spacing = new SpacingBetweenLines();
            if (formatting.BeforeTwips is ushort before)
                spacing.Before = before.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (formatting.AfterTwips is ushort after)
                spacing.After = after.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (formatting.BeforeAutoSpacing is bool beforeAutoSpacing)
                spacing.BeforeAutoSpacing = beforeAutoSpacing;
            if (formatting.AfterAutoSpacing is bool afterAutoSpacing)
                spacing.AfterAutoSpacing = afterAutoSpacing;
            if (formatting.BeforeLines is short beforeLines)
                spacing.BeforeLines = beforeLines;
            if (formatting.AfterLines is short afterLines)
                spacing.AfterLines = afterLines;
            if (formatting.LineValue is short line)
            {
                spacing.Line = Math.Abs((int)line).ToString(
                    System.Globalization.CultureInfo.InvariantCulture);
                spacing.LineRule = formatting.LineIsMultiple == true
                    ? LineSpacingRuleValues.Auto
                    : line < 0 ? LineSpacingRuleValues.Exact : LineSpacingRuleValues.AtLeast;
            }
            properties.AppendChild(spacing);
        }
        if (formatting.LeftTwips != null || formatting.RightTwips != null ||
            formatting.FirstLineTwips != null || formatting.LeftChars != null ||
            formatting.RightChars != null || formatting.FirstLineChars != null)
        {
            var indent = new Indentation();
            if (formatting.LeftTwips is short left)
                indent.Left = left.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (formatting.RightTwips is short right)
                indent.Right = right.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (formatting.LeftChars is short leftChars)
                indent.LeftChars = leftChars;
            if (formatting.RightChars is short rightChars)
                indent.RightChars = rightChars;
            if (formatting.FirstLineTwips is short firstLine)
            {
                if (firstLine >= 0)
                    indent.FirstLine = firstLine.ToString(System.Globalization.CultureInfo.InvariantCulture);
                else
                    indent.Hanging = (-firstLine).ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            if (formatting.FirstLineChars is short firstLineChars)
            {
                if (firstLineChars >= 0)
                    indent.FirstLineChars = firstLineChars;
                else
                    indent.HangingChars = -firstLineChars;
            }
            properties.AppendChild(indent);
        }
        if (formatting.ContextualSpacing is bool contextualSpacing)
            properties.AppendChild(new ContextualSpacing { Val = contextualSpacing });
        if (formatting.MirrorIndents is bool mirrorIndents)
            properties.AppendChild(new MirrorIndents { Val = mirrorIndents });
        if (formatting.Justification is byte value)
        {
            var alignment = value switch
            {
                0 => JustificationValues.Left,
                1 => JustificationValues.Center,
                2 => JustificationValues.Right,
                3 => JustificationValues.Both,
                4 => JustificationValues.Distribute,
                _ => (JustificationValues?)null
            };
            if (alignment != null) properties.AppendChild(new Justification { Val = alignment.Value });
        }
        if (formatting.TextAlignmentCode is short textAlignmentCode)
            properties.AppendChild(new TextAlignment { Val = textAlignmentCode switch
            {
                0 => VerticalTextAlignmentValues.Top,
                1 => VerticalTextAlignmentValues.Center,
                2 => VerticalTextAlignmentValues.Baseline,
                3 => VerticalTextAlignmentValues.Bottom,
                4 => VerticalTextAlignmentValues.Auto,
                _ => throw new InvalidDataException("Invalid paragraph text alignment.")
            } });
        if (formatting.OutlineLevel is byte outlineLevel)
            properties.AppendChild(new OutlineLevel { Val = outlineLevel });
    }

    private static void AppendControlRun(Paragraph paragraph,
        DocumentFormat.OpenXml.OpenXmlElement control, uint cp,
        IReadOnlyList<DocCharacterFormattingRange> ranges,
        IReadOnlyDictionary<int, DocStyleDefinition> styles)
    {
        var run = new Run();
        var range = FindCharacterRange(ranges, cp, out _);
        if (range is { Formatting.IsEmpty: false })
        {
            var properties = new RunProperties();
            AppendCharacterFormatting(properties, range.Formatting, styles);
            run.AppendChild(properties);
        }
        run.AppendChild(control);
        if (range?.Formatting.DeletedRevision == true)
            paragraph.AppendChild(new DeletedRun(run)
            {
                Id = cp.ToString(System.Globalization.CultureInfo.InvariantCulture),
                Author = range.Formatting.DeletedRevisionAuthor ?? "Unknown",
                Date = range.Formatting.DeletedRevisionAt
            });
        else if (range?.Formatting.InsertedRevision == true)
            paragraph.AppendChild(new InsertedRun(run)
            {
                Id = cp.ToString(System.Globalization.CultureInfo.InvariantCulture),
                Author = range.Formatting.InsertedRevisionAuthor ?? "Unknown",
                Date = range.Formatting.InsertedRevisionAt
            });
        else paragraph.AppendChild(run);
    }

    private static void AppendTextRuns(Paragraph paragraph, DocStoryAtom atom,
        IReadOnlyList<DocCharacterFormattingRange> ranges,
        IReadOnlyDictionary<int, DocStyleDefinition> styles, bool fieldCode = false)
    {
        var cp = atom.CpStart;
        while (cp < atom.CpEnd)
        {
            var range = FindCharacterRange(ranges, cp, out var lower);
            var next = Math.Min(atom.CpEnd, range?.CpEnd ??
                (lower < ranges.Count ? ranges[lower].CpStart : atom.CpEnd));
            if (next <= cp) throw new InvalidDataException("Character formatting ranges overlap.");
            var text = atom.Text.Substring(checked((int)(cp - atom.CpStart)),
                checked((int)(next - cp)));
            var run = new Run();
            if (range is { Formatting.IsEmpty: false })
            {
                var properties = new RunProperties();
                AppendCharacterFormatting(properties, range.Formatting, styles);
                run.AppendChild(properties);
            }
            if (fieldCode)
                run.AppendChild(new FieldCode
                {
                    Text = text,
                    Space = SpaceProcessingModeValues.Preserve
                });
            else if (!fieldCode && range?.Formatting.SymbolCharacter is ushort symbolCode &&
                range.Formatting.SymbolFontName is string symbolFont && text == "(")
                run.AppendChild(new SymbolChar
                {
                    Font = symbolFont,
                    Char = symbolCode.ToString("X4", CultureInfo.InvariantCulture)
                });
            else if (range?.Formatting.DeletedRevision == true)
                run.AppendChild(new DeletedText(text)
                    { Space = SpaceProcessingModeValues.Preserve });
            else
                run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            if (range?.Formatting.DeletedRevision == true)
            {
                paragraph.AppendChild(new DeletedRun(run)
                {
                    Id = cp.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    Author = range.Formatting.DeletedRevisionAuthor ?? "Unknown",
                    Date = range.Formatting.DeletedRevisionAt
                });
            }
            else if (range?.Formatting.InsertedRevision == true)
            {
                paragraph.AppendChild(new InsertedRun(run)
                {
                    Id = cp.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    Author = range.Formatting.InsertedRevisionAuthor ?? "Unknown",
                    Date = range.Formatting.InsertedRevisionAt
                });
            }
            else
            {
                paragraph.AppendChild(run);
            }
            cp = next;
        }
    }

    private static DocCharacterFormattingRange? FindCharacterRange(
        IReadOnlyList<DocCharacterFormattingRange> ranges, uint cp, out int nextIndex)
    {
        var lower = 0;
        var upper = ranges.Count;
        while (lower < upper)
        {
            var middle = (lower + upper) / 2;
            if (ranges[middle].CpStart <= cp) lower = middle + 1;
            else upper = middle;
        }
        nextIndex = lower;
        var index = lower - 1;
        return index >= 0 && cp < ranges[index].CpEnd ? ranges[index] : null;
    }

    private static EmphasisMarkValues? EmphasisValue(byte? code) => code switch
    {
        null => null,
        0 => EmphasisMarkValues.None,
        1 => EmphasisMarkValues.Dot,
        2 => EmphasisMarkValues.Comma,
        3 => EmphasisMarkValues.Circle,
        4 => EmphasisMarkValues.UnderDot,
        _ => throw new InvalidDataException("A DOC emphasis mark is invalid.")
    };

    private static bool? ResolveStyleRightToLeft(int styleIndex,
        IReadOnlyDictionary<int, DocStyleDefinition> styles)
    {
        var seen = new HashSet<int>();
        while (seen.Add(styleIndex) && styles.TryGetValue(styleIndex, out var style))
        {
            if (style.CharacterFormatting.RightToLeftText is bool direction)
                return direction;
            if (style.BasedOnIndex is not int parent) break;
            styleIndex = parent;
        }
        return null;
    }

    private static void ApplyStyleRunFormatting(
        IEnumerable<Paragraph> paragraphs,
        IReadOnlyDictionary<int, DocStyleDefinition> styles)
    {
        foreach (var paragraph in paragraphs)
        {
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            var paragraphStyle = styleId != null &&
                styleId.StartsWith("DocStyle", StringComparison.Ordinal) &&
                int.TryParse(styleId.Substring("DocStyle".Length), out var styleIndex) &&
                styles.TryGetValue(styleIndex, out var style) && style.Type == 1
                    ? styleIndex : styles.TryGetValue(0, out var normal) && normal.Type == 1
                        ? 0 : (int?)null;
            var direction = paragraphStyle is int paragraphIndex
                ? ResolveStyleRightToLeft(paragraphIndex, styles) : null;
            var paragraphHighlight = paragraphStyle is int highlightStyleIndex
                ? ResolveStyleHighlight(highlightStyleIndex, styles) : null;
            foreach (var run in paragraph.Descendants<Run>())
            {
                var properties = run.RunProperties ??= new RunProperties();
                if (direction is bool rightToLeft && properties.RightToLeftText == null)
                    properties.AppendChild(new RightToLeftText { Val = rightToLeft });
                if (properties.GetFirstChild<Highlight>() != null) continue;
                var runStyleId = properties.RunStyle?.Val?.Value;
                var characterHighlight = runStyleId != null &&
                    runStyleId.StartsWith("DocStyle", StringComparison.Ordinal) &&
                    int.TryParse(runStyleId.Substring("DocStyle".Length), out var runStyleIndex) &&
                    styles.TryGetValue(runStyleIndex, out var characterStyle) &&
                    characterStyle.Type == 2
                        ? ResolveStyleHighlight(runStyleIndex, styles) : null;
                var highlight = HighlightValue(characterHighlight ?? paragraphHighlight);
                if (highlight == null) continue;
                // DOC style CHPX can highlight text; w:highlight is invalid in
                // DOCX style rPr, so materialize its visible value on the run.
                var node = new Highlight { Val = highlight.Value };
                var following = properties.ChildElements.FirstOrDefault(x =>
                    x is Underline or Border or Shading or VerticalTextAlignment or
                        RightToLeftText or ComplexScript or Languages);
                if (following == null) properties.AppendChild(node);
                else properties.InsertBefore(node, following);
            }
        }
    }

    private static byte? ResolveStyleHighlight(int styleIndex,
        IReadOnlyDictionary<int, DocStyleDefinition> styles)
    {
        var seen = new HashSet<int>();
        while (seen.Add(styleIndex) && styles.TryGetValue(styleIndex, out var style))
        {
            if (style.CharacterFormatting.HighlightCode is byte highlight)
                return highlight;
            if (style.BasedOnIndex is not int parent) break;
            styleIndex = parent;
        }
        return null;
    }

    private static void AppendCharacterFormatting(
        DocumentFormat.OpenXml.OpenXmlCompositeElement properties,
        DocCharacterFormatting formatting,
        IReadOnlyDictionary<int, DocStyleDefinition> styles)
    {
        if (formatting.CharacterStyleIndex is int styleIndex &&
            styles.TryGetValue(styleIndex, out var style) && style.Type == 2)
            properties.AppendChild(new RunStyle { Val = style.StyleId });
        AppendRunFonts(properties, formatting);
        if (formatting.Bold is bool bold)
            properties.AppendChild(new Bold { Val = bold });
        if (formatting.ComplexScriptBold is bool complexBold)
            properties.AppendChild(new BoldComplexScript { Val = complexBold });
        if (formatting.Italic is bool italic)
            properties.AppendChild(new Italic { Val = italic });
        if (formatting.ComplexScriptItalic is bool complexItalic)
            properties.AppendChild(new ItalicComplexScript { Val = complexItalic });
        if (formatting.Caps is bool caps)
            properties.AppendChild(new Caps { Val = caps });
        if (formatting.SmallCaps is bool smallCaps)
            properties.AppendChild(new SmallCaps { Val = smallCaps });
        if (formatting.Strike is bool strike)
            properties.AppendChild(new Strike { Val = strike });
        if (formatting.DoubleStrike is bool doubleStrike)
            properties.AppendChild(new DoubleStrike { Val = doubleStrike });
        if (formatting.Outline is bool outline)
            properties.AppendChild(new Outline { Val = outline });
        if (formatting.Shadow is bool shadow)
            properties.AppendChild(new Shadow { Val = shadow });
        if (formatting.Emboss is bool emboss)
            properties.AppendChild(new Emboss { Val = emboss });
        if (formatting.Imprint is bool imprint)
            properties.AppendChild(new Imprint { Val = imprint });
        if (formatting.Hidden is bool hidden)
            properties.AppendChild(new Vanish { Val = hidden });
        if (formatting.SnapToGrid is bool snap)
            properties.AppendChild(new SnapToGrid { Val = snap });
        if (ColorValue(formatting.ColorRef) is string color)
            properties.AppendChild(new Color { Val = color });
        AppendCharacterSpacing(properties, formatting);
        if (formatting.CharacterScalePercent is ushort scale)
            properties.AppendChild(new CharacterScale { Val = scale });
        AppendKerning(properties, formatting);
        if (formatting.BaselineOffsetHalfPoints is short offset)
            properties.AppendChild(new Position { Val = offset.ToString(
                System.Globalization.CultureInfo.InvariantCulture) });
        if (formatting.SizeHalfPoints is ushort size)
            properties.AppendChild(new FontSize
            {
                Val = size.ToString(System.Globalization.CultureInfo.InvariantCulture)
            });
        if (formatting.ComplexScriptSizeHalfPoints is ushort complexSize)
            properties.AppendChild(new FontSizeComplexScript
            {
                Val = complexSize.ToString(System.Globalization.CultureInfo.InvariantCulture)
            });
        if (HighlightValue(formatting.HighlightCode) is HighlightColorValues highlight)
            properties.AppendChild(new Highlight { Val = highlight });
        AppendUnderline(properties, formatting);
        AppendRunBorder(properties, formatting);
        AppendRunShading(properties, formatting);
        AppendFitText(properties, formatting);
        if (ScriptValue(formatting.ScriptCode) is VerticalPositionValues script)
            properties.AppendChild(new VerticalTextAlignment { Val = script });
        bool? rightToLeft = formatting.RightToLeftText;
        if (rightToLeft == null && formatting.CharacterStyleIndex is int appliedStyle)
            rightToLeft = ResolveStyleRightToLeft(appliedStyle, styles);
        if (rightToLeft is bool rightToLeftValue)
            properties.AppendChild(new RightToLeftText { Val = rightToLeftValue });
        if (formatting.ForceComplexScript is bool forceComplexScript)
            properties.AppendChild(new ComplexScript { Val = forceComplexScript });
        if (EmphasisValue(formatting.EmphasisMarkCode) is EmphasisMarkValues emphasis)
            properties.AppendChild(new Emphasis { Val = emphasis });
        AppendLanguages(properties, formatting);
    }

    private static void AppendFitText(OpenXmlCompositeElement properties,
        DocCharacterFormatting formatting)
    {
        if (formatting.FitText is not { } fitText) return;
        if (fitText.WidthTwips < 0)
            throw new NotSupportedException("Negative DOC fit-text minimum widths have no equivalent w:fitText width.");
        if (fitText.WidthTwips == 0) return; // Native DOC ignores the operand.
        properties.AppendChild(new FitText
        { Val = checked((uint)fitText.WidthTwips), Id = fitText.Id });
    }

    private static void AppendCharacterSpacing(
        DocumentFormat.OpenXml.OpenXmlCompositeElement properties,
        DocCharacterFormatting formatting)
    {
        if (formatting.CharacterSpacingTwips is short spacing)
            properties.AppendChild(new Spacing { Val = spacing });
    }

    private static void AppendKerning(
        DocumentFormat.OpenXml.OpenXmlCompositeElement properties,
        DocCharacterFormatting formatting)
    {
        if (formatting.KerningThresholdHalfPoints is ushort threshold)
            properties.AppendChild(new Kern { Val = threshold });
    }

    private static void AppendLanguages(
        DocumentFormat.OpenXml.OpenXmlCompositeElement properties,
        DocCharacterFormatting formatting)
    {
        static string? Name(ushort? id)
        {
            if (id is not ushort value) return null;
            try { return System.Globalization.CultureInfo.GetCultureInfo(value).Name; }
            catch (System.Globalization.CultureNotFoundException) { return null; }
        }
        var normal = Name(formatting.LanguageId);
        var eastAsia = Name(formatting.EastAsiaLanguageId);
        var complex = Name(formatting.ComplexScriptLanguageId);
        if (normal == null && eastAsia == null && complex == null) return;
        var languages = new Languages();
        if (normal != null) languages.Val = normal;
        if (eastAsia != null) languages.EastAsia = eastAsia;
        if (complex != null) languages.Bidi = complex;
        properties.AppendChild(languages);
    }

    private static UnderlineValues? UnderlineValue(byte? code) => code switch
    {
        0 => UnderlineValues.None,
        1 => UnderlineValues.Single,
        2 => UnderlineValues.Words,
        3 => UnderlineValues.Double,
        4 => UnderlineValues.Dotted,
        6 => UnderlineValues.Thick,
        7 => UnderlineValues.Dash,
        9 => UnderlineValues.DotDash,
        10 => UnderlineValues.DotDotDash,
        11 => UnderlineValues.Wave,
        20 => UnderlineValues.DottedHeavy,
        23 => UnderlineValues.DashedHeavy,
        25 => UnderlineValues.DashDotHeavy,
        26 => UnderlineValues.DashDotDotHeavy,
        27 => UnderlineValues.WavyHeavy,
        39 => UnderlineValues.DashLong,
        43 => UnderlineValues.WavyDouble,
        55 => UnderlineValues.DashLongHeavy,
        _ => null
    };

    private static void AppendUnderline(
        DocumentFormat.OpenXml.OpenXmlCompositeElement properties,
        DocCharacterFormatting formatting)
    {
        var style = UnderlineValue(formatting.UnderlineCode);
        var color = ColorValue(formatting.UnderlineColorRef);
        if (style == null && color == null) return;
        var underline = new Underline();
        if (style is UnderlineValues value) underline.Val = value;
        if (color != null) underline.Color = color;
        properties.AppendChild(underline);
    }

    private static void AppendRunBorder(
        DocumentFormat.OpenXml.OpenXmlCompositeElement properties,
        DocCharacterFormatting formatting)
    {
        if (formatting.Border is not { } source) return;
        var border = new Border();
        source.ApplyTo(border);
        properties.AppendChild(border);
    }

    private static void AppendRunShading(
        DocumentFormat.OpenXml.OpenXmlCompositeElement properties,
        DocCharacterFormatting formatting)
    {
        if (formatting.Shading is not { } source) return;
        var pattern = source.Pattern == ushort.MaxValue
            ? ShadingPatternValues.Nil :
                DocShadingPatterns.ToOpenXml(source.Pattern);
        if (pattern is not { } value)
            return;
        var shading = new Shading { Val = value };
        if (ColorValue(source.FillRgb) is string fill) shading.Fill = fill;
        if (ColorValue(source.ForegroundRgb) is string foreground)
            shading.Color = foreground;
        properties.AppendChild(shading);
    }

    private static void AppendRunFonts(DocumentFormat.OpenXml.OpenXmlCompositeElement properties,
        DocCharacterFormatting formatting)
    {
        if (formatting.AsciiFontName == null && formatting.EastAsiaFontName == null &&
            formatting.HighAnsiFontName == null &&
            formatting.ComplexScriptFontName == null) return;
        var fonts = new RunFonts();
        if (formatting.AsciiFontName != null) fonts.Ascii = formatting.AsciiFontName;
        if (formatting.EastAsiaFontName != null) fonts.EastAsia = formatting.EastAsiaFontName;
        if (formatting.HighAnsiFontName != null) fonts.HighAnsi = formatting.HighAnsiFontName;
        if (formatting.ComplexScriptFontName != null)
            fonts.ComplexScript = formatting.ComplexScriptFontName;
        properties.AppendChild(fonts);
    }

    private static string? ColorValue(uint? colorRef)
    {
        if (colorRef is not uint value) return null;
        var auto = value >> 24;
        if (auto == 0xFF) return "auto";
        if (auto != 0) return null;
        return $"{(byte)value:X2}{(byte)(value >> 8):X2}{(byte)(value >> 16):X2}";
    }

    private static HighlightColorValues? HighlightValue(byte? code) => code switch
    {
        0 => HighlightColorValues.None,
        1 => HighlightColorValues.Black,
        2 => HighlightColorValues.Blue,
        3 => HighlightColorValues.Cyan,
        4 => HighlightColorValues.Green,
        5 => HighlightColorValues.Magenta,
        6 => HighlightColorValues.Red,
        7 => HighlightColorValues.Yellow,
        8 => HighlightColorValues.White,
        9 => HighlightColorValues.DarkBlue,
        10 => HighlightColorValues.DarkCyan,
        11 => HighlightColorValues.DarkGreen,
        12 => HighlightColorValues.DarkMagenta,
        13 => HighlightColorValues.DarkRed,
        14 => HighlightColorValues.DarkYellow,
        15 => HighlightColorValues.DarkGray,
        16 => HighlightColorValues.LightGray,
        _ => null
    };

    private static VerticalPositionValues? ScriptValue(byte? code) => code switch
    {
        0 => VerticalPositionValues.Baseline,
        1 => VerticalPositionValues.Superscript,
        2 => VerticalPositionValues.Subscript,
        _ => null
    };

    private static HeaderFooterValues ReferenceType(int slotInSection) => slotInSection switch
    {
        0 or 2 => HeaderFooterValues.Even,
        1 or 3 => HeaderFooterValues.Default,
        _ => HeaderFooterValues.First
    };

    private static void Count(IDictionary<ushort, int> counts, char value)
    {
        var key = (ushort)value;
        counts[key] = counts.TryGetValue(key, out var existing) ? existing + 1 : 1;
    }

    private static IEnumerable<string> EnumerateStreams(DocStructureNode node)
    {
        if (node.Kind == "Stream" && node.Attributes.TryGetValue("path", out var path))
            yield return path;
        foreach (var child in node.Children)
            foreach (var stream in EnumerateStreams(child))
                yield return stream;
    }
}
