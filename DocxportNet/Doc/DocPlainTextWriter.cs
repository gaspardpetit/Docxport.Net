using System.Buffers.Binary;
using System.Text;
using OpenMcdf;

namespace DocxportNet.Doc;

/// <summary>Writes Unicode body and section stories with their captured formatting.</summary>
internal static class DocPlainTextWriter
{
    private const int PageSize = 512;

    public static void Write(Stream output, DocPlainTextDocument document)
    {
        if (output == null) throw new ArgumentNullException(nameof(output));
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (!output.CanWrite) throw new ArgumentException("The output stream must be writable.", nameof(output));
        var mainStory = document.MainStory ?? throw new InvalidDataException(
            "The DOC body story is missing.");
        DocStoryParagraphRange.Validate(mainStory.Text, mainStory.Paragraphs);
        if (mainStory.ListParagraphs is { } mainLists)
            DocStoryListParagraph.Validate(mainLists, mainStory.ParagraphStyles,
                mainStory.TableRows);
        var text = mainStory.Text;
        if (!text.EndsWith("\r", StringComparison.Ordinal)) text += '\r';
        if (document.Sections.Count == 0 ||
            document.Sections[document.Sections.Count - 1].EndCp != text.Length ||
            document.Sections.Any(x => x.Stories.Count != 6))
            throw new InvalidDataException("The DOC section stories do not match the main text.");

        var precedingSectionEnd = 0;
        foreach (var section in document.Sections)
        {
            if (section.Model is { } model)
            {
                model.Validate();
                if (model.StartCp != precedingSectionEnd || model.EndCp != section.EndCp ||
                    Enumerable.Range(0, 6).Any(slot =>
                        model.Slots[slot].IsPresent != (section.Stories[slot] != null)))
                    throw new InvalidDataException("The captured section differs from its story slots.");
            }
            precedingSectionEnd = section.EndCp;
        }
        using var headerText = new StringWriter();
        var storyEnds = new List<int> { 0, 0, 0, 0, 0, 0, 0 };
        var formatRuns = new List<DocPlainTextFormatRun>();
        formatRuns.AddRange(mainStory.Runs.Select(x => x.ToWriter(0)));
        var paragraphStyleRuns = new List<DocPlainTextParagraphStyleRun>();
        paragraphStyleRuns.AddRange(mainStory.ParagraphStyles.Select(x => x.ToWriter(0)));
        paragraphStyleRuns.AddRange(mainStory.TableRows.Select(x =>
            new DocPlainTextParagraphStyleRun(x.Start, x.End, x.StyleIndex,
                x.Formatting)));
        var bookmarks = new List<DocPlainTextBookmark>();
        bookmarks.AddRange(mainStory.Bookmarks.Select(x =>
            new DocPlainTextBookmark(x.Name, x.Start, x.End)));
        var pictures = new List<DocPlainTextPicture>();
        pictures.AddRange(mainStory.Pictures.Select(x => new DocPlainTextPicture(
            x.Cp, x.Payload ?? throw new InvalidDataException(
                "A captured DOCX inline picture has no image payload."))));
        var floating = new List<DocPlainTextFloatingPicture>();
        floating.AddRange(mainStory.FloatingPictures.Select(x => x.ToWriter(x.Cp)));
        var headerFieldMarks = new List<DocStoryFieldMark>();
        var capturedHeaderFields = true;
        var hasHeaderStories = document.Sections.Any(x => x.Stories.Any(y => y != null));
        if (hasHeaderStories)
        {
            for (var sectionIndex = 0; sectionIndex < document.Sections.Count; sectionIndex++)
            {
                var section = document.Sections[sectionIndex];
                for (var slot = 0; slot < 6; slot++)
                {
                    var captured = section.Stories[slot];
                    if (captured != null)
                    {
                        var headerStart = headerText.GetStringBuilder().Length;
                        var start = checked(text.Length + headerStart);
                        var story = captured.Text;
                        if (captured.FieldMarks is { } storyFields)
                            headerFieldMarks.AddRange(storyFields.Select(x => x with
                            {
                                Cp = checked(headerStart + x.Cp)
                            }));
                        else capturedHeaderFields = false;
                        DocStoryParagraphRange.Validate(captured.Text, captured.Paragraphs);
                        if (captured.ListParagraphs is { } storyLists)
                            DocStoryListParagraph.Validate(storyLists,
                                captured.ParagraphStyles, captured.TableRows);
                        bookmarks.AddRange(captured.Bookmarks.Select(x =>
                            new DocPlainTextBookmark(x.Name,
                                checked(start + x.Start), checked(start + x.End))));
                        pictures.AddRange(captured.Pictures.Select(x =>
                            new DocPlainTextPicture(checked(start + x.Cp),
                                x.Payload ?? throw new InvalidDataException(
                                    "A captured DOCX inline picture has no image payload."))));
                        floating.AddRange(captured.FloatingPictures.Select(x =>
                            x.ToWriter(checked(start + x.Cp))));
                        formatRuns.AddRange(captured.Runs.Select(x => x.ToWriter(start)));
                        paragraphStyleRuns.AddRange(captured.ParagraphStyles.Select(
                            x => x.ToWriter(start)));
                        foreach (var row in captured.TableRows)
                            paragraphStyleRuns.Add(new DocPlainTextParagraphStyleRun(
                                checked(start + row.Start), checked(start + row.End),
                                row.StyleIndex, row.Formatting));
                        // A final table row mark already ends the story's content;
                        // another paragraph here would change header/footer height.
                        headerText.Write(story.EndsWith("\r", StringComparison.Ordinal) ||
                            story.EndsWith("\u0007", StringComparison.Ordinal)
                            ? story : story + '\r');
                        headerText.Write('\r'); // Guard paragraph mark.
                    }
                    storyEnds.Add(headerText.GetStringBuilder().Length);
                }
            }
            headerText.Write('\r'); // Final Header Document paragraph mark.
        }
        var header = headerText.ToString();
        var combined = hasHeaderStories ? text + header + '\r' : text;
        using var dataStream = new MemoryStream();
        var pictureOffsets = new Dictionary<int, int>();
        foreach (var picture in pictures)
            {
                if (picture.Cp < 0 || picture.Cp >= combined.Length ||
                    combined[picture.Cp] != '\u0001' ||
                    pictureOffsets.ContainsKey(picture.Cp))
                    throw new InvalidDataException("A DOC picture has no unique picture character.");
                pictureOffsets.Add(picture.Cp, DocInlinePictureWriter.Write(dataStream,
                    picture.Picture, checked((uint)(1025 + pictureOffsets.Count))));
            }
        foreach (var anchored in floating)
            if (anchored.Cp < 0 || anchored.Cp >= combined.Length ||
                combined[anchored.Cp] != '\u0008')
                throw new InvalidDataException("A DOC floating picture has no story shape character.");
        // Word reads the final main-story paragraph's PAPX from the terminal
        // paragraph record when a header subdocument follows the main story.
        // Mirror all paragraph properties on the nonvisible final mark.
        if (hasHeaderStories && paragraphStyleRuns.LastOrDefault(x => x.End == text.Length)
            is { } finalMainStyle)
            paragraphStyleRuns.Add(finalMainStyle with
            {
                Start = combined.Length - 1,
                End = combined.Length
            });
        if (hasHeaderStories && formatRuns.LastOrDefault(x => x.End == text.Length &&
            x.Start == text.Length - 1) is { } finalMainMark)
            formatRuns.Add(finalMainMark with
            {
                Start = combined.Length - 1,
                End = combined.Length
            });

        var fontNames = new List<string>();
        var fontIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        void AddFont(string? name)
        {
            if (string.IsNullOrWhiteSpace(name) || fontIndexes.ContainsKey(name)) return;
            fontIndexes.Add(name, fontNames.Count);
            fontNames.Add(name);
        }
        void AddFormattingFonts(DocCharacterFormatting formatting)
        {
            AddFont(formatting.AsciiFontName);
            AddFont(formatting.EastAsiaFontName);
            AddFont(formatting.HighAnsiFontName);
            AddFont(formatting.ComplexScriptFontName);
            AddFont(formatting.SymbolFontName);
        }
        var defaultFormatting = document.DefaultCharacterFormatting ?? DocCharacterFormatting.Empty;
        if (defaultFormatting.AsciiFontName == null &&
            defaultFormatting.EastAsiaFontName == null &&
            defaultFormatting.HighAnsiFontName == null)
            defaultFormatting = defaultFormatting with
            {
                AsciiFontName = "Times New Roman",
                EastAsiaFontName = "Times New Roman",
                HighAnsiFontName = "Times New Roman"
            };
        AddFormattingFonts(defaultFormatting);
        foreach (var run in formatRuns) AddFormattingFonts(run.Formatting);
        var revisionAuthorNames = new List<string> { "Unknown" };
        var revisionAuthorIndexes = new Dictionary<string, int>(StringComparer.Ordinal)
            { ["Unknown"] = 0 };
        void AddRevisionAuthor(string? name)
        {
            if (string.IsNullOrWhiteSpace(name) ||
                revisionAuthorIndexes.ContainsKey(name)) return;
            revisionAuthorIndexes.Add(name, revisionAuthorNames.Count);
            revisionAuthorNames.Add(name);
        }
        foreach (var run in formatRuns)
        {
            if (run.Formatting.DeletedRevision == true)
                AddRevisionAuthor(run.Formatting.DeletedRevisionAuthor);
            if (run.Formatting.InsertedRevision == true)
                AddRevisionAuthor(run.Formatting.InsertedRevisionAuthor);
        }
        foreach (var style in document.Styles ?? Array.Empty<DocStyleDefinition>())
            AddFormattingFonts(style.CharacterFormatting);
        if (document.Lists is { } fontLists)
        {
            foreach (var level in fontLists.Definitions.SelectMany(x => x.Levels))
                if (level.LabelFormatting is { } formatting)
                    AddFormattingFonts(formatting);
            foreach (var level in fontLists.Instances.SelectMany(x =>
                x.FormattingOverrides.Values))
                if (level.LabelFormatting is { } formatting)
                    AddFormattingFonts(formatting);
        }

        DocFitTextScaleResolver.Apply(formatRuns, combined, paragraphStyleRuns,
            document.Styles ?? [], defaultFormatting, document.DefaultParagraphFormatting);
        var encoded = Encoding.Unicode.GetBytes(combined);
        const int textOffset = 2048;
        var fcMac = checked(textOffset + encoded.Length);
        var chpxPage = Align(fcMac);
        var chpxPages = CreateCharacterPages(formatRuns, combined, textOffset,
            fontIndexes, revisionAuthorIndexes, hasHeaderStories ? text.Length : null, pictureOffsets,
            new HashSet<int>(floating.Select(x => x.Cp)));
        var paragraphEnds = new List<int>();
        for (var i = 0; i < combined.Length; i++)
            if (combined[i] is '\r' or '\f' or '\u0007')
                paragraphEnds.Add(checked(textOffset + (i + 1) * 2));
        var papxPages = CreateParagraphPages(paragraphEnds, paragraphStyleRuns,
            textOffset, dataStream, document.GrowAutofit);
        var pageEnd = checked(chpxPage + PageSize * (chpxPages.Count + papxPages.Count));
        var sectionBytes = document.Sections.Select(x =>
            (x.Model?.Formatting ?? x.Formatting ?? new DocSectionFormatting())
                .WithWriterDefaults().Encode()).ToArray();
        var sepxOffsets = new int[sectionBytes.Length];
        var wordLength = pageEnd;
        for (var i = 0; i < sectionBytes.Length; i++)
        {
            sepxOffsets[i] = sectionBytes[i].Length == 0 ? -1 : wordLength;
            if (sepxOffsets[i] >= 0) wordLength = checked(wordLength + 2 + sectionBytes[i].Length);
        }
        var floatingContent = floating.Count > 0
            ? DocFloatingPictureWriter.Write(floating, text.Length, combined.Length,
                wordLength) : default;
        if (floating.Count > 0) wordLength = checked(wordLength + floatingContent.Blips.Length);
        var word = new byte[wordLength];
        encoded.CopyTo(word, textOffset);

        for (var i = 0; i < chpxPages.Count; i++)
            chpxPages[i].Page.CopyTo(word, chpxPage + i * PageSize);

        for (var i = 0; i < papxPages.Count; i++)
            papxPages[i].Page.CopyTo(word, chpxPage + PageSize * (i + chpxPages.Count));
        for (var i = 0; i < sectionBytes.Length; i++)
            if (sepxOffsets[i] >= 0)
            {
                BinaryPrimitives.WriteInt16LittleEndian(word.AsSpan(sepxOffsets[i]),
                    checked((short)sectionBytes[i].Length));
                sectionBytes[i].CopyTo(word, sepxOffsets[i] + 2);
            }
        if (floating.Count > 0)
            floatingContent.Blips.CopyTo(word, wordLength - floatingContent.Blips.Length);

        var fib = new DocFibWriter(word, textOffset, fcMac, text.Length, header.Length);
        using var table = new MemoryStream();
        fib.AddTableBlock(table, 1, DocDefaultStructures.CreateStyleSheet(
            document.Styles ?? Array.Empty<DocStyleDefinition>(), fontIndexes,
            defaultFormatting, document.DefaultParagraphFormatting));
        if (fontNames.Count > 0)
            fib.AddTableBlock(table, 15, DocFontTable.Write(fontNames,
                document.FontMetadata));
        if (formatRuns.Any(x => x.Formatting.DeletedRevision == true ||
            x.Formatting.InsertedRevision == true))
            fib.AddTableBlock(table, 51, CreateRevisionAuthors(revisionAuthorNames));
        var dop = new byte[694]; // Word 2007 DOP, including compatibility options.
        if (!document.BalanceSingleByteDoubleByteWidth)
        {
            dop[9] |= 0x80; // DopBase.copts60.fDntBlnSbDbWid.
            dop[85] |= 0x80; // Dop95.copts80 repeats the legacy flags.
            dop[509] |= 0x80; // Dop2000.copts.copts80 repeats them for Word 2002.
        }
        if (paragraphStyleRuns.Any(x => x.Formatting is { Kinsoku: not null } or
            { WordWrap: not null }) || document.Styles?.Any(x =>
            x.ParagraphFormatting is { Kinsoku: not null } or { WordWrap: not null }) == true ||
            document.Sections.Any(x => x.Formatting?.GridMode is 1 or 3))
            dop[513] |= 0x20; // Dop2000.copts.fApplyBreakingRules.
        if (document.GrowAutofit)
            dop[514] |= 0x02; // Dop2000.copts.fGrowAutoFit.
        if (document.EvenAndOddHeaders) dop[0] |= 1; // DopBase.fFacingPages.
        if (document.AutoHyphenation) dop[5] |= 0x10; // DopBase.fAutoHyphen.
        if (document.HyphenateCaps) dop[5] |= 0x08; // DopBase.fHyphCapitals.
        if (document.MirrorMargins) dop[6] |= 0x20; // DopBase.fMirrorMargins.
        if (document.GutterAtTop) dop[83] |= 0x80; // DopBase.iGutterPos.
        if (document.DefaultParagraphFormatting is { BeforeLines: not null } or
            { AfterLines: not null } or { LeftChars: not null } or
            { RightChars: not null } or { FirstLineChars: not null } ||
            document.Styles?.Any(x => x.ParagraphFormatting is
                { BeforeLines: not null } or { AfterLines: not null } or
                { LeftChars: not null } or { RightChars: not null } or
                { FirstLineChars: not null }) == true ||
            paragraphStyleRuns.Any(x => x.Formatting is
                { BeforeLines: not null } or { AfterLines: not null } or
                { LeftChars: not null } or { RightChars: not null } or
                { FirstLineChars: not null }))
            dop[507] |= 0x40; // Dop2000.fCharLineUnits.
        BinaryPrimitives.WriteInt16LittleEndian(dop.AsSpan(10),
            document.DefaultTabStopTwips); // DopBase.dxaTab.
        BinaryPrimitives.WriteInt16LittleEndian(dop.AsSpan(14),
            document.HyphenationZoneTwips); // DopBase.dxaHotZ.
        BinaryPrimitives.WriteInt16LittleEndian(dop.AsSpan(16),
            document.ConsecutiveHyphenLimit); // DopBase.cConsecHypLim.
        fib.AddTableBlock(table, 31, dop);
        var tableGroups = paragraphStyleRuns.Where(x =>
                x.Formatting?.TableGroupId is > 0)
            .Select(x => x.Formatting!.TableGroupId!.Value).Distinct().OrderBy(x => x)
            .ToArray();
        if (tableGroups.Length > 0)
        {
            if (tableGroups.Length > ushort.MaxValue || paragraphStyleRuns.Any(x =>
                x.Formatting?.TableGroupId is uint id &&
                x.Formatting.ParagraphGroupId != id))
                throw new InvalidDataException("DOC table group references are invalid.");
            var pageProperties = new byte[checked(2 + tableGroups.Length * 14)];
            BinaryPrimitives.WriteUInt16LittleEndian(pageProperties,
                checked((ushort)tableGroups.Length));
            for (var i = 0; i < tableGroups.Length; i++)
                BinaryPrimitives.WriteUInt32LittleEndian(
                    pageProperties.AsSpan(2 + i * 14), tableGroups[i]);
            fib.AddTableBlock(table, 109, pageProperties);
        }
        // SttbfAssoc is mandatory even when all document association strings
        // are empty. It contains eighteen Unicode entries with no extra data.
        var associations = new byte[6 + 18 * 2];
        BinaryPrimitives.WriteUInt16LittleEndian(associations, 0xFFFF);
        BinaryPrimitives.WriteUInt16LittleEndian(associations.AsSpan(2), 18);
        fib.AddTableBlock(table, 32, associations);
        fib.AddTableBlock(table, 6, CreateSectionTable(document.Sections, sepxOffsets));
        if (hasHeaderStories)
        {
            var headerPlc = new byte[(storyEnds.Count + 1) * 4];
            for (var i = 0; i < storyEnds.Count; i++) U32(headerPlc, i * 4, storyEnds[i]);
            U32(headerPlc, storyEnds.Count * 4, header.Length + 2);
            fib.AddTableBlock(table, 11, headerPlc);
        }
        if (CreateFieldTable(text, mainStory.FieldMarks) is byte[] mainFields)
            fib.AddTableBlock(table, 16, mainFields);
        if (hasHeaderStories && CreateFieldTable(header,
            capturedHeaderFields ? headerFieldMarks : null) is byte[] headerFields)
            fib.AddTableBlock(table, 17, headerFields);
        if (bookmarks.Count > 0)
        {
            var blocks = CreateBookmarks(bookmarks, combined.Length);
            fib.AddTableBlock(table, 21, blocks.Names);
            fib.AddTableBlock(table, 22, blocks.Starts);
            fib.AddTableBlock(table, 23, blocks.Ends);
        }
        if (document.Lists is { Definitions.Count: > 0 } lists)
        {
            var blocks = DocListWriter.Write(lists, fontIndexes);
            fib.AddTableBlock(table, 73, blocks.FixedDefinitions);
            table.Write(blocks.Levels, 0, blocks.Levels.Length);
            fib.AddTableBlock(table, 74, blocks.Overrides);
        }
        if (floating.Count > 0)
        {
            if (floatingContent.MainAnchorPlc.Length > 0)
                fib.AddTableBlock(table, 40, floatingContent.MainAnchorPlc);
            if (floatingContent.HeaderAnchorPlc.Length > 0)
                fib.AddTableBlock(table, 41, floatingContent.HeaderAnchorPlc);
            fib.AddTableBlock(table, 50, floatingContent.DrawingContent);
        }

        var chpx = new byte[(chpxPages.Count * 2 + 1) * 4];
        U32(chpx, 0, textOffset);
        for (var i = 0; i < chpxPages.Count; i++)
        {
            U32(chpx, (i + 1) * 4, chpxPages[i].EndFc);
            U32(chpx, (chpxPages.Count + 1 + i) * 4, chpxPage / PageSize + i);
        }
        fib.AddTableBlock(table, 12, chpx);

        var papx = new byte[(papxPages.Count * 2 + 1) * 4];
        U32(papx, 0, textOffset);
        for (var page = 0; page < papxPages.Count; page++)
        {
            U32(papx, (page + 1) * 4, papxPages[page].EndFc);
            U32(papx, (papxPages.Count + 1 + page) * 4,
                chpxPage / PageSize + chpxPages.Count + page);
        }
        fib.AddTableBlock(table, 13, papx);

        var clx = new byte[21];
        clx[0] = 2;
        U32(clx, 1, 16);
        U32(clx, 9, combined.Length);
        U32(clx, 15, textOffset);
        fib.AddTableBlock(table, 33, clx);

        using var storage = RootStorage.Create(output, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen);
        WriteStream(storage, "WordDocument", word);
        WriteStream(storage, "1Table", table.ToArray());
        if (document.Title != null || document.Subject != null || document.Author != null ||
            document.Keywords != null || document.Comments != null ||
            document.RevisionNumber != null ||
            document.LastAuthor != null || document.PageCount != null ||
            document.WordCount != null || document.CharacterCount != null)
            WriteStream(storage, "\u0005SummaryInformation",
                DocSummaryInformation.Create(document.Title, document.Subject,
                    document.Author, document.Keywords, document.Comments,
                    document.LastAuthor, document.PageCount,
                    document.WordCount, document.CharacterCount,
                    document.RevisionNumber));
        if (dataStream.Length > 0)
            WriteStream(storage, "Data", dataStream.ToArray());
    }

    private static List<(byte[] Page, int EndFc)> CreateParagraphPages(
        IReadOnlyList<int> paragraphEnds,
        IReadOnlyList<DocPlainTextParagraphStyleRun> styleRuns, int textOffset,
        MemoryStream dataStream, bool growAutofit)
    {
        var byEnd = styleRuns.ToDictionary(x => x.End,
            x => ParagraphPropertyBlock(x, dataStream, growAutofit));
        var pages = new List<(byte[] Page, int EndFc)>();
        for (var start = 0; start < paragraphEnds.Count;)
        {
            var count = 0;
            var propertyBytes = 0;
            while (start + count < paragraphEnds.Count && count < 29)
            {
                var endCp = (paragraphEnds[start + count] - textOffset) / 2;
                var nextBytes = propertyBytes + (byEnd.TryGetValue(endCp, out var block)
                    ? block.Length : 0);
                var nextCount = count + 1;
                if ((nextCount + 1) * 4 + nextCount * 13 + nextBytes >= 510)
                    break;
                count = nextCount;
                propertyBytes = nextBytes;
            }
            if (count == 0) throw new InvalidDataException("A paragraph FKP cannot fit a run.");
            var page = new byte[PageSize];
            U32(page, 0, start == 0 ? textOffset : paragraphEnds[start - 1]);
            var recordStart = (count + 1) * 4;
            var propertyCursor = 510;
            for (var i = 0; i < count; i++)
            {
                var endFc = paragraphEnds[start + i];
                U32(page, (i + 1) * 4, endFc);
                var endCp = (endFc - textOffset) / 2;
                if (!byEnd.TryGetValue(endCp, out var block)) continue;
                propertyCursor -= block.Length;
                block.CopyTo(page, propertyCursor);
                page[recordStart + i * 13] = checked((byte)(propertyCursor / 2));
            }
            page[511] = checked((byte)count);
            pages.Add((page, paragraphEnds[start + count - 1]));
            start += count;
        }
        return pages;
    }

    private static byte[] ParagraphPropertyBlock(DocPlainTextParagraphStyleRun run,
        MemoryStream dataStream, bool growAutofit)
    {
        var formatting = run.Formatting ?? DocParagraphFormatting.Empty;
        var compatibility = formatting.Encode();
        byte[] properties;
        if ((formatting.TableAutoFit == true || UsesModernStyledLeftBorder(formatting)) &&
            formatting.TableCellEdges is { Count: > 1 } edges)
        {
            // Keep the PAPX TDefTable for older readers and give modern Word
            // a TInsert/TDxaCol row definition through sprmPTableProps.
            var offset = checked((uint)dataStream.Position);
            var modern = ModernRowProperties(formatting, edges, growAutofit);
            dataStream.Write(modern);
            properties = new byte[6 + compatibility.Length];
            BinaryPrimitives.WriteUInt16LittleEndian(properties, 0x646B);
            BinaryPrimitives.WriteUInt32LittleEndian(properties.AsSpan(2), offset);
            compatibility.CopyTo(properties, 6);
        }
        else properties = compatibility;
        var payloadLength = 2 + properties.Length;
        var blockLength = 1 + payloadLength + (payloadLength % 2 == 0 ? 1 : 0);
        if (blockLength > 510) throw new InvalidDataException("A paragraph PAPX is too large.");
        var block = new byte[blockLength];
        block[0] = checked((byte)(blockLength / 2));
        BinaryPrimitives.WriteUInt16LittleEndian(block.AsSpan(1), checked((ushort)run.StyleIndex));
        properties.CopyTo(block, 3);
        return block;
    }

    // Word positions a fixed styled row with a visible left cell border
    // through a modern table-property block and a half-border origin.
    // The zero first edge excludes inherited outer-border placement.
    private static bool UsesModernStyledLeftBorder(DocParagraphFormatting formatting) =>
        formatting.TableAutoFit == false && formatting.TableStyleIndex != null &&
        formatting.TableIndentTwips == null && formatting.TableBorders?.Left == null &&
        formatting.TableCellEdges is { Count: > 1 } edges && edges[0] == 0 &&
        formatting.TableCellBorders?.FirstOrDefault()?.Left is { WidthEighthPoints: > 0 };

    private static byte[] ModernRowProperties(DocParagraphFormatting formatting,
        IReadOnlyList<short> edges, bool growAutofit)
    {
        using var modifiers = new MemoryStream();
        static void WriteI16(Stream stream, short value)
        {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
        }
        void WriteByteProperty(ushort code, byte value)
        {
            WriteI16(modifiers, unchecked((short)code));
            modifiers.WriteByte(value);
        }
        void WriteWordProperty(ushort code, short value)
        {
            WriteI16(modifiers, unchecked((short)code));
            WriteI16(modifiers, value);
        }
        WriteByteProperty(0x2416, 1);
        WriteByteProperty(0x2417, 1);
        WriteI16(modifiers, 0x6649);
        var depth = new byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(depth, formatting.TableDepth ?? 1);
        modifiers.Write(depth);
        var styledLeftBorder = UsesModernStyledLeftBorder(formatting)
            ? formatting.TableCellBorders![0]!.Left : null;
        WriteWordProperty(0x9601, styledLeftBorder is { } leftBorder
            ? checked((short)-((leftBorder.WidthEighthPoints * 5 + 2) / 4))
            : formatting.TableRowOriginTwips ?? 0);
        if (growAutofit)
        {
            var rowMargins = formatting.TableDefaultCellMargins;
            WriteWordProperty(0x9602, checked((short)(((rowMargins?.Left ?? 108) +
                (rowMargins?.Right ?? 108)) / 2)));
        }

        WriteI16(modifiers, 0x7621);
        modifiers.WriteByte(0);
        modifiers.WriteByte(checked((byte)(edges.Count - 1)));
        WriteI16(modifiers, 360);
        for (var i = 0; i < edges.Count - 1; i++)
        {
            WriteI16(modifiers, 0x7623);
            modifiers.WriteByte(checked((byte)i));
            modifiers.WriteByte(checked((byte)(i + 1)));
            WriteI16(modifiers, checked((short)(edges[i + 1] - edges[i])));
        }
        if (formatting.TableCellNoWraps is { } noWraps)
            // Word uses the modern row block for auto-fit tables; the
            // compatibility PAPX no-wrap operand alone is not applied there.
            for (var i = 0; i < noWraps.Count; i++)
                if (noWraps[i] is bool noWrap)
                {
                    WriteI16(modifiers, unchecked((short)0xD639));
                    modifiers.WriteByte(3);
                    modifiers.WriteByte(checked((byte)i));
                    modifiers.WriteByte(checked((byte)(i + 1)));
                    modifiers.WriteByte(noWrap ? (byte)1 : (byte)0);
                }
        if (formatting.TableCellTextFlows is { } textFlows)
            for (var i = 0; i < Math.Min(textFlows.Count, edges.Count - 1); i++)
                if (textFlows[i] is ushort flow)
                {
                    if (flow is not (0 or 1 or 3 or 4 or 5))
                        throw new InvalidDataException("A table cell text flow is invalid.");
                    WriteI16(modifiers, 0x7629);
                    modifiers.WriteByte(checked((byte)i));
                    modifiers.WriteByte(checked((byte)(i + 1)));
                    WriteI16(modifiers, checked((short)flow));
                }
        if (formatting.TableCellHideMarks is { } hideMarks)
            for (var i = 0; i < Math.Min(hideMarks.Count, edges.Count - 1); i++)
                if (hideMarks[i] is bool hideMark)
                {
                    WriteI16(modifiers, unchecked((short)0xD642));
                    modifiers.WriteByte(3);
                    modifiers.WriteByte(checked((byte)i));
                    modifiers.WriteByte(checked((byte)(i + 1)));
                    modifiers.WriteByte(hideMark ? (byte)1 : (byte)0);
                }
        if (formatting.TableCellHorizontalMerges is { } merges)
            for (var i = 0; i < merges.Count; i++)
            {
                if (merges[i] is not (2 or 3)) continue;
                var end = i + 1;
                while (end < merges.Count && merges[end] == 1) end++;
                if (end == i + 1) continue;
                WriteI16(modifiers, 0x5624);
                modifiers.WriteByte(checked((byte)i));
                modifiers.WriteByte(checked((byte)end));
                i = end - 1;
            }
        if (formatting.TableStyleIndex is ushort style)
            WriteWordProperty(0x563A, checked((short)style));
        if (formatting.TablePreferredWidth is { } preferred)
        {
            WriteI16(modifiers, unchecked((short)0xF614));
            modifiers.WriteByte(preferred.Unit);
            WriteI16(modifiers, checked((short)preferred.Value));
        }
        WriteByteProperty(0x3615, formatting.TableAutoFit == true ? (byte)1 : (byte)0);
        if (formatting.TableCellPreferredWidths is { } widths)
            for (var i = 0; i < Math.Min(widths.Count, edges.Count - 1); i++)
                if (widths[i] is { } width)
                {
                    WriteI16(modifiers, unchecked((short)0xD635));
                    modifiers.WriteByte(5);
                    modifiers.WriteByte(checked((byte)i));
                    modifiers.WriteByte(checked((byte)(i + 1)));
                    modifiers.WriteByte(width.Unit);
                    WriteI16(modifiers, checked((short)width.Value));
                }
        if (formatting.TableCellSpacingTwips is ushort cellSpacing)
        {
            WriteI16(modifiers, unchecked((short)0xD633));
            modifiers.WriteByte(6);
            modifiers.WriteByte(0);
            modifiers.WriteByte(1);
            modifiers.WriteByte(0x0F);
            modifiers.WriteByte(3);
            WriteI16(modifiers, checked((short)cellSpacing));
        }
        if (styledLeftBorder != null && formatting.TableCellBorders is { } directBorders)
            for (var i = 0; i < directBorders.Count; i++)
            {
                var borders = directBorders[i];
                WriteBorder(borders?.Top, 1); WriteBorder(borders?.Left, 2);
                WriteBorder(borders?.Bottom, 4); WriteBorder(borders?.Right, 8);
                WriteBorder(borders?.TopLeftToBottomRight, 0x10);
                WriteBorder(borders?.TopRightToBottomLeft, 0x20);
                void WriteBorder(DocParagraphBorder? border, byte side)
                {
                    if (border == null) return;
                    WriteI16(modifiers, unchecked((short)0xD62F));
                    modifiers.WriteByte(11); modifiers.WriteByte(checked((byte)i));
                    modifiers.WriteByte(checked((byte)(i + 1))); modifiers.WriteByte(side);
                    modifiers.Write(border.EncodeRaw());
                }
            }
        var defaultMargins = formatting.TableDefaultCellMargins ?? new DocCellMargins();
        WriteMargins(0xD634, 0, defaultMargins with
        {
            Left = defaultMargins.Left ?? 108,
            Right = defaultMargins.Right ?? 108
        });
        if (formatting.TableCellMargins is { } cellMargins)
            for (var i = 0; i < Math.Min(cellMargins.Count, edges.Count - 1); i++)
                if (cellMargins[i] is { } cellMargin)
                    WriteMargins(0xD632, i, cellMargin);
        if (formatting.TableBackgroundShading is { } backgroundShading)
        {
            var shadingProperty = (DocParagraphFormatting.Empty with
                { TableBackgroundShading = backgroundShading }).Encode();
            modifiers.Write(shadingProperty, 0, shadingProperty.Length);
        }
        if (formatting.TableIndentTwips is short tableIndent)
        {
            WriteI16(modifiers, unchecked((short)0xF661));
            modifiers.WriteByte(3);
            WriteI16(modifiers, tableIndent);
        }
        if (formatting.TableGroupId is uint groupId)
        {
            WriteI16(modifiers, 0x7479);
            var group = new byte[4];
            BinaryPrimitives.WriteUInt32LittleEndian(group, groupId);
            modifiers.Write(group);
        }
        var data = modifiers.ToArray();
        var result = new byte[2 + data.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(result,
            checked((ushort)data.Length));
        data.CopyTo(result, 2);
        return result;

        void WriteMargins(ushort code, int index, DocCellMargins margins)
        {
            WriteMargin(1, margins.Top);
            WriteMargin(2, margins.Left);
            WriteMargin(4, margins.Bottom);
            WriteMargin(8, margins.Right);

            void WriteMargin(byte side, ushort? value)
            {
                if (value == null) return;
                if (value > 31680)
                    throw new InvalidDataException("A DOC cell margin exceeds 22 inches.");
                WriteI16(modifiers, unchecked((short)code));
                modifiers.WriteByte(6);
                modifiers.WriteByte(checked((byte)index));
                modifiers.WriteByte(checked((byte)(index + 1)));
                modifiers.WriteByte(side);
                modifiers.WriteByte(3);
                WriteI16(modifiers, checked((short)value.Value));
            }
        }
    }


    private static List<(byte[] Page, int EndFc)> CreateCharacterPages(
        IReadOnlyList<DocPlainTextFormatRun> runs, string characters, int textOffset,
        IReadOnlyDictionary<string, int> fontIndexes,
        IReadOnlyDictionary<string, int> revisionAuthorIndexes,
        int? storyBoundary = null,
        IReadOnlyDictionary<int, int>? pictureOffsets = null,
        ISet<int>? floatingCps = null)
    {
        var characterCount = characters.Length;
        var segments = new List<DocPlainTextFormatRun>();
        var cursor = 0;
        foreach (var run in runs.OrderBy(x => x.Start))
        {
            if (run.Start < cursor || run.End <= run.Start || run.End > characterCount)
                throw new InvalidDataException("Character formatting runs overlap or exceed the text.");
            if (run.Start > cursor)
                segments.Add(new DocPlainTextFormatRun(cursor, run.Start, DocCharacterFormatting.Empty));
            segments.Add(run);
            cursor = run.End;
        }
        if (cursor < characterCount)
            segments.Add(new DocPlainTextFormatRun(cursor, characterCount,
                DocCharacterFormatting.Empty));
        if (storyBoundary is int boundary)
            for (var i = 0; i < segments.Count; i++)
                if (segments[i].Start < boundary && segments[i].End > boundary)
                {
                    var segment = segments[i];
                    segments[i] = segment with { End = boundary };
                    segments.Insert(i + 1, segment with { Start = boundary });
                    break;
                }
        for (var cp = 0; cp < characterCount; cp++)
        {
            var isPicture = pictureOffsets != null && pictureOffsets.ContainsKey(cp);
            if (!isPicture && floatingCps?.Contains(cp) != true &&
                characters[cp] is not ('\u0013' or '\u0014' or '\u0015')) continue;
            var index = segments.FindIndex(x => x.Start <= cp && cp < x.End);
            if (index < 0) throw new InvalidDataException("A field mark has no character run.");
            var segment = segments[index];
            segments.RemoveAt(index);
            if (segment.Start < cp)
                segments.Insert(index++, segment with { End = cp });
            segments.Insert(index++, segment with
            {
                Start = cp, End = cp + 1,
                Formatting = segment.Formatting with
                {
                    Special = true,
                    PictureDataOffset = isPicture ? pictureOffsets![cp] :
                        floatingCps?.Contains(cp) == true ? 0 : null
                }
            });
            if (cp + 1 < segment.End)
                segments.Insert(index, segment with { Start = cp + 1 });
        }
        var pages = new List<(byte[] Page, int EndFc)>();
        for (var start = 0; start < segments.Count;)
        {
            var count = 0;
            var nextPropertyCursor = 510;
            while (start + count < segments.Count && count < 20)
            {
                var properties = EncodeCharacterProperties(segments[start + count].Formatting,
                    fontIndexes, revisionAuthorIndexes);
                if (properties.Length > byte.MaxValue)
                    throw new InvalidDataException("A CHPX property list is too large.");
                var candidateCursor = properties.Length == 0 ? nextPropertyCursor :
                    (nextPropertyCursor - properties.Length - 1) & ~1;
                var candidateCount = count + 1;
                if (candidateCursor <= (candidateCount + 1) * 4 + candidateCount)
                    break;
                nextPropertyCursor = candidateCursor;
                count = candidateCount;
            }
            if (count == 0)
                throw new InvalidDataException("A character formatting page cannot fit a run.");
            var page = new byte[PageSize];
            var propertyCursor = 510;
            U32(page, 0, checked(textOffset + segments[start].Start * 2));
            for (var i = 0; i < count; i++)
            {
                var run = segments[start + i];
                U32(page, (i + 1) * 4, checked(textOffset + run.End * 2));
                var properties = EncodeCharacterProperties(run.Formatting, fontIndexes,
                    revisionAuthorIndexes);
                if (properties.Length == 0) continue;
                var length = properties.Length + 1;
                propertyCursor = (propertyCursor - length) & ~1;
                if (propertyCursor <= (count + 1) * 4 + count)
                    throw new InvalidDataException("A character formatting page is full.");
                page[propertyCursor] = checked((byte)properties.Length);
                properties.CopyTo(page, propertyCursor + 1);
                page[(count + 1) * 4 + i] = checked((byte)(propertyCursor / 2));
            }
            page[511] = checked((byte)count);
            pages.Add((page, checked(textOffset + segments[start + count - 1].End * 2)));
            start += count;
        }
        return pages;
    }

    internal static byte[] EncodeCharacterProperties(DocCharacterFormatting formatting,
        IReadOnlyDictionary<string, int>? fontIndexes = null,
        IReadOnlyDictionary<string, int>? revisionAuthorIndexes = null)
    {
        var bytes = new List<byte>(14);
        void WriteFont(string? name, byte opcode)
        {
            if (name == null || fontIndexes == null ||
                !fontIndexes.TryGetValue(name, out var index)) return;
            bytes.Add(opcode); bytes.Add(0x4A);
            bytes.Add((byte)index); bytes.Add((byte)(index >> 8));
        }
        if (formatting.SymbolCharacter is ushort symbol)
        {
            if (formatting.SymbolFontName == null || fontIndexes == null ||
                !fontIndexes.TryGetValue(formatting.SymbolFontName, out var symbolFont))
                throw new InvalidDataException("A DOC symbol needs a font table entry.");
            bytes.AddRange([0x09, 0x6A, (byte)symbolFont, (byte)(symbolFont >> 8),
                (byte)symbol, (byte)(symbol >> 8)]);
        }
        if (formatting.CharacterStyleIndex is int styleIndex)
        {
            bytes.AddRange([0x30, 0x4A]);
            bytes.Add((byte)styleIndex);
            bytes.Add((byte)(styleIndex >> 8));
        }
        if (formatting.DeletedRevision is bool deletedRevision)
            bytes.AddRange([0x00, 0x08, deletedRevision ? (byte)1 : (byte)0]);
        if (formatting.InsertedRevision is bool insertedRevision)
            bytes.AddRange([0x01, 0x08, insertedRevision ? (byte)1 : (byte)0]);
        void WriteRevisionDetails(bool? marked, string? author, DateTime? at,
            byte authorOpcode, byte dateOpcode)
        {
            if (marked != true) return;
            if (author != null && revisionAuthorIndexes != null)
            {
                var index = revisionAuthorIndexes.TryGetValue(author, out var found)
                    ? found : 0;
                bytes.AddRange([authorOpcode, 0x48, (byte)index, (byte)(index >> 8)]);
            }
            if (at is DateTime date)
            {
                var packed = DocRevisionDate.Encode(date);
                bytes.AddRange([dateOpcode, 0x68, (byte)packed, (byte)(packed >> 8),
                    (byte)(packed >> 16), (byte)(packed >> 24)]);
            }
        }
        WriteRevisionDetails(formatting.InsertedRevision,
            formatting.InsertedRevisionAuthor, formatting.InsertedRevisionAt, 0x04, 0x05);
        WriteRevisionDetails(formatting.DeletedRevision,
            formatting.DeletedRevisionAuthor, formatting.DeletedRevisionAt, 0x63, 0x64);
        if (formatting.Bold is bool bold)
            bytes.AddRange([0x35, 0x08, bold ? (byte)1 : (byte)0]);
        if (formatting.Italic is bool italic)
            bytes.AddRange([0x36, 0x08, italic ? (byte)1 : (byte)0]);
        if (formatting.ComplexScriptBold is bool complexBold)
            bytes.AddRange([0x5C, 0x08, complexBold ? (byte)1 : (byte)0]);
        if (formatting.ComplexScriptItalic is bool complexItalic)
            bytes.AddRange([0x5D, 0x08, complexItalic ? (byte)1 : (byte)0]);
        if (formatting.RightToLeftText is bool rightToLeft)
            bytes.AddRange([0x5A, 0x08, rightToLeft ? (byte)1 : (byte)0]);
        if (formatting.ForceComplexScript is bool forceComplexScript)
            bytes.AddRange([0x82, 0x08, forceComplexScript ? (byte)1 : (byte)0]);
        if (formatting.Strike is bool strike)
            bytes.AddRange([0x37, 0x08, strike ? (byte)1 : (byte)0]);
        if (formatting.DoubleStrike is bool doubleStrike)
            bytes.AddRange([0x53, 0x2A, doubleStrike ? (byte)1 : (byte)0]);
        if (formatting.CharacterScalePercent is ushort characterScale)
        {
            if (characterScale is < 1 or > 600)
                throw new InvalidDataException("A character scale is outside 1–600 percent.");
            bytes.AddRange([0x52, 0x48]);
            bytes.AddRange(BitConverter.GetBytes(characterScale));
        }
        if (formatting.BaselineOffsetHalfPoints is short baselineOffset)
        {
            if (baselineOffset is < -3168 or > 3168)
                throw new InvalidDataException("A baseline offset is outside ±3168 half-points.");
            bytes.AddRange([0x45, 0x48]);
            bytes.AddRange(BitConverter.GetBytes(baselineOffset));
        }
        if (formatting.Outline is bool outline)
            bytes.AddRange([0x38, 0x08, outline ? (byte)1 : (byte)0]);
        if (formatting.Shadow is bool shadow)
            bytes.AddRange([0x39, 0x08, shadow ? (byte)1 : (byte)0]);
        if (formatting.Emboss is bool emboss)
            bytes.AddRange([0x58, 0x08, emboss ? (byte)1 : (byte)0]);
        if (formatting.Imprint is bool imprint)
            bytes.AddRange([0x54, 0x08, imprint ? (byte)1 : (byte)0]);
        if (formatting.SmallCaps is bool smallCaps)
            bytes.AddRange([0x3A, 0x08, smallCaps ? (byte)1 : (byte)0]);
        if (formatting.Caps is bool caps)
            bytes.AddRange([0x3B, 0x08, caps ? (byte)1 : (byte)0]);
        if (formatting.Hidden is bool hidden)
            bytes.AddRange([0x3C, 0x08, hidden ? (byte)1 : (byte)0]);
        if (formatting.SnapToGrid is bool snapToGrid)
            bytes.AddRange([0x68, 0x08, snapToGrid ? (byte)1 : (byte)0]);
        if (formatting.Special is bool special)
            bytes.AddRange([0x55, 0x08, special ? (byte)1 : (byte)0]);
        if (formatting.PictureDataOffset is int pictureOffset)
        {
            bytes.AddRange([0x03, 0x6A]);
            bytes.AddRange(BitConverter.GetBytes(pictureOffset));
        }
        if (formatting.UnderlineCode is byte underline)
            bytes.AddRange([0x3E, 0x2A, underline]);
        if (formatting.FitText is { } fitText)
        {
            bytes.AddRange([0x76, 0xCA, 8]);
            bytes.AddRange(BitConverter.GetBytes(fitText.WidthTwips));
            bytes.AddRange(BitConverter.GetBytes(fitText.Id));
        }
        if (formatting.EmphasisMarkCode is byte emphasis)
        {
            if (emphasis > 4)
                throw new InvalidDataException("A DOC emphasis mark is invalid.");
            bytes.AddRange([0x34, 0x2A, emphasis]);
        }
        if (formatting.ScriptCode is byte script)
            bytes.AddRange([0x48, 0x2A, script]);
        if (formatting.SizeHalfPoints is ushort size)
        {
            bytes.AddRange([0x43, 0x4A]);
            bytes.Add((byte)size);
            bytes.Add((byte)(size >> 8));
        }
        if (formatting.ComplexScriptSizeHalfPoints is ushort complexSize)
        {
            bytes.AddRange([0x61, 0x4A]);
            bytes.Add((byte)complexSize);
            bytes.Add((byte)(complexSize >> 8));
        }
        WriteFont(formatting.AsciiFontName, 0x4F);
        WriteFont(formatting.EastAsiaFontName, 0x50);
        WriteFont(formatting.HighAnsiFontName, 0x51);
        WriteFont(formatting.ComplexScriptFontName, 0x5E);
        if (formatting.HighlightCode is byte highlight)
            bytes.AddRange([0x0C, 0x2A, highlight]);
        if (formatting.ColorRef is uint color)
        {
            bytes.AddRange([0x70, 0x68]);
            bytes.Add((byte)color);
            bytes.Add((byte)(color >> 8));
            bytes.Add((byte)(color >> 16));
            bytes.Add((byte)(color >> 24));
        }
        if (formatting.UnderlineColorRef is uint underlineColor)
        {
            bytes.AddRange([0x77, 0x68]);
            bytes.Add((byte)underlineColor);
            bytes.Add((byte)(underlineColor >> 8));
            bytes.Add((byte)(underlineColor >> 16));
            bytes.Add((byte)(underlineColor >> 24));
        }
        if (formatting.Border is { } border)
        {
            if (border.Type == 0)
            {
                bytes.AddRange([0x65, 0x68, 0xFF, 0xFF, 0xFF, 0xFF]);
                bytes.AddRange([0x72, 0xCA, 8, 0xFF, 0xFF, 0xFF, 0xFF,
                    0xFF, 0xFF, 0xFF, 0xFF]);
            }
            else
            {
                if (border.Encode80() is { } legacy)
                {
                    bytes.AddRange([0x65, 0x68]);
                    bytes.AddRange(legacy);
                }
                bytes.AddRange([0x72, 0xCA]);
                bytes.AddRange(border.Encode());
            }
        }
        if (formatting.Shading is { } shading)
        {
            var nil = shading.Pattern == ushort.MaxValue;
            static byte? LegacyShadingColor(uint? value) => value switch
            {
                null or 0xFF000000u => 0,
                0x000000u => 1,
                0xFF0000u => 2,
                0xFFFF00u => 3,
                0x00FF00u => 4,
                0xFF00FFu => 5,
                0x0000FFu => 6,
                0x00FFFFu => 7,
                0xFFFFFFu => 8,
                0x800000u => 9,
                0x808000u => 10,
                0x008000u => 11,
                0x800080u => 12,
                0x008080u => 14,
                0x808080u => 15,
                0xC0C0C0u => 16,
                _ => null
            };
            var legacyForeground = LegacyShadingColor(shading.ForegroundRgb);
            var legacyBackground = LegacyShadingColor(shading.FillRgb);
            if (!nil && legacyForeground is byte foreground &&
                legacyBackground is byte background && shading.Pattern <= 63)
            {
                var shd80 = (ushort)(foreground | (background << 5) |
                    (shading.Pattern << 10));
                bytes.AddRange([0x66, 0x48, (byte)shd80, (byte)(shd80 >> 8)]);
            }
            bytes.AddRange([0x71, 0xCA, 10]);
            void WriteShadingColor(uint? color)
            {
                var value = color ?? 0xFF000000;
                bytes.Add((byte)value);
                bytes.Add((byte)(value >> 8));
                bytes.Add((byte)(value >> 16));
                bytes.Add((byte)(value >> 24));
            }
            // ShdAuto clears inherited run shading and remains renderable by
            // Word when DOCX supplied w:shd val="nil".
            WriteShadingColor(nil ? null : shading.ForegroundRgb);
            WriteShadingColor(nil ? null : shading.FillRgb);
            bytes.Add(nil ? (byte)0 : (byte)shading.Pattern);
            bytes.Add(nil ? (byte)0 : (byte)(shading.Pattern >> 8));
        }
        void WriteWord(byte opcode, byte prefix, short value)
        {
            bytes.Add(opcode); bytes.Add(prefix);
            bytes.Add((byte)value); bytes.Add((byte)(value >> 8));
        }
        if (formatting.CharacterSpacingTwips is short spacing)
            WriteWord(0x40, 0x88, spacing);
        if (formatting.KerningThresholdHalfPoints is ushort kerning)
            WriteWord(0x4B, 0x48, unchecked((short)kerning));
        if (formatting.LanguageId is ushort language)
            WriteWord(0x73, 0x48, unchecked((short)language));
        if (formatting.EastAsiaLanguageId is ushort eastAsiaLanguage)
            WriteWord(0x74, 0x48, unchecked((short)eastAsiaLanguage));
        if (formatting.ComplexScriptLanguageId is ushort complexLanguage)
            WriteWord(0x5F, 0x48, unchecked((short)complexLanguage));
        return bytes.ToArray();
    }

    private static byte[] CreateRevisionAuthors(IReadOnlyList<string> authors)
    {
        using var output = new MemoryStream();
        void Word(ushort value)
        {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
        }
        Word(0xFFFF);
        Word(checked((ushort)authors.Count));
        Word(0);
        foreach (var author in authors)
        {
            Word(checked((ushort)author.Length));
            var text = Encoding.Unicode.GetBytes(author);
            output.Write(text, 0, text.Length);
        }
        return output.ToArray();
    }

    private static byte[] CreateSectionTable(IReadOnlyList<DocPlainTextSection> sections,
        IReadOnlyList<int> sepxOffsets)
    {
        var bytes = new byte[checked(sections.Count * 16 + 4)];
        var previous = 0;
        for (var i = 0; i < sections.Count; i++)
        {
            var end = sections[i].EndCp;
            if (end <= previous) throw new InvalidDataException("DOC section positions must increase.");
            U32(bytes, (i + 1) * 4, end);
            BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan((sections.Count + 1) * 4 + i * 12 + 2), sepxOffsets[i]);
            previous = end;
        }
        return bytes;
    }

    internal static IReadOnlyList<DocStoryFieldMark> ReadFieldMarks(string story)
    {
        var marks = new List<DocStoryFieldMark>();
        var nested = new Stack<bool>();
        for (var cp = 0; cp < story.Length; cp++)
        {
            var kind = story[cp];
            if (kind == '\u0013') nested.Push(false);
            else if (kind == '\u0014')
            {
                if (nested.Count == 0 || nested.Peek())
                    throw new InvalidDataException("A DOC field separator is unmatched.");
                nested.Pop();
                nested.Push(true);
            }
            else if (kind == '\u0015')
            {
                if (nested.Count == 0)
                    throw new InvalidDataException("A DOC field end is unmatched.");
                nested.Pop();
            }
            else continue;
            var fieldType = (byte)0;
            if (kind == '\u0013')
            {
                var instructionEnd = story.IndexOfAny(['\u0014', '\u0015'], cp + 1);
                if (instructionEnd > cp)
                {
                    var instruction = story.AsSpan(cp + 1, instructionEnd - cp - 1).Trim();
                    fieldType = GetFieldType(instruction);
                }
            }
            marks.Add(new DocStoryFieldMark(cp, checked((byte)kind), fieldType));
        }
        if (nested.Count != 0)
            throw new InvalidDataException("A DOC field has no end.");
        return marks;
    }

    private static byte[]? CreateFieldTable(string story,
        IReadOnlyList<DocStoryFieldMark>? capturedMarks = null)
    {
        var marks = capturedMarks ?? ReadFieldMarks(story);
        if (story.Count(x => x is '\u0013' or '\u0014' or '\u0015') != marks.Count)
            throw new InvalidDataException("The DOC story field marks are incomplete.");
        if (marks.Count == 0) return null;
        var previous = -1;
        foreach (var mark in marks)
        {
            if (mark.Cp <= previous || mark.Cp >= story.Length ||
                story[mark.Cp] != mark.Kind)
                throw new InvalidDataException("A DOC field mark differs from its captured story.");
            previous = mark.Cp;
        }
        var bytes = new byte[checked((marks.Count + 1) * 4 + marks.Count * 2)];
        for (var i = 0; i < marks.Count; i++)
        {
            U32(bytes, i * 4, marks[i].Cp);
            bytes[(marks.Count + 1) * 4 + i * 2] = marks[i].Kind;
            bytes[(marks.Count + 1) * 4 + i * 2 + 1] = marks[i].FieldType;
        }
        U32(bytes, marks.Count * 4, story.Length);
        return bytes;
    }

    private static byte GetFieldType(ReadOnlySpan<char> instruction)
    {
        var end = instruction.IndexOfAny(" \t\r\n\\".AsSpan());
        var keyword = end < 0 ? instruction : instruction.Slice(0, end);
        if (keyword.Equals("PAGE", StringComparison.OrdinalIgnoreCase)) return 33;
        if (keyword.Equals("NUMPAGES", StringComparison.OrdinalIgnoreCase)) return 26;
        if (keyword.Equals("NUMWORDS", StringComparison.OrdinalIgnoreCase)) return 27;
        if (keyword.Equals("NUMCHARS", StringComparison.OrdinalIgnoreCase)) return 28;
        if (keyword.Equals("FILENAME", StringComparison.OrdinalIgnoreCase)) return 29;
        if (keyword.Equals("CREATEDATE", StringComparison.OrdinalIgnoreCase)) return 21;
        if (keyword.Equals("SAVEDATE", StringComparison.OrdinalIgnoreCase)) return 22;
        if (keyword.Equals("PRINTDATE", StringComparison.OrdinalIgnoreCase)) return 23;
        if (keyword.Equals("DATE", StringComparison.OrdinalIgnoreCase)) return 31;
        if (keyword.Equals("TIME", StringComparison.OrdinalIgnoreCase)) return 32;
        if (keyword.Equals("HYPERLINK", StringComparison.OrdinalIgnoreCase)) return 88;
        if (keyword.Equals("REF", StringComparison.OrdinalIgnoreCase)) return 3;
        if (keyword.Equals("PAGEREF", StringComparison.OrdinalIgnoreCase)) return 37;
        if (keyword.Equals("SEQ", StringComparison.OrdinalIgnoreCase)) return 12;
        if (keyword.Equals("STYLEREF", StringComparison.OrdinalIgnoreCase)) return 10;
        if (keyword.Equals("IF", StringComparison.OrdinalIgnoreCase)) return 7;
        if (keyword.Equals("SECTIONPAGES", StringComparison.OrdinalIgnoreCase)) return 66;
        if (keyword.Equals("SECTION", StringComparison.OrdinalIgnoreCase)) return 65;
        if (keyword.Equals("MERGEFIELD", StringComparison.OrdinalIgnoreCase)) return 59;
        if (keyword.Equals("DOCPROPERTY", StringComparison.OrdinalIgnoreCase)) return 85;
        if (keyword.Equals("LASTSAVEDBY", StringComparison.OrdinalIgnoreCase)) return 20;
        if (keyword.Equals("AUTHOR", StringComparison.OrdinalIgnoreCase)) return 17;
        if (keyword.Equals("TITLE", StringComparison.OrdinalIgnoreCase)) return 15;
        if (keyword.Equals("SUBJECT", StringComparison.OrdinalIgnoreCase)) return 16;
        if (keyword.Equals("KEYWORDS", StringComparison.OrdinalIgnoreCase)) return 18;
        if (keyword.Equals("COMMENTS", StringComparison.OrdinalIgnoreCase)) return 19;
        return 0;
    }

    private static (byte[] Names, byte[] Starts, byte[] Ends) CreateBookmarks(
        IReadOnlyList<DocPlainTextBookmark> bookmarks, int storyLength)
    {
        if (bookmarks.Count > ushort.MaxValue || bookmarks.Any(x =>
            x.StartCp < 0 || x.EndCp < x.StartCp || x.EndCp > storyLength ||
            string.IsNullOrWhiteSpace(x.Name)) ||
            bookmarks.Select(x => x.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count()
                != bookmarks.Count)
            throw new InvalidDataException("The DOC bookmarks have invalid names or positions.");
        var starts = bookmarks.OrderBy(x => x.StartCp).ThenByDescending(x => x.EndCp)
            .ToArray();
        var ends = starts.Select((bookmark, index) => (Bookmark: bookmark, StartIndex: index))
            .OrderBy(x => x.Bookmark.EndCp).ThenByDescending(x => x.Bookmark.StartCp)
            .ToArray();
        var endIndexes = new int[starts.Length];
        for (var i = 0; i < ends.Length; i++) endIndexes[ends[i].StartIndex] = i;
        using var names = new MemoryStream();
        names.WriteByte(0xFF); names.WriteByte(0xFF);
        WriteU16(names, checked((ushort)starts.Length));
        WriteU16(names, 0);
        foreach (var bookmark in starts)
        {
            WriteU16(names, checked((ushort)bookmark.Name.Length));
            var bytes = Encoding.Unicode.GetBytes(bookmark.Name);
            names.Write(bytes, 0, bytes.Length);
        }
        var startPlc = new byte[checked(starts.Length * 8 + 4)];
        var endPlc = new byte[checked((ends.Length + 1) * 4)];
        for (var i = 0; i < starts.Length; i++)
        {
            U32(startPlc, i * 4, starts[i].StartCp);
            BinaryPrimitives.WriteUInt16LittleEndian(startPlc.AsSpan((starts.Length + 1) * 4 + i * 4),
                checked((ushort)endIndexes[i]));
        }
        U32(startPlc, starts.Length * 4, storyLength);
        for (var i = 0; i < ends.Length; i++)
            U32(endPlc, i * 4, ends[i].Bookmark.EndCp);
        U32(endPlc, ends.Length * 4, storyLength);
        return (names.ToArray(), startPlc, endPlc);
    }

    private static void WriteU16(Stream stream, ushort value)
    {
        stream.WriteByte((byte)value);
        stream.WriteByte((byte)(value >> 8));
    }

    private static void WriteStream(RootStorage storage, string name, byte[] bytes)
    {
        using var stream = storage.CreateStream(name);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static int Align(int value) => checked((value + PageSize - 1) / PageSize * PageSize);
    private static void U32(byte[] target, int offset, int value) =>
        BinaryPrimitives.WriteUInt32LittleEndian(target.AsSpan(offset), checked((uint)value));
}
