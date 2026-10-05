using System.Buffers.Binary;
using OpenMcdf;

namespace DocxportNet.Doc;

/// <summary>Enters a structural node. Dispose the returned scope after its children; return null to skip them.</summary>
public interface IDocStructureVisitor
{
    IDisposable? Enter(DocStructureNode node, int depth);
}

/// <summary>Discovers the compound-file tree and FIB directory of a binary .doc.</summary>
public sealed class DocStructureWalker
{
    private static readonly byte[] CompoundSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    public DocStructure Accept(string path, IDocStructureVisitor visitor)
    {
        var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        try
        {
            var structure = Accept(input, visitor);
            structure.OwnInput(input);
            return structure;
        }
        catch
        {
            input.Dispose();
            throw;
        }
    }

    public DocStructure Accept(Stream input, IDocStructureVisitor visitor)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (visitor == null) throw new ArgumentNullException(nameof(visitor));
        var structure = Read(input);
        try
        {
            Walk(structure, structure.Root, 0, visitor);
            return structure;
        }
        catch
        {
            structure.Dispose();
            throw;
        }
    }

    /// <summary>Walks a password-protected RC4 CryptoAPI DOC after decrypting its content streams.</summary>
    public DocStructure Accept(Stream input, IDocStructureVisitor visitor, string password)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (visitor == null) throw new ArgumentNullException(nameof(visitor));
        var decrypted = DocCryptoApiReader.Decrypt(input, password);
        try
        {
            var structure = Read(decrypted);
            try
            {
                structure.Root.Attributes["sourceEncryption"] = "RC4CryptoAPI";
                Walk(structure, structure.Root, 0, visitor);
                structure.OwnInput(decrypted);
                return structure;
            }
            catch
            {
                structure.Dispose();
                throw;
            }
        }
        catch
        {
            decrypted.Dispose();
            throw;
        }
    }

    public DocStructure Accept(string path, IDocStructureVisitor visitor, string password)
    {
        using var input = new FileStream(path, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        return Accept(input, visitor, password);
    }

    private static void Walk(DocStructure structure, DocStructureNode node, int depth, IDocStructureVisitor visitor)
    {
        using var scope = visitor.Enter(node, depth);
        if (scope == null) return;
        if (node.Kind == "Stream" && node.Name == "ObjInfo")
            DocOleNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "Sections")
            DocSectionNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "HeadersAndFooters" or "FootnoteReferences" or
            "FootnoteText" or "EndnoteReferences" or "EndnoteText" or "CommentReferences" or "CommentText")
            DocStoryPlcNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "BookmarkStarts" or "BookmarkEnds" or
            "AnnotationBookmarkStarts" or "AnnotationBookmarkEnds" or
            "FactoidBookmarkStarts" or "FactoidBookmarkEnds" or
            "ConsistencyBookmarkStarts" or "ConsistencyBookmarkEnds" or
            "RepairBookmarkStarts" or "RepairBookmarkEnds")
            DocBookmarkNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "SdtBookmarkStarts" or "SdtBookmarkEnds" or
            "ProtectionBookmarkStarts" or "ProtectionBookmarkEnds")
            DocModernBookmarkNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "MainFields" or "HeaderFields" or
            "FootnoteFields" or "CommentFields" or "EndnoteFields" or "TextboxFields" or
            "HeaderTextboxFields")
            DocFieldNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "ListDefinitions" or "ListOverrides")
            DocListNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "MainShapeAnchors" or "HeaderShapeAnchors")
            DocShapeNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "TableCharacterCache")
            DocTableCharacterNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "DrawingContent" &&
            node.Children.Count == 0)
            node.Children.Add(new DocStructureNode("OfficeArtContent", "Drawings",
                node.StreamName, node.Offset, node.Length));
        if (node.Kind == "OfficeArtContent")
            DocOfficeArtNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "SpellingRanges" or "GrammarRanges" or
            "AutoSummaryRanges" or "SubdocumentRanges" or "LanguageDetectionRanges" or
            "SmartTagRanges")
            DocRangePlcNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "GlossaryRanges")
            DocGlossaryNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "DocumentProperties")
            DocDopNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "FramesAndListStyles")
            DocFrameSetNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "SdtBookmarkNames" or "SchemaReferences")
            DocSdtNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "CommentTree" or "RevisionSaveIds" or
            "SmartTagData" or "OfficeDataSource" or "PageProperties" or
            "UserInfoRanges" or "UserInfoGuids" or "CommandCustomizations" or
            "PrinterDriver" or "PrinterPort" or "PrinterEnvironment" or
            "PrintMergeState" or "UserStrings" or "ToolbarData" or
            "RouteSlip" or "ThreadingMetadata" or "OutlineStyles" or
            "GrammarOptionSets" or "OleControlInfo" or
            "GrammarCookieData" or
            "OldCookieRanges" or "CookieRanges")
            DocMetadataNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "TextboxText" or "HeaderTextboxText" or
            "TextboxBoundaries" or "HeaderTextboxBoundaries")
            DocTextboxNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "BookmarkNames" or "FontTable" or
            "AssociatedStrings" or "CommentBookmarkNames" or "RevisionAuthors" or "ListNames" or
            "FactoidBookmarkNames" or "ConsistencyBookmarkNames" or "RepairBookmarkNames" or
            "Captions" or "AutoCaptions" or "SaveHistory" or "ExternalFileNames" or
            "GlossaryNames" or "GlossaryStyles" or "ListTemplates" or
            "ProtectionBookmarkNames" or "ProtectionUsers")
            DocStringTableNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "TextPieceTable")
            DocClxNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "StyleSheet")
            DocStyleSheetNavigator.Expand(structure, node);
        if (node.Kind == "STSHIB")
            DocStyleSheetNavigator.ExpandHeaderTail(structure, node);
        if (node.Kind == "OfficeDataSource")
            DocOdsoNavigator.Expand(structure, node);
        if (node.Kind is "RecipientInfo" or "FieldMapInfo")
            DocOdsoNavigator.ExpandList(structure, node);
        if (node.Kind == "DofrRglstsf")
            DocFrameSetNavigator.ExpandListStyles(structure, node);
        if (node.Kind == "DofrFsn")
            DocFrameSetNavigator.ExpandFrame(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "LastSelection")
            DocLegacyNavigator.ExpandSelection(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "AuthorFilter")
            DocLegacyNavigator.ExpandAuthorFilter(structure, node);
        if (node.Kind == "STD")
            DocStyleDefinitionNavigator.Expand(structure, node);
        if (node.Kind is "Tcg" or "Tcg255" or "PlfMcd" or "PlfAcd" or "PlfKme" or
            "Kme" or "Acd" or "Cid" or "CidFci" or "TcgSttbf" or "TcgSttbfCore" or "MacroNames" or
            "CTBWRAPPER" or "Customization" or "CTB" or "TBDelta")
            DocCustomizationNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "EmailEnvelope")
            DocEnvelopeNavigator.ExpandLocation(structure, node);
        if (node.Kind == "MsoEnvelope")
            DocEnvelopeNavigator.ExpandBody(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "CharacterFormatting" or "ParagraphFormatting")
            DocFormattingNavigator.ExpandPageTable(structure, node);
        if (node.Kind is "ChpxFkp" or "PapxFkp")
            DocFormattingNavigator.ExpandPage(structure, node);
        if (node.Kind is "Chpx" or "PapxInFkp" or "PrcData" or "Sepx" or
            "UpxPapx" or "UpxChpx" or "UpxTapx")
            DocPropertyNavigator.Expand(structure, node);
        if (node.Kind == "Operand" && node.Attributes.ContainsKey("dataOffset"))
            DocPictureNavigator.Expand(structure, node);
        if (node.Kind == "FarEastLayoutOperand" && node.Children.Count == 0 &&
            node.StreamName != null && node.Offset != null && node.Length == 7 &&
            structure.ReadRange(node.StreamName, node.Offset.Value, 1)[0] == 6)
            node.Children.Add(new DocStructureNode("UFEL", "EastAsianLayoutFlags",
                node.StreamName, node.Offset + 1, 2));
        if (node.Kind == "TDefTableOperand")
            DocTableOperandNavigator.Expand(structure, node);
        if (node.Kind == "NumRMOperand" && node.StreamName != null &&
            node.Offset != null && node.Length == 129 &&
            structure.ReadRange(node.StreamName, node.Offset.Value, 1)[0] == 128)
        {
            var revision = new DocStructureNode("NumRM", "NumberingRevision",
                node.StreamName, node.Offset + 1, 128);
            revision.Children.Add(new DocStructureNode("DTTM", "RevisedAt",
                node.StreamName, revision.Offset + 4, 4));
            node.Children.Add(revision);
        }
        if (node.Kind == "CSSAOperand" && node.StreamName != null &&
            node.Offset != null && node.Length == 7 &&
            structure.ReadRange(node.StreamName, node.Offset.Value, 1)[0] == 6)
        {
            var cssa = new DocStructureNode("CSSA", "CellSpacing", node.StreamName,
                node.Offset + 1, 6);
            cssa.Children.Add(new DocStructureNode("ItcFirstLim", "Cells", node.StreamName,
                cssa.Offset, 2));
            cssa.Children.Add(new DocStructureNode("Fts", "WidthUnit", node.StreamName,
                cssa.Offset + 3, 1));
            node.Children.Add(cssa);
        }
        if (node.Kind == "PbiGrfOperand" && node.StreamName != null &&
            node.Offset != null && node.Length == 2)
        {
            var flags = BinaryPrimitives.ReadUInt16LittleEndian(
                structure.ReadRange(node.StreamName, node.Offset.Value, 2));
            node.Attributes["pictureBullet"] = (flags & 1) != 0 ? "true" : "false";
            node.Attributes["noAutoSize"] = (flags & 2) != 0 ? "true" : "false";
        }
        if (node.Kind == "TLP" && node.StreamName != null &&
            node.Offset != null && node.Length == 4)
        {
            var fatl = new DocStructureNode("Fatl", "TableAutoFormatFlags",
                node.StreamName, node.Offset + 2, 2);
            fatl.Attributes["value"] = BinaryPrimitives.ReadUInt16LittleEndian(
                structure.ReadRange(node.StreamName, node.Offset.Value + 2, 2))
                .ToString(System.Globalization.CultureInfo.InvariantCulture);
            node.Children.Add(fatl);
        }
        if (node.Kind == "TCGRF" && node.Children.Count == 0 &&
            node.StreamName != null && node.Offset != null && node.Length == 2)
        {
            var flags = BinaryPrimitives.ReadUInt16LittleEndian(
                structure.ReadRange(node.StreamName, node.Offset.Value, 2));
            foreach (var (kind, name, value) in new[]
            {
                ("TextFlow", "TextFlow", (flags >> 2) & 7),
                ("VerticalMergeFlag", "VerticalMerge", (flags >> 5) & 3),
                ("VerticalAlign", "VerticalAlignment", (flags >> 7) & 3),
                ("Fts", "WidthUnit", (flags >> 9) & 7)
            })
            {
                var field = new DocStructureNode(kind, name, node.StreamName, node.Offset, 2);
                field.Attributes["value"] = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                node.Children.Add(field);
            }
        }
        if (node.Kind is "BrcOperand" or "TableBordersOperand" or "TableBordersOperand80" or
            "TableBrcOperand" or "TableShadeOperand" or "DefTableShdOperand" or
            "DefTableShd80Operand" or "BrcMayBeNil" or "Brc80MayBeNil" or
            "Brc" or "Brc80" or "Shd" or "Shd80")
            DocBorderShadingNavigator.Expand(structure, node);
        if (node.Kind is "PropRMarkOperand" or "PropRMark" or "SPgbPropOperand")
            DocSectionOperandNavigator.Expand(structure, node);
        if (node.Kind is "PChgTabsOperand" or "PChgTabsPapxOperand" or
            "PChgTabsDel" or "PChgTabsDelClose" or "PChgTabsAdd" or "TBD")
            DocTabOperandNavigator.Expand(structure, node);
        if (node.Kind == "ATRDPre10" && node.Children.Count == 0 &&
            node.StreamName != null && node.Offset != null && node.Length == 30)
        {
            var initials = structure.ReadRange(node.StreamName, node.Offset.Value, 20);
            var count = BinaryPrimitives.ReadUInt16LittleEndian(initials);
            if (count <= 9)
            {
                var author = new DocStructureNode("LPXCharBuffer9", "AuthorInitials",
                    node.StreamName, node.Offset, 20);
                author.Attributes["text"] = System.Text.Encoding.Unicode.GetString(initials,
                    2, count * 2);
                node.Children.Add(author);
            }
        }
        foreach (var child in node.Children)
            Walk(structure, child, depth + 1, visitor);
    }

    public DocStructure Read(Stream input)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (!input.CanRead || !input.CanSeek)
            throw new ArgumentException("The DOC input must be a readable, seekable stream.", nameof(input));
        var signature = new byte[8];
        input.Position = 0;
        ReadFully(input, signature);
        if (!signature.SequenceEqual(CompoundSignature))
            throw new InvalidDataException("The input is not a compound binary file.");
        input.Position = 0;
        var storage = RootStorage.Open(input, StorageModeFlags.LeaveOpen);
        try
        {
        if (!storage.ContainsEntry("WordDocument"))
            throw new InvalidDataException("The compound file has no WordDocument stream.");
        using var wordStream = storage.OpenStream("WordDocument");
        var (fib, lwOffset, pairOffset, pairCount) = ParseFib(wordStream);
        if (fib.IsEncrypted || fib.IsObfuscated)
            throw new NotSupportedException("Encrypted or obfuscated DOC files are not supported.");
        if (!storage.ContainsEntry(fib.TableStreamName))
            throw new InvalidDataException("The selected table stream is missing.");
        using var tableStream = storage.OpenStream(fib.TableStreamName);

        var root = new DocStructureNode("Document", "DOC");
        root.Attributes["tableStream"] = fib.TableStreamName;
        AddStorageEntries(storage, root, "");
        var wordNode = FindChild(root, "WordDocument");
        var fibEnd = pairOffset + pairCount * 8L;
        if (wordStream.Length < fibEnd + 2)
            throw new InvalidDataException("The WordDocument FIB extension count is truncated.");
        wordStream.Position = fibEnd;
        var extensionCountBytes = new byte[2];
        ReadFully(wordStream, extensionCountBytes);
        var extensionWords = U16(extensionCountBytes, 0);
        CheckRange(wordStream.Length, fibEnd + 2, extensionWords * 2L);
        var fibNode = new DocStructureNode("FIB", "FileInformationBlock", "WordDocument", 0,
            fibEnd + 2 + extensionWords * 2L);
        wordNode.Children.Add(fibNode);
        var baseNode = new DocStructureNode("FibBase", "FibBase", "WordDocument", 0, 32);
        baseNode.Attributes["identifier"] = $"0x{fib.Identifier:X4}";
        baseNode.Attributes["version"] = $"0x{fib.Version:X4}";
        baseNode.Attributes["flags"] = $"0x{fib.Flags:X4}";
        baseNode.Attributes["tableStream"] = fib.TableStreamName;
        baseNode.Attributes["fcMin"] = fib.FcMin.ToString(System.Globalization.CultureInfo.InvariantCulture);
        baseNode.Attributes["fcMac"] = fib.FcMac.ToString(System.Globalization.CultureInfo.InvariantCulture);
        fibNode.Children.Add(baseNode);
        fibNode.Children.Add(new DocStructureNode("FibRgW97", "WordFlagsAndKeys",
            "WordDocument", 34, 28));
        var lwBytes = new byte[88];
        wordStream.Position = lwOffset;
        ReadFully(wordStream, lwBytes);
        var parts = ReadParts(lwBytes);
        var lwNode = new DocStructureNode("FibRgLw97", "DocumentPartCounts", "WordDocument", lwOffset, 88);
        lwNode.Attributes["cbMac"] = U32(lwBytes, 0).ToString(System.Globalization.CultureInfo.InvariantCulture);
        foreach (var part in parts)
        {
            var partNode = new DocStructureNode("DocumentPart", part.Name);
            partNode.Attributes["cpStart"] = part.CpStart.ToString(System.Globalization.CultureInfo.InvariantCulture);
            partNode.Attributes["cpEnd"] = part.CpEnd.ToString(System.Globalization.CultureInfo.InvariantCulture);
            partNode.Attributes["length"] = part.Length.ToString(System.Globalization.CultureInfo.InvariantCulture);
            lwNode.Children.Add(partNode);
        }
        fibNode.Children.Add(lwNode);
        var directory = new DocStructureNode("FibDirectory", "FibRgFcLcb", "WordDocument", pairOffset, pairCount * 8L);
        directory.Attributes["entryCount"] = pairCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
        fibNode.Children.Add(directory);
        if (extensionWords > 0)
        {
            var extension = new DocStructureNode("FibRgCswNew", "FibExtension", "WordDocument",
                fibEnd + 2, extensionWords * 2L);
            var extensionBytes = new byte[extensionWords * 2];
            wordStream.Position = fibEnd + 2;
            ReadFully(wordStream, extensionBytes);
            var version = U16(extensionBytes, 0);
            extension.Attributes["version"] = $"0x{version:X4}";
            if (extensionBytes.Length > 2)
                extension.Children.Add(new DocStructureNode(version == 0x0112 ?
                    "FibRgCswNewData2007" : "FibRgCswNewData2000", "VersionData",
                    "WordDocument", fibEnd + 4, extensionBytes.Length - 2));
            fibNode.Children.Add(extension);
        }

        var locations = new List<DocLocation>(pairCount);
        var locationNodes = new Dictionary<int, DocStructureNode>();
        var pair = new byte[8];
        for (var i = 0; i < pairCount; i++)
        {
            var field = pairOffset + i * 8;
            wordStream.Position = field;
            ReadFully(wordStream, pair);
            var offset = U32(pair, 0);
            var length = U32(pair, 4);
            var name = LocationName(i);
            var streamName = i == 79 ? "WordDocument" : fib.TableStreamName;
            var isRange = i is not (0 or 7 or 14 or 20 or 25 or 26 or 34 or 35 or 38 or 39 or 49 or 87);
            var location = new DocLocation(name, streamName, offset, length, i, field, isRange);
            locations.Add(location);
            if (length == 0 || !isRange) continue;
            var node = new DocStructureNode("FibLocation", name, streamName, offset, length);
            node.Attributes["fibIndex"] = i.ToString(System.Globalization.CultureInfo.InvariantCulture);
            node.Attributes["fibFieldOffset"] = field.ToString(System.Globalization.CultureInfo.InvariantCulture);
            directory.Children.Add(node);
            locationNodes[i] = node;
        }
        ValidateLocation(locations, 31, tableStream.Length);
        if (pairCount > 6) ValidateLocation(locations, 6, tableStream.Length);
        if (pairCount > 97) ValidateLocation(locations, 97, tableStream.Length);
        if (pairCount > 33) ValidateLocation(locations, 33, tableStream.Length);
        var structure = new DocStructure(root, fib, locations, parts, storage);
        if (locations.Count > 31 && locations[31].IsPresent && locations[31].Length >= 506)
        {
            var node = new DocStructureNode("Dop2000", "EnvelopeVisibility", fib.TableStreamName,
                locations[31].Offset + 504L, 1);
            node.SetPayloadFactory(() => structure.ReadDopVisibility(fib.TableStreamName, node.Offset!.Value));
            locationNodes[31].Children.Add(node);
        }
        if (locations.Count > 97 && locations[97].IsPresent && locations[97].Length >= 20)
        {
            var node = new DocStructureNode("MsoEnvelopeCLSID", "EmailEnvelopeHeader", fib.TableStreamName,
                locations[97].Offset, 20);
            node.SetPayloadFactory(() => structure.ReadEnvelopeHeader(fib.TableStreamName, node.Offset!.Value));
            locationNodes[97].Children.Add(node);
        }
        return structure;
        }
        catch
        {
            storage.Dispose();
            throw;
        }
    }

    private static void ValidateLocation(IReadOnlyList<DocLocation> locations, int index, long streamLength)
    {
        if (index >= locations.Count || !locations[index].IsPresent) return;
        var location = locations[index];
        if (location.Offset > streamLength || location.Length > streamLength - location.Offset)
            throw new InvalidDataException($"{location.Name} points outside the table stream.");
    }

    private static (DocFibBase Base, int LwOffset, int PairOffset, int PairCount) ParseFib(Stream word)
    {
        if (word.Length < 34)
            throw new InvalidDataException("The WordDocument FIB is invalid.");
        var baseBytes = new byte[34];
        word.Position = 0;
        ReadFully(word, baseBytes);
        if (U16(baseBytes, 0) != 0xA5EC)
            throw new InvalidDataException("The WordDocument FIB is invalid.");
        var version = U16(baseBytes, 2);
        if (version < 0x00C1)
            throw new NotSupportedException("Pre-Word 97 DOC files are not supported.");
        var flags = U16(baseBytes, 10);
        var position = 32;
        position = checked(position + 2 + U16(baseBytes, position) * 2);
        CheckRange(word.Length, position, 2);
        word.Position = position;
        var countBytes = new byte[2];
        ReadFully(word, countBytes);
        var lwCount = U16(countBytes, 0);
        if (lwCount != 22)
            throw new InvalidDataException("The Word FIB has an unexpected FibRgLw97 size.");
        var lwOffset = position + 2;
        position = checked(lwOffset + lwCount * 4);
        CheckRange(word.Length, position, 2);
        word.Position = position;
        ReadFully(word, countBytes);
        var count = U16(countBytes, 0);
        position += 2;
        CheckRange(word.Length, position, count * 8L);
        return (new DocFibBase(U16(baseBytes, 0), version, flags, (flags & 0x0200) != 0 ? "1Table" : "0Table")
        {
            FcMin = U32(baseBytes, 24),
            FcMac = U32(baseBytes, 28)
        }, lwOffset, position, count);
    }

    private static IReadOnlyList<DocPartRange> ReadParts(byte[] lw)
    {
        var fields = new (string Name, int Index)[]
        {
            ("Main", 3), ("Footnotes", 4), ("Headers", 5), ("Comments", 7),
            ("Endnotes", 8), ("Textboxes", 9), ("HeaderTextboxes", 10)
        };
        var parts = new List<DocPartRange>(fields.Length);
        uint start = 0;
        foreach (var (name, index) in fields)
        {
            var count = U32(lw, index * 4);
            if (count >= 0x80000000)
                throw new InvalidDataException($"The {name} character count is negative.");
            var end = checked(start + count);
            if (end >= 0x80000000)
                throw new InvalidDataException("The document-part CP range is too large.");
            if (count != 0) parts.Add(new DocPartRange(name, start, end));
            start = end;
        }
        return parts;
    }

    private static string LocationName(int index) => index switch
    {
        1 => "StyleSheet",
        2 => "FootnoteReferences",
        3 => "FootnoteText",
        4 => "CommentReferences",
        5 => "CommentText",
        6 => "Sections",
        8 => "GlossaryNames",
        9 => "GlossaryRanges",
        11 => "HeadersAndFooters",
        12 => "CharacterFormatting",
        13 => "ParagraphFormatting",
        15 => "FontTable",
        16 => "MainFields",
        17 => "HeaderFields",
        18 => "FootnoteFields",
        19 => "CommentFields",
        21 => "BookmarkNames",
        24 => "CommandCustomizations",
        27 => "PrinterDriver",
        28 => "PrinterPort",
        29 => "PrinterEnvironment",
        30 => "LastSelection",
        22 => "BookmarkStarts",
        23 => "BookmarkEnds",
        32 => "AssociatedStrings",
        37 => "CommentBookmarkNames",
        40 => "MainShapeAnchors",
        41 => "HeaderShapeAnchors",
        42 => "AnnotationBookmarkStarts",
        43 => "AnnotationBookmarkEnds",
        44 => "PrintMergeState",
        79 => "UndoInformation",
        87 => "LastSavedTime",
        31 => "DocumentProperties",
        33 => "TextPieceTable",
        46 => "EndnoteReferences",
        47 => "EndnoteText",
        48 => "EndnoteFields",
        50 => "DrawingContent",
        52 => "Captions",
        53 => "AutoCaptions",
        54 => "SubdocumentRanges",
        55 => "SpellingRanges",
        51 => "RevisionAuthors",
        56 => "TextboxText",
        57 => "TextboxFields",
        58 => "HeaderTextboxText",
        59 => "HeaderTextboxFields",
        60 => "UserStrings",
        61 => "ToolbarData",
        62 => "GrammarCookieData",
        70 => "RouteSlip",
        71 => "SaveHistory",
        72 => "ExternalFileNames",
        73 => "ListDefinitions",
        74 => "ListOverrides",
        75 => "TextboxBoundaries",
        76 => "HeaderTextboxBoundaries",
        82 => "GlossaryStyles",
        83 => "GrammarOptionSets",
        84 => "OleControlInfo",
        89 => "AutoSummaryRanges",
        90 => "GrammarRanges",
        91 => "ListNames",
        93 => "TableCharacterCache",
        96 => "ListTemplates",
        97 => "EmailEnvelope",
        98 => "LanguageDetectionRanges",
        99 => "FramesAndListStyles",
        94 => "ThreadingMetadata",
        100 => "OutlineStyles",
        101 => "OldCookieRanges",
        109 => "PageProperties",
        110 => "UserInfoRanges",
        111 => "UserInfoGuids",
        112 => "CommentTree",
        113 => "RevisionSaveIds",
        114 => "FactoidBookmarkNames",
        115 => "FactoidBookmarkStarts",
        116 => "CookieRanges",
        117 => "FactoidBookmarkEnds",
        118 => "SmartTagData",
        120 => "ConsistencyBookmarkNames",
        121 => "ConsistencyBookmarkStarts",
        122 => "ConsistencyBookmarkEnds",
        123 => "RepairBookmarkNames",
        124 => "RepairBookmarkStarts",
        125 => "RepairBookmarkEnds",
        130 => "OfficeDataSource",
        132 => "SmartTagRanges",
        135 => "AuthorFilter",
        136 => "SchemaReferences",
        137 => "SdtBookmarkNames",
        138 => "SdtBookmarkStarts",
        139 => "SdtBookmarkEnds",
        141 => "ProtectionBookmarkNames",
        142 => "ProtectionBookmarkStarts",
        143 => "ProtectionBookmarkEnds",
        144 => "ProtectionUsers",
        _ => $"FibEntry{index}"
    };

    private static void AddStorageEntries(Storage storage, DocStructureNode parent, string parentPath)
    {
        foreach (var entry in storage.EnumerateEntries().OrderBy(x => x.Name, StringComparer.Ordinal))
        {
            var path = parentPath.Length == 0 ? entry.Name : parentPath + "/" + entry.Name;
            var node = new DocStructureNode(entry.Type == EntryType.Storage ? "Storage" : "Stream", entry.Name,
                entry.Type == EntryType.Stream ? path : null, null,
                entry.Type == EntryType.Stream ? entry.Length : null);
            node.Attributes["path"] = path;
            parent.Children.Add(node);
            if (entry.Type == EntryType.Storage)
            {
                var child = storage.OpenStorage(entry.Name);
                AddStorageEntries(child, node, path);
            }
        }
    }

    private static DocStructureNode FindChild(DocStructureNode node, string name) =>
        node.Children.First(x => x.Name == name);

    private static ushort U16(byte[] bytes, int offset)
    {
        CheckRange(bytes.Length, offset, 2);
        return BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset));
    }

    private static uint U32(byte[] bytes, int offset)
    {
        CheckRange(bytes.Length, offset, 4);
        return BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset));
    }

    private static void CheckRange(long available, long offset, long length)
    {
        if (offset < 0 || length < 0 || offset > available || length > available - offset)
            throw new InvalidDataException("The FIB references bytes outside WordDocument.");
    }

    private static void ReadFully(Stream stream, byte[] bytes)
    {
        var position = 0;
        while (position < bytes.Length)
        {
            var read = stream.Read(bytes, position, bytes.Length - position);
            if (read == 0) throw new InvalidDataException("The DOC stream ended unexpectedly.");
            position += read;
        }
    }
}
