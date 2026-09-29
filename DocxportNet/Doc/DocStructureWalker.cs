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

    private static void Walk(DocStructure structure, DocStructureNode node, int depth, IDocStructureVisitor visitor)
    {
        using var scope = visitor.Enter(node, depth);
        if (scope == null) return;
        if (node.Kind == "FibLocation" && node.Name == "Sections")
            DocSectionNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "TextPieceTable")
            DocClxNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "StyleSheet")
            DocStyleSheetNavigator.Expand(structure, node);
        if (node.Kind == "FibLocation" && node.Name == "EmailEnvelope")
            DocEnvelopeNavigator.ExpandLocation(structure, node);
        if (node.Kind == "MsoEnvelope")
            DocEnvelopeNavigator.ExpandBody(structure, node);
        if (node.Kind == "FibLocation" && node.Name is "CharacterFormatting" or "ParagraphFormatting")
            DocFormattingNavigator.ExpandPageTable(structure, node);
        if (node.Kind is "ChpxFkp" or "PapxFkp")
            DocFormattingNavigator.ExpandPage(structure, node);
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
        var fibNode = new DocStructureNode("FIB", "FileInformationBlock", "WordDocument", 0, pairOffset + pairCount * 8L);
        wordNode.Children.Add(fibNode);
        var baseNode = new DocStructureNode("FibBase", "FibBase", "WordDocument", 0, 32);
        baseNode.Attributes["identifier"] = $"0x{fib.Identifier:X4}";
        baseNode.Attributes["version"] = $"0x{fib.Version:X4}";
        baseNode.Attributes["flags"] = $"0x{fib.Flags:X4}";
        baseNode.Attributes["tableStream"] = fib.TableStreamName;
        baseNode.Attributes["fcMin"] = fib.FcMin.ToString(System.Globalization.CultureInfo.InvariantCulture);
        baseNode.Attributes["fcMac"] = fib.FcMac.ToString(System.Globalization.CultureInfo.InvariantCulture);
        fibNode.Children.Add(baseNode);
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
        6 => "Sections",
        11 => "HeadersAndFooters",
        12 => "CharacterFormatting",
        13 => "ParagraphFormatting",
        15 => "FontTable",
        79 => "UndoInformation",
        87 => "LastSavedTime",
        31 => "DocumentProperties",
        33 => "TextPieceTable",
        97 => "EmailEnvelope",
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
