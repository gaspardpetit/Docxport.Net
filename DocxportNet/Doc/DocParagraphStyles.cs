using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>A paragraph style assignment mapped back from PAPX FCs to story CPs.</summary>
public sealed record DocParagraphStyleRange(uint CpStart, uint CpEnd, int StyleIndex,
    DocParagraphFormatting? Formatting = null);

internal static class DocParagraphStyleReader
{
    public static IReadOnlyList<DocParagraphStyleRange> Read(DocTextIndex index)
    {
        var result = new List<DocParagraphStyleRange>();
        var pieceProperties = new DocPiecePropertyReader(index);
        foreach (var page in index.FormattingPages.Where(x => !x.IsCharacterFormatting))
        foreach (var run in page.Runs.Where(x => x.Kind == "PapxRange"))
        {
            var property = run.Children.FirstOrDefault(x => x.Kind == "PapxInFkp");
            var styleIndex = 0;
            var formatting = DocParagraphFormatting.Empty;
            DocParagraphFormatting? referencedTable = null;
            var directSprms = Array.Empty<byte>();
            if (property?.Offset is long offset && property.Length is long length)
            {
                var bytes = index.Structure.ReadRange("WordDocument", offset, checked((int)length));
                var styleOffset = bytes[0] == 0 ? 2 : 1;
                if (bytes.Length < styleOffset + 2)
                    throw new InvalidDataException("A PAPX has no paragraph style index.");
                styleIndex = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(styleOffset));
                // PAPX may end with one alignment byte after its SPRMs.
                directSprms = bytes.AsSpan(styleOffset + 2).ToArray();
                var (resolvedSprms, resolvedNode, resolvedOffset) = ResolveHugePapx(
                    index.Structure, directSprms, property, offset + styleOffset + 2);
                directSprms = resolvedSprms;
                referencedTable = ReadReferencedTableProperties(index.Structure,
                    resolvedNode, directSprms, resolvedOffset);
                try { formatting = DocParagraphFormatting.Parse(directSprms); }
                catch (InvalidDataException) when (directSprms.Length > 0 &&
                    directSprms[directSprms.Length - 1] == 0)
                {
                    directSprms = directSprms.AsSpan(0, directSprms.Length - 1).ToArray();
                    formatting = DocParagraphFormatting.Parse(directSprms);
                }
                if (referencedTable != null)
                    formatting = ApplyReferencedTableProperties(formatting, referencedTable);
            }
            var fcStart = long.Parse(run.Attributes["fcStart"], CultureInfo.InvariantCulture);
            var fcEnd = long.Parse(run.Attributes["fcEnd"], CultureInfo.InvariantCulture);
            foreach (var piece in index.Pieces)
            {
                var start = Math.Max(fcStart, piece.TextOffset);
                var end = Math.Min(fcEnd, piece.TextOffset + piece.TextByteLength);
                if (start >= end) continue;
                var pieceSprms = pieceProperties.Read(piece, character: false);
                var combined = formatting;
                if (pieceSprms.Length != 0)
                {
                    var sprms = new byte[directSprms.Length + pieceSprms.Length];
                    directSprms.CopyTo(sprms, 0);
                    pieceSprms.CopyTo(sprms, directSprms.Length);
                    combined = DocParagraphFormatting.Parse(sprms);
                    if (referencedTable != null)
                        combined = ApplyReferencedTableProperties(combined, referencedTable);
                }
                var width = piece.Encoding == "utf16" ? 2 : 1;
                if ((start - piece.TextOffset) % width != 0 ||
                    (end - piece.TextOffset) % width != 0)
                    throw new InvalidDataException("A PAPX boundary splits a character.");
                result.Add(new DocParagraphStyleRange(
                    piece.GetCharacterPosition(start),
                    piece.GetCharacterPosition(end),
                    styleIndex, combined));
            }
        }
        return result.OrderBy(x => x.CpStart).ToArray();
    }

    private static (byte[] Sprms, DocStructureNode Node, long Offset) ResolveHugePapx(
        DocStructure structure, byte[] sprms, DocStructureNode property, long sprmOffset)
    {
        var offsets = new HashSet<uint>();
        for (var depth = 0; sprms.Length >= 6 &&
            BinaryPrimitives.ReadUInt16LittleEndian(sprms) == 0x6646; depth++)
        {
            if (depth == 32)
                throw new InvalidDataException("A DOC huge PAPX chain is too deep.");
            if (sprms.AsSpan(6).IndexOfAnyExcept((byte)0) >= 0)
                throw new InvalidDataException("A DOC huge PAPX has trailing modifiers.");
            var dataOffset = BinaryPrimitives.ReadUInt32LittleEndian(sprms.AsSpan(2));
            if (!offsets.Add(dataOffset))
                throw new InvalidDataException("A DOC huge PAPX reference is cyclic.");
            var length = BinaryPrimitives.ReadUInt16LittleEndian(
                structure.ReadRange("Data", dataOffset, 2));
            if (length < 10)
                throw new InvalidDataException("A DOC huge PAPX has too few property bytes.");
            property = new DocStructureNode("PrcData", "HugeParagraphProperties",
                "Data", dataOffset, checked(2 + length));
            sprmOffset = dataOffset + 2;
            sprms = structure.ReadRange("Data", sprmOffset, length);
        }
        return (sprms, property, sprmOffset);
    }

    private static DocParagraphFormatting ApplyReferencedTableProperties(
        DocParagraphFormatting fallback, DocParagraphFormatting modern) => fallback with
    {
        TableCellEdges = modern.TableCellEdges ?? fallback.TableCellEdges,
        TableCellPreferredWidths = modern.TableCellPreferredWidths ??
            fallback.TableCellPreferredWidths,
        TableCellHorizontalMerges = modern.TableCellHorizontalMerges ??
            fallback.TableCellHorizontalMerges,
        TableCellBorders = modern.TableCellBorders ?? fallback.TableCellBorders,
        TableAutoFit = modern.TableAutoFit ?? fallback.TableAutoFit,
        TablePreferredWidth = modern.TablePreferredWidth ?? fallback.TablePreferredWidth,
        TableRowOriginTwips = modern.TableRowOriginTwips ?? fallback.TableRowOriginTwips
    };

    private static DocParagraphFormatting? ReadReferencedTableProperties(DocStructure structure,
        DocStructureNode property, byte[] directSprms, long sprmOffset)
    {
        var prefix = new List<byte>();
        var found = false;
        var node = property;
        var bytes = directSprms;
        var start = sprmOffset;
        for (var depth = 0; depth < 32; depth++)
        {
            if (bytes.AsSpan().IndexOf(new byte[] { 0x6B, 0x64 }) >= 0)
                DocPropertyNavigator.Expand(structure, node);
            var modifiers = node.Kind == "PapxInFkp"
                ? node.Children.SelectMany(x => x.Children) : node.Children;
            var reference = modifiers.FirstOrDefault(x => x.Kind == "Prl" &&
                x.Children.FirstOrDefault()?.Attributes.TryGetValue("code", out var code) == true &&
                code == "0x646B");
            if (reference == null)
            {
                if (!found) return null;
                prefix.AddRange(bytes);
                return DocParagraphFormatting.Parse(prefix.ToArray());
            }
            found = true;
            var precedingLength = checked((int)(reference.Offset!.Value - start));
            if (precedingLength < 0 || precedingLength > bytes.Length)
                throw new InvalidDataException("A DOC table-property reference is misplaced.");
            prefix.AddRange(bytes.AsSpan(0, precedingLength).ToArray());
            var target = reference.Children.SelectMany(x => x.Children)
                .Single(x => x.Kind == "PrcData");
            bytes = structure.ReadRange("Data", target.Offset!.Value + 2,
                checked((int)target.Length!.Value - 2));
            start = target.Offset.Value + 2;
            node = target;
        }
        throw new InvalidDataException("A DOC table-property reference is too deep.");
    }
}
