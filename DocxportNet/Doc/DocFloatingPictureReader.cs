using System.Buffers.Binary;

namespace DocxportNet.Doc;

public sealed record DocFloatingPicture(uint Cp, uint ShapeId, int LeftTwips, int TopTwips,
    int RightTwips, int BottomTwips, byte[] Bytes, string ContentType,
    DocInlinePictureCrop? Crop = null, byte WrapCode = 2, bool BehindText = false,
    byte WrapSide = 0, byte HorizontalOrigin = 0, byte VerticalOrigin = 0,
    byte HorizontalAlignment = 0, byte VerticalAlignment = 0,
    int DistanceTopEmu = 0, int DistanceBottomEmu = 0,
    int DistanceLeftEmu = 114300, int DistanceRightEmu = 114300,
    bool FlipHorizontal = false, bool FlipVertical = false,
    double RotationDegrees = 0);

/// <summary>Reads picture shapes linked by PlcfSpa, OfficeArt shape properties and the blip store.</summary>
public static class DocFloatingPictureReader
{
    public static IReadOnlyList<DocFloatingPicture> ReadMain(DocStructure structure)
        => Read(structure, "MainShapeAnchors", 0, 0);

    public static IReadOnlyList<DocFloatingPicture> ReadHeaders(DocStructure structure)
        => Read(structure, "HeaderShapeAnchors", 1,
            structure.Parts.FirstOrDefault(x => x.Name == "Headers")?.CpStart ?? 0);

    private static IReadOnlyList<DocFloatingPicture> Read(DocStructure structure,
        string anchorName, byte drawingKind, uint storyStart)
    {
        var anchors = structure.FindLocation(anchorName);
        var drawing = structure.FindLocation("DrawingContent");
        if (anchors?.IsPresent != true || drawing?.IsPresent != true) return [];
        var plc = structure.ReadRange(anchors.StreamName, anchors.Offset, checked((int)anchors.Length));
        if (plc.Length < 4 || (plc.Length - 4) % 30 != 0)
            throw new InvalidDataException("The shape anchor PLC has an invalid length.");
        var count = (plc.Length - 4) / 30;
        var cpBytes = checked((count + 1) * 4);
        var art = structure.ReadRange(drawing.StreamName, drawing.Offset, checked((int)drawing.Length));
        var group = Record(art, 0);
        if (group.Type != 0xF000) throw new InvalidDataException("The drawing group is missing.");
        var store = Children(art, 8, group.End).FirstOrDefault(x => x.Type == 0xF001);
        if (store.Type != 0xF001) return [];
        var blips = Children(art, store.Start + 8, store.End)
            .Where(x => x.Type == 0xF007).ToArray();
        var shapes = new Dictionary<uint, (int BlipId, DocInlinePictureCrop? Crop,
            byte HorizontalAlignment, byte VerticalAlignment,
            int DistanceTopEmu, int DistanceBottomEmu,
            int DistanceLeftEmu, int DistanceRightEmu,
            bool FlipHorizontal, bool FlipVertical, double RotationDegrees)>();
        for (var cursor = group.End; cursor < art.Length;)
        {
            var kind = art[cursor++];
            var container = Record(art, cursor);
            if (container.Type != 0xF002) break;
            if (kind == drawingKind)
                FindShapes(art, cursor + 8, container.End, shapes);
            cursor = container.End;
        }
        var result = new List<DocFloatingPicture>();
        for (var i = 0; i < count; i++)
        {
            var cp = BinaryPrimitives.ReadUInt32LittleEndian(plc.AsSpan(i * 4));
            var at = cpBytes + i * 26;
            var id = BinaryPrimitives.ReadUInt32LittleEndian(plc.AsSpan(at));
            if (!shapes.TryGetValue(id, out var shape) || shape.BlipId <= 0 ||
                shape.BlipId > blips.Length)
                continue;
            var bse = blips[shape.BlipId - 1];
            var bseBody = art.AsSpan(bse.Start + 8, bse.End - bse.Start - 8);
            if (bseBody.Length < 36) throw new InvalidDataException("The picture blip entry is truncated.");
            var delay = BinaryPrimitives.ReadInt32LittleEndian(bseBody.Slice(28));
            var blip = ReadBlip(structure, delay);
            var flags = BinaryPrimitives.ReadUInt16LittleEndian(plc.AsSpan(at + 20));
            result.Add(new DocFloatingPicture(checked(storyStart + cp), id,
                BinaryPrimitives.ReadInt32LittleEndian(plc.AsSpan(at + 4)),
                BinaryPrimitives.ReadInt32LittleEndian(plc.AsSpan(at + 8)),
                BinaryPrimitives.ReadInt32LittleEndian(plc.AsSpan(at + 12)),
                BinaryPrimitives.ReadInt32LittleEndian(plc.AsSpan(at + 16)),
                blip.Bytes, blip.ContentType, shape.Crop,
                checked((byte)((flags >> 5) & 0xF)), (flags & 0x4000) != 0,
                checked((byte)((flags >> 9) & 0xF)),
                checked((byte)((flags >> 1) & 0x3)),
                checked((byte)((flags >> 3) & 0x3)),
                shape.HorizontalAlignment, shape.VerticalAlignment,
                shape.DistanceTopEmu, shape.DistanceBottomEmu,
                shape.DistanceLeftEmu, shape.DistanceRightEmu,
                shape.FlipHorizontal, shape.FlipVertical, shape.RotationDegrees));
        }
        return result;
    }

    private static void FindShapes(byte[] art, int start, int end,
        Dictionary<uint, (int BlipId, DocInlinePictureCrop? Crop,
            byte HorizontalAlignment, byte VerticalAlignment,
            int DistanceTopEmu, int DistanceBottomEmu,
            int DistanceLeftEmu, int DistanceRightEmu,
            bool FlipHorizontal, bool FlipVertical, double RotationDegrees)> shapes)
    {
        foreach (var record in Children(art, start, end))
        {
            if (record.Type == 0xF004)
            {
                uint id = 0;
                var blipId = 0;
                double left = 0, top = 0, right = 0, bottom = 0;
                byte horizontalAlignment = 0, verticalAlignment = 0;
                var distanceTop = 0;
                var distanceBottom = 0;
                var distanceLeft = 114300;
                var distanceRight = 114300;
                var flipHorizontal = false;
                var flipVertical = false;
                var rotationDegrees = 0d;
                foreach (var child in Children(art, record.Start + 8, record.End))
                {
                    if (child.Type == 0xF00A && child.End - child.Start >= 16)
                    {
                        id = BinaryPrimitives.ReadUInt32LittleEndian(art.AsSpan(child.Start + 8));
                        var flags = BinaryPrimitives.ReadUInt32LittleEndian(art.AsSpan(child.Start + 12));
                        flipHorizontal = (flags & 0x40) != 0;
                        flipVertical = (flags & 0x80) != 0;
                    }
                    if (child.Type is 0xF00B or 0xF122)
                    {
                        var propertyCount = BinaryPrimitives.ReadUInt16LittleEndian(art.AsSpan(child.Start)) >> 4;
                        for (var i = 0; i < propertyCount; i++)
                        {
                            var at = child.Start + 8 + i * 6;
                            if (at + 6 > child.End) break;
                            var property = BinaryPrimitives.ReadUInt16LittleEndian(art.AsSpan(at));
                            var value = BinaryPrimitives.ReadInt32LittleEndian(art.AsSpan(at + 2));
                            switch (property)
                            {
                                case 0x0004: rotationDegrees = value / 65536.0; break;
                                case 0x0100: top = value / 65536.0; break;
                                case 0x0101: bottom = value / 65536.0; break;
                                case 0x0102: left = value / 65536.0; break;
                                case 0x0103: right = value / 65536.0; break;
                                case 0x4104: blipId = value; break;
                                case 0x0384: distanceLeft = value; break;
                                case 0x0385: distanceTop = value; break;
                                case 0x0386: distanceRight = value; break;
                                case 0x0387: distanceBottom = value; break;
                                case 0x038F:
                                    if (value is < 0 or > 5)
                                        throw new NotSupportedException(
                                            $"DOC floating picture horizontal alignment {value} is unsupported.");
                                    horizontalAlignment = checked((byte)value); break;
                                case 0x0391:
                                    if (value is < 0 or > 5)
                                        throw new NotSupportedException(
                                            $"DOC floating picture vertical alignment {value} is unsupported.");
                                    verticalAlignment = checked((byte)value); break;
                            }
                        }
                    }
                }
                if (id != 0 && blipId != 0)
                {
                    if (distanceTop < 0 || distanceBottom < 0 ||
                        distanceLeft < 0 || distanceRight < 0)
                        throw new NotSupportedException(
                            "DOC floating picture wrap distance is negative.");
                    var crop = new DocInlinePictureCrop(left, top, right, bottom);
                    shapes[id] = (blipId, crop.IsEmpty ? null : crop,
                        horizontalAlignment, verticalAlignment,
                        distanceTop, distanceBottom, distanceLeft, distanceRight,
                        flipHorizontal, flipVertical,
                        flipHorizontal != flipVertical
                            ? -rotationDegrees : rotationDegrees);
                }
            }
            else if (record.Type is 0xF003 or 0xF002)
                FindShapes(art, record.Start + 8, record.End, shapes);
        }
    }

    private static (byte[] Bytes, string ContentType) ReadBlip(DocStructure structure, int offset)
    {
        if (offset < 0) throw new InvalidDataException("The picture blip has a negative offset.");
        var header = structure.ReadRange("WordDocument", offset, 8);
        var type = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(2));
        var length = BinaryPrimitives.ReadInt32LittleEndian(header.AsSpan(4));
        if (length < 18) throw new InvalidDataException("The picture blip is truncated.");
        var record = structure.ReadRange("WordDocument", offset, checked(length + 8));
        var instance = BinaryPrimitives.ReadUInt16LittleEndian(header) >> 4;
        var prefix = instance switch
        {
            0x46A or 0x6E0 or 0x6E2 or 0x6E4 or 0x7A8 => 17,
            0x46B or 0x6E1 or 0x6E3 or 0x6E5 or 0x7A9 => 33,
            _ => throw new NotSupportedException($"OfficeArt BLIP instance 0x{instance:X3} is unsupported.")
        };
        var bytes = record.AsSpan(8 + prefix).ToArray();
        if (type == 0xF01F)
            return (DocDibBitmap.ToBmp(bytes), "image/bmp");
        var mime = type switch
        {
            0xF01E when bytes.AsSpan().StartsWith(new byte[] { 0x89, 0x50, 0x4E, 0x47 }) => "image/png",
            0xF01D or 0xF02A when bytes.AsSpan().StartsWith(new byte[] { 0xFF, 0xD8 }) => "image/jpeg",
            0xF029 when DocBinaryCompat.IsTiff(bytes) => "image/tiff",
            _ => throw new NotSupportedException($"OfficeArt BLIP type 0x{type:X4} is unsupported.")
        };
        return (bytes, mime);
    }

    private static IEnumerable<(int Start, int End, ushort Type)> Children(byte[] bytes, int start, int end)
    {
        for (var cursor = start; cursor < end;)
        {
            var record = Record(bytes, cursor);
            if (record.End > end) throw new InvalidDataException("An OfficeArt record exceeds its container.");
            yield return record;
            cursor = record.End;
        }
    }

    private static (int Start, int End, ushort Type) Record(byte[] bytes, int start)
    {
        if (start < 0 || start + 8 > bytes.Length)
            throw new InvalidDataException("An OfficeArt record header is truncated.");
        var length = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(start + 4));
        if (length > bytes.Length - start - 8)
            throw new InvalidDataException("An OfficeArt record exceeds the drawing stream.");
        return (start, checked(start + 8 + (int)length),
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(start + 2)));
    }
}
