using System.Buffers.Binary;

namespace DocxportNet.Doc;

/// <summary>Writes pictures with margin-relative square-wrapped SPAs.</summary>
internal static class DocFloatingPictureWriter
{
    public static (byte[] MainAnchorPlc, byte[] HeaderAnchorPlc,
        byte[] DrawingContent, byte[] Blips) Write(
        IReadOnlyList<DocPlainTextFloatingPicture> floatingPictures, int mainLength,
        int totalLength, int blipOffset)
    {
        if (floatingPictures.Count == 0)
            throw new ArgumentException("At least one floating picture is required.",
                nameof(floatingPictures));
        var ordered = floatingPictures.OrderBy(x => x.Cp).ToArray();
        var main = ordered.Where(x => x.Cp < mainLength).ToArray();
        var headers = ordered.Where(x => x.Cp >= mainLength).ToArray();
        var bseRecords = new List<byte[]>();
        var blipRecords = new List<byte[]>();
        var mainShapes = new List<byte[]>();
        var headerShapes = new List<byte[]>();
        for (var i = 0; i < ordered.Length; i++)
        {
            var floating = ordered[i];
            var picture = floating.Picture;
            if (floating.Cp < 0 || floating.Cp >= totalLength ||
                i > 0 && floating.Cp == ordered[i - 1].Cp)
                throw new InvalidDataException("A floating picture has no unique story anchor.");
            if (picture.ContentType is not ("image/png" or "image/jpeg" or "image/tiff" or "image/bmp"))
                throw new NotSupportedException("DOC floating pictures require PNG, JPEG, TIFF, or BMP bytes.");
            var width = checked((int)Math.Round(picture.WidthEmu / 635m));
            var height = checked((int)Math.Round(picture.HeightEmu / 635m));
            if (width <= 0 || height <= 0)
                throw new InvalidDataException("A DOC floating picture needs positive bounds.");
            var uid = DocBinaryCompat.Md5(picture.Bytes);
            var jpeg = picture.ContentType == "image/jpeg";
            var tiff = picture.ContentType == "image/tiff";
            var bmp = picture.ContentType == "image/bmp";
            if (!bmp && (tiff ? !DocBinaryCompat.IsTiff(picture.Bytes) : jpeg ?
                !picture.Bytes.AsSpan().StartsWith(new byte[] { 0xFF, 0xD8 }) :
                !picture.Bytes.AsSpan().StartsWith(new byte[] { 0x89, 0x50, 0x4E, 0x47 })))
                throw new InvalidDataException("The floating picture bytes do not match their type.");
            var blipBytes = bmp ? DocDibBitmap.ToDib(picture.Bytes) : picture.Bytes;
            var blip = Record(0, bmp ? 0x7A8 : tiff ? 0x6E4 : jpeg ? 0x46A : 0x6E0,
                bmp ? 0xF01F : tiff ? 0xF029 : jpeg ? 0xF01D : 0xF01E,
                Join(uid, [0xFF], blipBytes));
            var bse = new byte[36];
            bse[0] = bse[1] = bmp ? (byte)7 : tiff ? (byte)0x11 : jpeg ? (byte)5 : (byte)6;
            uid.CopyTo(bse, 2);
            BinaryPrimitives.WriteUInt16LittleEndian(bse.AsSpan(18), 0x00FF);
            BinaryPrimitives.WriteUInt32LittleEndian(bse.AsSpan(20), checked((uint)blip.Length));
            BinaryPrimitives.WriteUInt32LittleEndian(bse.AsSpan(24), 1);
            BinaryPrimitives.WriteUInt32LittleEndian(bse.AsSpan(28), checked((uint)blipOffset));
            blipOffset = checked(blipOffset + blip.Length);
            bseRecords.Add(Record(2, bmp ? 7 : tiff ? 0x11 : jpeg ? 5 : 6, 0xF007, bse));
            blipRecords.Add(blip);
            var shape = PictureShape(picture.Crop, floating.HorizontalOrigin,
                floating.VerticalOrigin, floating.HorizontalAlignment,
                floating.VerticalAlignment, floating.DistanceTopEmu,
                floating.DistanceBottomEmu, floating.DistanceLeftEmu,
                floating.DistanceRightEmu, picture.FlipHorizontal,
                picture.FlipVertical, picture.RotationDegrees);
            var inHeader = floating.Cp >= mainLength;
            var storyIndex = inHeader ? headerShapes.Count : mainShapes.Count;
            var shapeId = inHeader
                ? (main.Length > 0 ? 2049 : 1025) + storyIndex
                : 1026 + storyIndex;
            BinaryPrimitives.WriteUInt32LittleEndian(shape.AsSpan(16), checked((uint)shapeId));
            var propertyCount = BinaryPrimitives.ReadUInt16LittleEndian(shape.AsSpan(24)) >> 4;
            for (var propertyIndex = 0; propertyIndex < propertyCount; propertyIndex++)
            {
                var at = 32 + propertyIndex * 6;
                if (BinaryPrimitives.ReadUInt16LittleEndian(shape.AsSpan(at)) == 0x4104)
                {
                    BinaryPrimitives.WriteUInt32LittleEndian(shape.AsSpan(at + 2),
                        checked((uint)(i + 1)));
                }
                if ((floating.WrapCode == 3 && floating.BehindText ||
                    floating.WrapCode == 4) &&
                    BinaryPrimitives.ReadUInt16LittleEndian(shape.AsSpan(at)) == 0x03BF)
                {
                    var flags = BinaryPrimitives.ReadUInt32LittleEndian(shape.AsSpan(at + 2));
                    BinaryPrimitives.WriteUInt32LittleEndian(shape.AsSpan(at + 2), flags | 0x20);
                }
            }
            if (inHeader) headerShapes.Add(shape);
            else mainShapes.Add(shape);
        }
        var drawingCount = (main.Length > 0 ? 1 : 0) + (headers.Length > 0 ? 1 : 0);
        var dggBody = new byte[16 + drawingCount * 8];
        var maxShapeId = headers.Length > 0
            ? (main.Length > 0 ? 2048 : 1024) + headers.Length
            : 1025 + main.Length;
        BinaryPrimitives.WriteUInt32LittleEndian(dggBody, checked((uint)(maxShapeId + 1)));
        BinaryPrimitives.WriteUInt32LittleEndian(dggBody.AsSpan(4),
            checked((uint)(drawingCount + 1)));
        BinaryPrimitives.WriteUInt32LittleEndian(dggBody.AsSpan(8),
            checked((uint)(ordered.Length + drawingCount)));
        BinaryPrimitives.WriteUInt32LittleEndian(dggBody.AsSpan(12),
            checked((uint)drawingCount));
        var cluster = 16;
        if (main.Length > 0)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(dggBody.AsSpan(cluster), 1);
            BinaryPrimitives.WriteUInt32LittleEndian(dggBody.AsSpan(cluster + 4),
                checked((uint)(main.Length + 2)));
            cluster += 8;
        }
        if (headers.Length > 0)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(dggBody.AsSpan(cluster),
                main.Length > 0 ? 2u : 1u);
            BinaryPrimitives.WriteUInt32LittleEndian(dggBody.AsSpan(cluster + 4),
                checked((uint)(headers.Length + 1)));
        }
        var group = Record(15, 0, 0xF000, Join(
            Record(0, 0, 0xF006, dggBody),
            Record(15, ordered.Length, 0xF001, Join(bseRecords.ToArray())),
            DocBinaryCompat.Hex("40001ef110000000ffff00000000ff0080808000f7000010")));
        var drawings = new List<byte[]> { group };
        if (mainShapes.Count > 0)
            drawings.Add(Join([0], Drawing(1, 1024, mainShapes, true)));
        if (headerShapes.Count > 0)
            drawings.Add(Join([1], Drawing(mainShapes.Count > 0 ? 2 : 1,
                mainShapes.Count > 0 ? 2048 : 1024, headerShapes, false)));
        var art = Join(drawings.ToArray());
        var mainPlc = main.Length > 0 ? AnchorPlc(main, 0, mainLength,
            i => checked((uint)(1026 + i))) : [];
        var headerBase = main.Length > 0 ? 2049 : 1025;
        var headerPlc = headers.Length > 0 ? AnchorPlc(headers, mainLength,
            totalLength - mainLength,
            i => checked((uint)(headerBase + i))) : [];
        return (mainPlc, headerPlc, art, Join(blipRecords.ToArray()));
    }

    private static byte[] Drawing(int drawingId, int baseShapeId,
        IReadOnlyList<byte[]> pictureShapes, bool includeBackground)
    {
        var root = DocBinaryCompat.Hex(
            "0f0004f028000000010009f0100000000000000000000000000000000000000002000af0080000000004000005000000");
        BinaryPrimitives.WriteUInt32LittleEndian(root.AsSpan(40), checked((uint)baseShapeId));
        var groupShapes = Record(15, 0, 0xF003, Join([root, .. pictureShapes]));
        var dgBody = new byte[8];
        BinaryPrimitives.WriteUInt32LittleEndian(dgBody,
            checked((uint)(pictureShapes.Count + 1)));
        BinaryPrimitives.WriteUInt32LittleEndian(dgBody.AsSpan(4),
            checked((uint)(baseShapeId + pictureShapes.Count + (includeBackground ? 1 : 0))));
        var records = new List<byte[]>
        {
            Record(0, drawingId, 0xF008, dgBody), groupShapes
        };
        if (includeBackground)
        {
            var background = DocBinaryCompat.Hex(
                "0f0004f04200000012000af00800000001040000000e000053000bf01e000000" +
                "bf0100001000cb0100000000ff01000008000403090000003f0301000100" +
                "000011f00400000001000000");
            BinaryPrimitives.WriteUInt32LittleEndian(background.AsSpan(16),
                checked((uint)(baseShapeId + 1)));
            records.Add(background);
        }
        return Record(15, 0, 0xF002, Join(records.ToArray()));
    }

    private static byte[] AnchorPlc(IReadOnlyList<DocPlainTextFloatingPicture> pictures,
        int storyStart, int storyLength, Func<int, uint> shapeId)
    {
        var cpBytes = checked((pictures.Count + 1) * 4);
        var plc = new byte[checked(cpBytes + pictures.Count * 26)];
        for (var i = 0; i < pictures.Count; i++)
        {
            var floating = pictures[i];
            var width = checked((int)Math.Round(floating.Picture.WidthEmu / 635m));
            var height = checked((int)Math.Round(floating.Picture.HeightEmu / 635m));
            BinaryPrimitives.WriteUInt32LittleEndian(plc.AsSpan(i * 4),
                checked((uint)(floating.Cp - storyStart)));
            var at = cpBytes + i * 26;
            BinaryPrimitives.WriteUInt32LittleEndian(plc.AsSpan(at), shapeId(i));
            BinaryPrimitives.WriteInt32LittleEndian(plc.AsSpan(at + 4), floating.LeftTwips);
            BinaryPrimitives.WriteInt32LittleEndian(plc.AsSpan(at + 8), floating.TopTwips);
            BinaryPrimitives.WriteInt32LittleEndian(plc.AsSpan(at + 12),
                checked(floating.LeftTwips + width));
            BinaryPrimitives.WriteInt32LittleEndian(plc.AsSpan(at + 16),
                checked(floating.TopTwips + height));
            if (floating.WrapCode is not (1 or 2 or 3 or 4 or 5))
                throw new NotSupportedException(
                    $"DOC floating picture wrap code {floating.WrapCode} is unsupported.");
            if (floating.WrapSide > 3)
                throw new NotSupportedException(
                    $"DOC floating picture wrap side {floating.WrapSide} is unsupported.");
            if (floating.HorizontalOrigin > 2 || floating.VerticalOrigin > 2)
                throw new NotSupportedException("DOC floating picture position origin is unsupported.");
            BinaryPrimitives.WriteUInt16LittleEndian(plc.AsSpan(at + 20),
                checked((ushort)((floating.HorizontalOrigin << 1) |
                    (floating.VerticalOrigin << 3) | (floating.WrapCode << 5) |
                    (floating.WrapCode is 2 or 4 or 5 ? floating.WrapSide << 9 : 0) |
                    (floating.WrapCode == 3 && floating.BehindText ? 0x4000 : 0))));
        }
        BinaryPrimitives.WriteUInt32LittleEndian(plc.AsSpan(pictures.Count * 4),
            checked((uint)storyLength));
        return plc;
    }

    private static byte[] PictureShape(DocInlinePictureCrop? crop,
        byte horizontalOrigin, byte verticalOrigin,
        byte horizontalAlignment, byte verticalAlignment,
        int distanceTop, int distanceBottom, int distanceLeft, int distanceRight,
        bool flipHorizontal, bool flipVertical, double rotationDegrees)
    {
        var template = DocBinaryCompat.Hex(
            "0f0004f08e000000b2040af00800000002040000000a000063000bf038000000" +
            "0441010000003f0100000600bf0100001000ff010000080080c314000000bf0300002200" +
            "5000690063007400750072006500200031000000" +
            "530022f11e000000900300000000920300000000bf0300820082c40700000000c50700000000" +
            "000010f00400000000000000000011f00400000001000000");
        BinaryPrimitives.WriteUInt32LittleEndian(template.AsSpan(20),
            0x00000A00U | (flipHorizontal ? 0x40U : 0U) |
            (flipVertical ? 0x80U : 0U));
        var properties = new List<byte>();
        void Add(ushort property, double value)
        {
            if (value == 0) return;
            Span<byte> entry = stackalloc byte[6];
            BinaryPrimitives.WriteUInt16LittleEndian(entry, property);
            BinaryPrimitives.WriteInt32LittleEndian(entry.Slice(2),
                checked((int)Math.Round(value * 65536)));
            properties.AddRange(entry.ToArray());
        }
        if (!DocBinaryCompat.IsFinite(rotationDegrees))
            throw new NotSupportedException("DOC floating image rotation is invalid.");
        Add(0x0004, flipHorizontal != flipVertical
            ? -rotationDegrees : rotationDegrees);
        if (crop is { IsEmpty: false })
        {
            if (!DocBinaryCompat.IsFinite(crop.Left) || !DocBinaryCompat.IsFinite(crop.Top) ||
                !DocBinaryCompat.IsFinite(crop.Right) || !DocBinaryCompat.IsFinite(crop.Bottom) ||
                crop.Left < 0 || crop.Top < 0 || crop.Right < 0 || crop.Bottom < 0 ||
                crop.Left + crop.Right >= 1 || crop.Top + crop.Bottom >= 1)
                throw new NotSupportedException("DOC floating image crop fractions are invalid.");
            Add(0x0100, crop.Top);
            Add(0x0101, crop.Bottom);
            Add(0x0102, crop.Left);
            Add(0x0103, crop.Right);
        }
        if (properties.Count == 0)
            return WithPositionProperties(template, horizontalOrigin, verticalOrigin,
                horizontalAlignment, verticalAlignment,
                distanceTop, distanceBottom, distanceLeft, distanceRight);
        var result = new byte[template.Length + properties.Count];
        template.AsSpan(0, 32).CopyTo(result);
        properties.ToArray().CopyTo(result, 32);
        template.AsSpan(32).CopyTo(result.AsSpan(32 + properties.Count));
        BinaryPrimitives.WriteUInt32LittleEndian(result.AsSpan(4),
            checked((uint)(142 + properties.Count)));
        BinaryPrimitives.WriteUInt16LittleEndian(result.AsSpan(24),
            checked((ushort)(((6 + properties.Count / 6) << 4) | 3)));
        BinaryPrimitives.WriteUInt32LittleEndian(result.AsSpan(28),
            checked((uint)(56 + properties.Count)));
        return WithPositionProperties(result, horizontalOrigin, verticalOrigin,
            horizontalAlignment, verticalAlignment,
            distanceTop, distanceBottom, distanceLeft, distanceRight);
    }

    private static byte[] WithPositionProperties(byte[] shape,
        byte horizontalOrigin, byte verticalOrigin,
        byte horizontalAlignment, byte verticalAlignment,
        int distanceTop, int distanceBottom, int distanceLeft, int distanceRight)
    {
        if (horizontalOrigin > 2 || verticalOrigin > 2)
            throw new NotSupportedException("DOC floating picture position origin is unsupported.");
        if (horizontalAlignment > 5 || verticalAlignment > 5)
            throw new NotSupportedException("DOC floating picture alignment is unsupported.");
        if (distanceTop < 0 || distanceBottom < 0 ||
            distanceLeft < 0 || distanceRight < 0)
            throw new NotSupportedException("DOC floating picture wrap distance is negative.");
        if (horizontalOrigin == 0 && verticalOrigin == 0 &&
            horizontalAlignment == 0 && verticalAlignment == 0 &&
            distanceTop == 0 && distanceBottom == 0 &&
            distanceLeft == 114300 && distanceRight == 114300) return shape;
        var propertyStart = 24;
        var propertyLength = BinaryPrimitives.ReadInt32LittleEndian(shape.AsSpan(propertyStart + 4));
        var positionStart = checked(propertyStart + 8 + propertyLength);
        if (BinaryPrimitives.ReadUInt16LittleEndian(shape.AsSpan(positionStart + 2)) != 0xF122)
            throw new InvalidDataException("The OfficeArt position property record is missing.");
        var properties = new List<byte[]>();
        static byte[] PositionProperty(ushort property, uint value)
        {
            var bytes = new byte[6];
            BinaryPrimitives.WriteUInt16LittleEndian(bytes, property);
            BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(2), value);
            return bytes;
        }
        if (distanceLeft != 114300)
            properties.Add(PositionProperty(0x0384, checked((uint)distanceLeft)));
        if (distanceTop != 0)
            properties.Add(PositionProperty(0x0385, checked((uint)distanceTop)));
        if (distanceRight != 114300)
            properties.Add(PositionProperty(0x0386, checked((uint)distanceRight)));
        if (distanceBottom != 0)
            properties.Add(PositionProperty(0x0387, checked((uint)distanceBottom)));
        if (horizontalAlignment != 0)
            properties.Add(PositionProperty(0x038F, horizontalAlignment));
        if (horizontalOrigin != 2)
            properties.Add(PositionProperty(0x0390, horizontalOrigin));
        if (verticalAlignment != 0)
            properties.Add(PositionProperty(0x0391, verticalAlignment));
        if (verticalOrigin != 2)
            properties.Add(PositionProperty(0x0392, verticalOrigin));
        properties.AddRange([
            PositionProperty(0x03BF, 0x82008200),
            PositionProperty(0x07C4, 0),
            PositionProperty(0x07C5, 0)
        ]);
        var record = Record(3, properties.Count, 0xF122, Join(properties.ToArray()));
        var oldLength = BinaryPrimitives.ReadInt32LittleEndian(shape.AsSpan(positionStart + 4)) + 8;
        var output = Join(shape.AsSpan(0, positionStart).ToArray(), record,
            shape.AsSpan(positionStart + oldLength).ToArray());
        BinaryPrimitives.WriteUInt32LittleEndian(output.AsSpan(4),
            checked((uint)(output.Length - 8)));
        return output;
    }

    private static byte[] Record(int version, int instance, int type, byte[] body)
    {
        var result = new byte[8 + body.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(result, checked((ushort)((instance << 4) | version)));
        BinaryPrimitives.WriteUInt16LittleEndian(result.AsSpan(2), checked((ushort)type));
        BinaryPrimitives.WriteUInt32LittleEndian(result.AsSpan(4), checked((uint)body.Length));
        body.CopyTo(result, 8);
        return result;
    }

    private static byte[] Join(params byte[][] parts)
    {
        var output = new byte[parts.Sum(x => x.Length)];
        var cursor = 0;
        foreach (var part in parts)
        {
            part.CopyTo(output, cursor);
            cursor += part.Length;
        }
        return output;
    }
}
