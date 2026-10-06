using System.Buffers.Binary;

namespace DocxportNet.Doc;

/// <summary>An inline DOC image and its displayed extent in English metric units.</summary>
internal sealed record DocInlinePicture(byte[] Bytes, string ContentType, long WidthEmu,
    long HeightEmu, DocInlinePictureCrop? Crop = null, bool FlipHorizontal = false,
    bool FlipVertical = false, double RotationDegrees = 0, Uri? LinkedImage = null);

public sealed record DocInlinePictureCrop(double Left, double Top, double Right,
    double Bottom)
{
    public bool IsEmpty => Left == 0 && Top == 0 && Right == 0 && Bottom == 0;
}

internal static class DocInlinePictureReader
{
    public static DocInlinePicture Read(DocStructure structure, int dataOffset)
        => TryRead(structure, dataOffset) ?? throw new NotSupportedException(
            "The DOC inline picture has no supported embedded blip or linked image.");

    public static DocInlinePicture? TryRead(DocStructure structure, int dataOffset)
    {
        if (dataOffset < 0)
            throw new InvalidDataException("A DOC picture has a negative Data-stream offset.");
        var header = structure.ReadRange("Data", dataOffset, 68);
        var totalLength = BinaryPrimitives.ReadInt32LittleEndian(header);
        var headerLength = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(4));
        if (headerLength != 68 || totalLength < 76)
            throw new InvalidDataException("A DOC picture has an invalid PICF header.");
        var block = structure.ReadRange("Data", dataOffset, totalLength);
        var goalWidth = BinaryPrimitives.ReadInt16LittleEndian(block.AsSpan(28));
        var goalHeight = BinaryPrimitives.ReadInt16LittleEndian(block.AsSpan(30));
        var scaleX = BinaryPrimitives.ReadUInt16LittleEndian(block.AsSpan(32));
        var scaleY = BinaryPrimitives.ReadUInt16LittleEndian(block.AsSpan(34));
        if (goalWidth <= 0 || goalHeight <= 0 || scaleX == 0 || scaleY == 0)
            throw new InvalidDataException("A DOC picture has invalid display dimensions.");
        var width = checked((long)Math.Round(goalWidth * (decimal)scaleX * 635 / 1000));
        var height = checked((long)Math.Round(goalHeight * (decimal)scaleY * 635 / 1000));
        var metafileType = BinaryPrimitives.ReadUInt16LittleEndian(block.AsSpan(6));
        var recordsStart = 68;
        if (metafileType == 0x66)
            recordsStart = checked(recordsStart + 1 + block[68]);
        else if (metafileType != 0x64)
            throw new NotSupportedException($"DOC picture format 0x{metafileType:X4} is unsupported.");
        var found = FindBlip(block.AsSpan(recordsStart));
        var records = block.AsSpan(recordsStart);
        var linkedImage = found == null ? FindLinkedImage(records) : null;
        if (found == null && linkedImage == null) return null;
        var flips = FindFlips(records);
        return new DocInlinePicture(found?.Bytes ?? [], found?.ContentType ?? "image/png",
            width, height, FindCrop(records), flips.Horizontal, flips.Vertical,
            FindRotation(records), linkedImage);
    }

    private static Uri? FindLinkedImage(ReadOnlySpan<byte> records)
    {
        for (var offset = 0; offset + 8 <= records.Length;)
        {
            var word = BinaryPrimitives.ReadUInt16LittleEndian(records.Slice(offset));
            var version = word & 0xF;
            var instance = word >> 4;
            var type = BinaryPrimitives.ReadUInt16LittleEndian(records.Slice(offset + 2));
            var length = BinaryPrimitives.ReadUInt32LittleEndian(records.Slice(offset + 4));
            if (length > records.Length - offset - 8)
                throw new InvalidDataException("An OfficeArt shape record exceeds its container.");
            var body = records.Slice(offset + 8, checked((int)length));
            if (type == 0xF00B)
            {
                var tableLength = checked(instance * 6);
                if (tableLength > body.Length)
                    throw new InvalidDataException("An OfficeArt shape property table is truncated.");
                var dataOffset = tableLength;
                for (var i = 0; i < instance; i++)
                {
                    var entry = body.Slice(i * 6);
                    var property = BinaryPrimitives.ReadUInt16LittleEndian(entry);
                    var value = BinaryPrimitives.ReadUInt32LittleEndian(entry.Slice(2));
                    if ((property & 0x8000) == 0) continue;
                    if (value > body.Length - dataOffset)
                        throw new InvalidDataException("An OfficeArt complex property is truncated.");
                    if (property == 0xC105 && value >= 4 && value % 2 == 0)
                    {
                        // pibName is the UTF-16 path of a linked picture; it has no embedded BLIP.
                        var name = System.Text.Encoding.Unicode.GetString(
                            body.Slice(dataOffset, checked((int)value)).ToArray()).TrimEnd('\0');
                        if (Uri.TryCreate(name, UriKind.Absolute, out var uri))
                            return uri;
                    }
                    dataOffset += checked((int)value);
                }
            }
            if (version == 0xF && FindLinkedImage(body) is { } nested)
                return nested;
            offset += checked(8 + (int)length);
        }
        return null;
    }

    private static double FindRotation(ReadOnlySpan<byte> records)
    {
        for (var offset = 0; offset + 8 <= records.Length;)
        {
            var word = BinaryPrimitives.ReadUInt16LittleEndian(records.Slice(offset));
            var version = word & 0xF;
            var instance = word >> 4;
            var type = BinaryPrimitives.ReadUInt16LittleEndian(records.Slice(offset + 2));
            var length = BinaryPrimitives.ReadUInt32LittleEndian(records.Slice(offset + 4));
            if (length > records.Length - offset - 8)
                throw new InvalidDataException("An OfficeArt shape record exceeds its container.");
            var body = records.Slice(offset + 8, checked((int)length));
            if (type == 0xF00B)
            {
                if (instance * 6 > body.Length)
                    throw new InvalidDataException("An OfficeArt shape property table is truncated.");
                for (var i = 0; i < instance; i++)
                {
                    var entry = body.Slice(i * 6);
                    if (BinaryPrimitives.ReadUInt16LittleEndian(entry) == 0x0004)
                        return BinaryPrimitives.ReadInt32LittleEndian(entry.Slice(2)) /
                            65536.0;
                }
            }
            if (version == 0xF && FindRotation(body) is var nested && nested != 0)
                return nested;
            offset += checked(8 + (int)length);
        }
        return 0;
    }

    private static (bool Horizontal, bool Vertical) FindFlips(ReadOnlySpan<byte> records)
    {
        for (var offset = 0; offset + 8 <= records.Length;)
        {
            var word = BinaryPrimitives.ReadUInt16LittleEndian(records.Slice(offset));
            var version = word & 0xF;
            var type = BinaryPrimitives.ReadUInt16LittleEndian(records.Slice(offset + 2));
            var length = BinaryPrimitives.ReadUInt32LittleEndian(records.Slice(offset + 4));
            if (length > records.Length - offset - 8)
                throw new InvalidDataException("An OfficeArt shape record exceeds its container.");
            var body = records.Slice(offset + 8, checked((int)length));
            if (type == 0xF00A && body.Length == 8)
            {
                var flags = BinaryPrimitives.ReadUInt32LittleEndian(body.Slice(4));
                return ((flags & 0x40) != 0, (flags & 0x80) != 0);
            }
            if (version == 0xF && FindFlips(body) is { } nested &&
                (nested.Horizontal || nested.Vertical))
                return nested;
            offset += checked(8 + (int)length);
        }
        return (false, false);
    }

    private static DocInlinePictureCrop? FindCrop(ReadOnlySpan<byte> records)
    {
        for (var offset = 0; offset + 8 <= records.Length;)
        {
            var word = BinaryPrimitives.ReadUInt16LittleEndian(records.Slice(offset));
            var version = word & 0xF;
            var instance = word >> 4;
            var type = BinaryPrimitives.ReadUInt16LittleEndian(records.Slice(offset + 2));
            var length = BinaryPrimitives.ReadUInt32LittleEndian(records.Slice(offset + 4));
            if (length > records.Length - offset - 8)
                throw new InvalidDataException("An OfficeArt shape record exceeds its container.");
            var body = records.Slice(offset + 8, checked((int)length));
            if (type == 0xF00B)
            {
                if (instance * 6 > body.Length)
                    throw new InvalidDataException("An OfficeArt shape property table is truncated.");
                double left = 0, top = 0, right = 0, bottom = 0;
                for (var i = 0; i < instance; i++)
                {
                    var entry = body.Slice(i * 6);
                    var property = BinaryPrimitives.ReadUInt16LittleEndian(entry);
                    var fraction = BinaryPrimitives.ReadInt32LittleEndian(entry.Slice(2)) /
                        65536.0;
                    switch (property)
                    {
                        case 0x0100: top = fraction; break;
                        case 0x0101: bottom = fraction; break;
                        case 0x0102: left = fraction; break;
                        case 0x0103: right = fraction; break;
                    }
                }
                var crop = new DocInlinePictureCrop(left, top, right, bottom);
                if (!crop.IsEmpty) return crop;
            }
            if (version == 0xF && FindCrop(body) is { } nested)
                return nested;
            offset += checked(8 + (int)length);
        }
        return null;
    }

    private static (byte[] Bytes, string ContentType)? FindBlip(ReadOnlySpan<byte> records)
    {
        for (var offset = 0; offset + 8 <= records.Length;)
        {
            var word = BinaryPrimitives.ReadUInt16LittleEndian(records.Slice(offset));
            var version = word & 0xF;
            var instance = word >> 4;
            var type = BinaryPrimitives.ReadUInt16LittleEndian(records.Slice(offset + 2));
            var length = BinaryPrimitives.ReadUInt32LittleEndian(records.Slice(offset + 4));
            if (length > records.Length - offset - 8)
                throw new InvalidDataException("An OfficeArt picture record exceeds its container.");
            var body = records.Slice(offset + 8, checked((int)length));
            if (type is 0xF01D or 0xF01E or 0xF01F or 0xF029 or 0xF02A)
            {
                var prefix = instance switch
                {
                    0x46A or 0x6E0 or 0x6E2 or 0x6E4 or 0x7A8 => 17,
                    0x46B or 0x6E1 or 0x6E3 or 0x6E5 or 0x7A9 => 33,
                    _ => throw new NotSupportedException(
                        $"OfficeArt BLIP instance 0x{instance:X3} is unsupported.")
                };
                if (body.Length <= prefix)
                    throw new InvalidDataException("An OfficeArt image BLIP is truncated.");
                var data = body.Slice(prefix);
                if (type == 0xF01F)
                    return (DocDibBitmap.ToBmp(data), "image/bmp");
                var isJpeg = type is 0xF01D or 0xF02A;
                var isTiff = type == 0xF029;
                if (isTiff ? !DocBinaryCompat.IsTiff(data) :
                    isJpeg ? !(data[0] == 0xFF && data[1] == 0xD8) :
                    !(data.Length >= 8 && data[0] == 0x89 &&
                        data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47))
                    throw new InvalidDataException("The OfficeArt image bytes do not match their BLIP type.");
                return (data.ToArray(), isTiff ? "image/tiff" : isJpeg ? "image/jpeg" : "image/png");
            }
            if (version == 0xF)
            {
                var nested = FindBlip(body);
                if (nested != null) return nested;
            }
            else if (type == 0xF007 && body.Length >= 36)
            {
                var nameLength = body[33];
                if (36 + nameLength < body.Length)
                {
                    var nested = FindBlip(body.Slice(36 + nameLength));
                    if (nested != null) return nested;
                }
            }
            offset += checked(8 + (int)length);
        }
        return null;
    }
}
