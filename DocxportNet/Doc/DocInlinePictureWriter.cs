using System.Buffers.Binary;

namespace DocxportNet.Doc;

/// <summary>Writes the PICF and embedded OfficeArt BLIP used by inline DOC pictures.</summary>
internal static class DocInlinePictureWriter
{
    public static int Write(Stream data, DocInlinePicture picture, uint shapeId)
    {
        if (picture.ContentType is not ("image/jpeg" or "image/png" or "image/tiff" or "image/bmp"))
            throw new NotSupportedException($"DOC inline image type '{picture.ContentType}' is unsupported.");
        if (picture.WidthEmu <= 0 || picture.HeightEmu <= 0)
            throw new InvalidDataException("A DOC inline image needs positive display dimensions.");
        var widthTwips = checked((short)Math.Round(picture.WidthEmu / 635m));
        var heightTwips = checked((short)Math.Round(picture.HeightEmu / 635m));
        if (widthTwips <= 0 || heightTwips <= 0)
            throw new NotSupportedException("A DOC inline image exceeds the PICF size range.");
        var jpeg = picture.ContentType == "image/jpeg";
        var tiff = picture.ContentType == "image/tiff";
        var bmp = picture.ContentType == "image/bmp";
        var bytes = picture.Bytes;
        if (!bmp && (tiff ? !DocBinaryCompat.IsTiff(bytes) : jpeg ?
            bytes.Length < 3 || bytes[0] != 0xFF || bytes[1] != 0xD8 :
            bytes.Length < 8 || bytes[0] != 0x89 || bytes[1] != 0x50 ||
            bytes[2] != 0x4E || bytes[3] != 0x47))
            throw new InvalidDataException("The inline image bytes do not match their content type.");
        var blipBytes = bmp ? DocDibBitmap.ToDib(bytes) : bytes;

        var offset = checked((int)data.Position);
        var uid = DocBinaryCompat.Md5(bytes);
        using var records = new MemoryStream();
        WriteShape(records, shapeId, picture.Crop, picture.FlipHorizontal,
            picture.FlipVertical, picture.RotationDegrees);
        using var blip = new MemoryStream();
        DocBinaryCompat.Write(blip, uid);
        blip.WriteByte(0xFF); // BLIP tag.
        DocBinaryCompat.Write(blip, blipBytes);
        var blipType = bmp ? (ushort)0xF01F : tiff ? (ushort)0xF029 :
            jpeg ? (ushort)0xF01D : (ushort)0xF01E;
        var blipInstance = bmp ? (ushort)0x7A8 : tiff ? (ushort)0x6E4 :
            jpeg ? (ushort)0x46A : (ushort)0x6E0;
        using var blipRecord = new MemoryStream();
        WriteRecord(blipRecord, 0, blipInstance, blipType, blip.ToArray());
        using var fbse = new MemoryStream();
        var imageType = bmp ? (byte)7 : tiff ? (byte)0x11 : jpeg ? (byte)5 : (byte)6;
        fbse.WriteByte(imageType); // btWin32
        fbse.WriteByte(imageType); // btMacOS
        DocBinaryCompat.Write(fbse, uid);
        WriteU16(fbse, 0x00FF);
        WriteU32(fbse, checked((uint)blipRecord.Length));
        WriteU32(fbse, 1); // cRef
        WriteU32(fbse, checked((uint)(offset + 68))); // foDelay
        WriteU32(fbse, 0); // usage, name length, and unused bytes
        DocBinaryCompat.Write(fbse, blipRecord.ToArray());
        WriteRecord(records, 2, (ushort)imageType, 0xF007, fbse.ToArray());

        var pictureRecord = records.ToArray();
        var picf = new byte[68];
        BinaryPrimitives.WriteInt32LittleEndian(picf, checked(68 + pictureRecord.Length));
        BinaryPrimitives.WriteUInt16LittleEndian(picf.AsSpan(4), 68);
        BinaryPrimitives.WriteUInt16LittleEndian(picf.AsSpan(6), 0x64);
        BinaryPrimitives.WriteUInt16LittleEndian(picf.AsSpan(14), 8);
        BinaryPrimitives.WriteInt16LittleEndian(picf.AsSpan(28), widthTwips);
        BinaryPrimitives.WriteInt16LittleEndian(picf.AsSpan(30), heightTwips);
        BinaryPrimitives.WriteUInt16LittleEndian(picf.AsSpan(32), 1000);
        BinaryPrimitives.WriteUInt16LittleEndian(picf.AsSpan(34), 1000);
        DocBinaryCompat.Write(data, picf);
        DocBinaryCompat.Write(data, pictureRecord);
        return offset;
    }

    private static void WriteShape(Stream output, uint shapeId,
        DocInlinePictureCrop? crop, bool flipHorizontal, bool flipVertical,
        double rotationDegrees)
    {
        using var shape = new MemoryStream();
        var sp = new byte[8];
        BinaryPrimitives.WriteUInt32LittleEndian(sp, shapeId);
        BinaryPrimitives.WriteUInt32LittleEndian(sp.AsSpan(4), 0x00000A00U |
            (flipHorizontal ? 0x40U : 0U) | (flipVertical ? 0x80U : 0U));
        WriteRecord(shape, 2, 0x4B, 0xF00A, sp);
        // The shape properties identify the picture BLIP and allow Word to
        // render it from the following embedded BLIP store entry.
        byte[] properties = [
            0x04, 0x41, 0x01, 0x00, 0x00, 0x00,
            0x3F, 0x01, 0x00, 0x00, 0x06, 0x00,
            0xBF, 0x01, 0x00, 0x00, 0x10, 0x00,
            0xFF, 0x01, 0x00, 0x00, 0x08, 0x00,
            0x80, 0xC3, 0x14, 0x00, 0x00, 0x00,
            0xBF, 0x03, 0x00, 0x00, 0x02, 0x00,
            0x50, 0x00, 0x69, 0x00, 0x63, 0x00, 0x74, 0x00,
            0x75, 0x00, 0x72, 0x00, 0x65, 0x00, 0x20, 0x00,
            0x31, 0x00, 0x00, 0x00];
        var extraProperties = new List<byte>();
        if (rotationDegrees != 0)
        {
            if (!DocBinaryCompat.IsFinite(rotationDegrees))
                throw new NotSupportedException("DOC inline picture rotation is invalid.");
            Span<byte> entry = stackalloc byte[6];
            BinaryPrimitives.WriteUInt16LittleEndian(entry, 0x0004);
            BinaryPrimitives.WriteInt32LittleEndian(entry.Slice(2),
                checked((int)Math.Round(rotationDegrees * 65536)));
            extraProperties.AddRange(entry.ToArray());
        }
        if (crop is { IsEmpty: false })
        {
            if (!DocBinaryCompat.IsFinite(crop.Left) || !DocBinaryCompat.IsFinite(crop.Top) ||
                !DocBinaryCompat.IsFinite(crop.Right) || !DocBinaryCompat.IsFinite(crop.Bottom) ||
                crop.Left < 0 || crop.Top < 0 || crop.Right < 0 || crop.Bottom < 0 ||
                crop.Left + crop.Right >= 1 || crop.Top + crop.Bottom >= 1)
                throw new NotSupportedException("DOC inline image crop fractions are invalid.");
            void AddCrop(ushort property, double fraction)
            {
                if (fraction == 0) return;
                Span<byte> entry = stackalloc byte[6];
                BinaryPrimitives.WriteUInt16LittleEndian(entry, property);
                BinaryPrimitives.WriteInt32LittleEndian(entry.Slice(2),
                    checked((int)Math.Round(fraction * 65536)));
                extraProperties.AddRange(entry.ToArray());
            }
            AddCrop(0x0100, crop.Top);
            AddCrop(0x0101, crop.Bottom);
            AddCrop(0x0102, crop.Left);
            AddCrop(0x0103, crop.Right);
        }
        var allProperties = extraProperties.Concat(properties).ToArray();
        WriteRecord(shape, 3, checked((ushort)(6 + extraProperties.Count / 6)),
            0xF00B, allProperties);
        WriteRecord(shape, 3, 1, 0xF122, [0xAA, 0x03, 0, 0, 0, 0x0F]);
        WriteRecord(shape, 0, 0, 0xF010, [0, 0, 0, 0x80]);
        WriteRecord(output, 15, 0, 0xF004, shape.ToArray());
    }

    private static void WriteRecord(Stream output, ushort version, ushort instance,
        ushort type, byte[] body)
    {
        WriteU16(output, checked((ushort)((instance << 4) | version)));
        WriteU16(output, type);
        WriteU32(output, checked((uint)body.Length));
        DocBinaryCompat.Write(output, body);
    }

    private static void WriteU16(Stream output, ushort value)
    {
        Span<byte> buffer = stackalloc byte[2];
        BinaryPrimitives.WriteUInt16LittleEndian(buffer, value);
        DocBinaryCompat.Write(output, buffer);
    }

    private static void WriteU32(Stream output, uint value)
    {
        Span<byte> buffer = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32LittleEndian(buffer, value);
        DocBinaryCompat.Write(output, buffer);
    }
}
