using System.Buffers.Binary;

namespace DocxportNet.Doc;

/// <summary>Converts canonical uncompressed BMP files to OfficeArt DIB data.</summary>
internal static class DocDibBitmap
{
    public static byte[] ToDib(byte[] bitmap)
    {
        if (bitmap.Length < 54 || bitmap[0] != 'B' || bitmap[1] != 'M' ||
            BinaryPrimitives.ReadUInt32LittleEndian(bitmap.AsSpan(2)) != bitmap.Length)
            throw new NotSupportedException("DOC BMP pictures require a canonical BMP file header.");
        var dib = bitmap.AsSpan(14);
        var pixelOffset = Validate(dib);
        if (BinaryPrimitives.ReadUInt32LittleEndian(bitmap.AsSpan(10)) !=
            checked((uint)(14 + pixelOffset)))
            throw new NotSupportedException("DOC BMP pictures require a canonical pixel offset.");
        if (BinaryPrimitives.ReadUInt32LittleEndian(dib) != 40)
            return NormalizeExtendedHeader(dib, pixelOffset);
        return BinaryPrimitives.ReadUInt32LittleEndian(dib.Slice(16)) == 2
            ? ExpandRle4(dib, pixelOffset) : dib.ToArray();
    }

    public static byte[] ToBmp(ReadOnlySpan<byte> dib)
    {
        var pixelOffset = Validate(dib);
        var bitmap = new byte[checked(dib.Length + 14)];
        bitmap[0] = (byte)'B';
        bitmap[1] = (byte)'M';
        BinaryPrimitives.WriteUInt32LittleEndian(bitmap.AsSpan(2),
            checked((uint)bitmap.Length));
        BinaryPrimitives.WriteUInt32LittleEndian(bitmap.AsSpan(10),
            checked((uint)(14 + pixelOffset)));
        dib.CopyTo(bitmap.AsSpan(14));
        return bitmap;
    }

    private static int Validate(ReadOnlySpan<byte> dib)
    {
        if (dib.Length < 40 ||
            BinaryPrimitives.ReadUInt32LittleEndian(dib) is not (40 or 108 or 124) ||
            BinaryPrimitives.ReadUInt16LittleEndian(dib.Slice(12)) != 1 ||
            BinaryPrimitives.ReadUInt16LittleEndian(dib.Slice(14)) is not (1 or 4 or 8 or 16 or 24 or 32))
            throw new NotSupportedException("DOC BMP pictures require indexed or 16/24/32-bit DIB data.");
        var headerSize = checked((int)BinaryPrimitives.ReadUInt32LittleEndian(dib));
        if (dib.Length < headerSize)
            throw new InvalidDataException("A DOC BMP picture has a truncated DIB header.");
        var width = BinaryPrimitives.ReadInt32LittleEndian(dib.Slice(4));
        var height = BinaryPrimitives.ReadInt32LittleEndian(dib.Slice(8));
        if (width <= 0 || height is 0 or int.MinValue)
            throw new InvalidDataException("A DOC BMP picture has invalid dimensions.");
        var bits = BinaryPrimitives.ReadUInt16LittleEndian(dib.Slice(14));
        var compression = BinaryPrimitives.ReadUInt32LittleEndian(dib.Slice(16));
        if (compression != 0 &&
            !(compression == 1 && bits == 8 || compression == 2 && bits == 4 ||
                compression == 3 && bits is 16 or 32))
            throw new NotSupportedException("The DOC BMP compression mode is unsupported.");
        if (compression is 1 or 2 && height < 0)
            throw new NotSupportedException("DOC RLE bitmaps require bottom-up rows.");
        var rowBytes = checked(((long)width * bits + 31) / 32 * 4);
        var pixelBytes = checked(rowBytes * Math.Abs((long)height));
        var usedColors = BinaryPrimitives.ReadUInt32LittleEndian(dib.Slice(32));
        if (bits > 8 && usedColors != 0 ||
            bits <= 8 && usedColors > (1u << bits))
            throw new NotSupportedException("DOC BMP pictures require a canonical color table.");
        var paletteColors = bits <= 8
            ? usedColors == 0 ? 1u << bits : usedColors : 0u;
        var pixelOffset = checked(headerSize + (int)paletteColors * 4 +
            (compression == 3 && headerSize == 40 ? 12 : 0));
        if (compression is 1 or 2)
        {
            var encodedSize = BinaryPrimitives.ReadUInt32LittleEndian(dib.Slice(20));
            if (encodedSize < 2 || encodedSize != dib.Length - pixelOffset ||
                dib[dib.Length - 2] != 0 || dib[dib.Length - 1] != 1)
                throw new NotSupportedException("DOC RLE bitmaps require a canonical encoded pixel array.");
        }
        else if (pixelBytes != dib.Length - pixelOffset)
            throw new NotSupportedException("DOC BMP pictures require a canonical pixel array.");
        if (compression == 3)
        {
            var red = BinaryPrimitives.ReadUInt32LittleEndian(dib.Slice(40));
            var green = BinaryPrimitives.ReadUInt32LittleEndian(dib.Slice(44));
            var blue = BinaryPrimitives.ReadUInt32LittleEndian(dib.Slice(48));
            var allowed = bits == 16 ? 0xFFFFu : uint.MaxValue;
            if (red == 0 || green == 0 || blue == 0 ||
                ((red | green | blue) & ~allowed) != 0 ||
                (red & green) != 0 || (red & blue) != 0 || (green & blue) != 0 ||
                !Contiguous(red) || !Contiguous(green) || !Contiguous(blue))
                throw new NotSupportedException("DOC BMP color masks are invalid.");
        }
        return pixelOffset;
    }

    private static bool Contiguous(uint mask)
    {
        while ((mask & 1) == 0) mask >>= 1;
        return (mask & (mask + 1)) == 0;
    }

    private static byte[] NormalizeExtendedHeader(ReadOnlySpan<byte> dib,
        int pixelOffset)
    {
        var headerSize = checked((int)BinaryPrimitives.ReadUInt32LittleEndian(dib));
        var bits = BinaryPrimitives.ReadUInt16LittleEndian(dib.Slice(14));
        var compression = BinaryPrimitives.ReadUInt32LittleEndian(dib.Slice(16));
        if (bits is not (24 or 32) || compression is not (0 or 3) ||
            BinaryPrimitives.ReadUInt32LittleEndian(dib.Slice(56)) != 0x73524742 ||
            headerSize == 124 && BinaryPrimitives.ReadUInt32LittleEndian(dib.Slice(116)) != 0)
            throw new NotSupportedException("DOC BMP pictures require an opaque sRGB V4/V5 bitmap.");
        var alphaMask = BinaryPrimitives.ReadUInt32LittleEndian(dib.Slice(52));
        if (alphaMask != 0)
        {
            if (bits != 32 || alphaMask != 0xFF000000)
                throw new NotSupportedException("DOC BMP pictures require opaque V4/V5 pixels.");
            for (var at = pixelOffset; at < dib.Length; at += 4)
                if (dib[at + 3] != 0xFF)
                    throw new NotSupportedException("DOC BMP pictures require opaque V4/V5 pixels.");
        }
        var newOffset = compression == 3 ? 52 : 40;
        var output = new byte[checked(newOffset + dib.Length - pixelOffset)];
        dib.Slice(0, 40).CopyTo(output);
        BinaryPrimitives.WriteUInt32LittleEndian(output, 40);
        if (compression == 3) dib.Slice(40, 12).CopyTo(output.AsSpan(40));
        dib.Slice(pixelOffset).CopyTo(output.AsSpan(newOffset));
        return output;
    }

    private static byte[] ExpandRle4(ReadOnlySpan<byte> dib, int pixelOffset)
    {
        var width = BinaryPrimitives.ReadInt32LittleEndian(dib.Slice(4));
        var height = BinaryPrimitives.ReadInt32LittleEndian(dib.Slice(8));
        var stride = checked((width + 7) / 8 * 4);
        var output = new byte[checked(pixelOffset + stride * height)];
        dib.Slice(0, pixelOffset).CopyTo(output);
        BinaryPrimitives.WriteUInt32LittleEndian(output.AsSpan(16), 0);
        BinaryPrimitives.WriteUInt32LittleEndian(output.AsSpan(20),
            checked((uint)(stride * height)));
        var x = 0;
        var y = 0;
        var at = pixelOffset;
        while (at < dib.Length)
        {
            if (at + 2 > dib.Length)
                throw new InvalidDataException("A DOC RLE4 bitmap has a truncated packet.");
            var count = dib[at++];
            var value = dib[at++];
            if (count != 0)
            {
                for (var i = 0; i < count; i++)
                    Put(i % 2 == 0 ? value >> 4 : value & 15);
                continue;
            }
            if (value == 0) { x = 0; y++; continue; }
            if (value == 1)
            {
                if (at != dib.Length)
                    throw new InvalidDataException("A DOC RLE4 bitmap has data after its end marker.");
                return output;
            }
            if (value == 2)
            {
                if (at + 2 > dib.Length)
                    throw new InvalidDataException("A DOC RLE4 bitmap has a truncated delta.");
                x += dib[at++];
                y += dib[at++];
                continue;
            }
            var packetBytes = (value + 1) / 2;
            if (at + packetBytes + packetBytes % 2 > dib.Length)
                throw new InvalidDataException("A DOC RLE4 bitmap has a truncated absolute run.");
            for (var i = 0; i < value; i++)
            {
                var packed = dib[at + i / 2];
                Put(i % 2 == 0 ? packed >> 4 : packed & 15);
            }
            at += packetBytes + packetBytes % 2;
        }
        throw new InvalidDataException("A DOC RLE4 bitmap has no end marker.");

        void Put(int color)
        {
            if (x >= width || y >= height)
                throw new InvalidDataException("A DOC RLE4 bitmap exceeds its dimensions.");
            var index = pixelOffset + y * stride + x / 2;
            if (x % 2 == 0) output[index] |= checked((byte)(color << 4));
            else output[index] |= checked((byte)color);
            x++;
        }
    }
}
