using System.Buffers.Binary;
using DocxportNet.Doc;

namespace DocxportNet.Visitors.Pdf;

internal static class DxpBmp24Converter
{
    public static byte[] TransformOrOriginal(byte[] bitmap, double rotationDegrees,
        bool flipHorizontal, bool flipVertical)
    {
        try { return Transform(bitmap, rotationDegrees, flipHorizontal, flipVertical); }
        catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or
            OverflowException or ArgumentException or IndexOutOfRangeException)
        {
            return bitmap;
        }
    }

    public static byte[] Transform(byte[] bitmap, double rotationDegrees, bool flipHorizontal, bool flipVertical)
    {
        var sourceTurns = (int)Math.Round(rotationDegrees / 90.0) % 4;
        var quarterTurns = (4 - sourceTurns) % 4;
        if (Math.Abs(rotationDegrees - sourceTurns * 90.0) > 0.001 ||
            quarterTurns == 0 && !flipHorizontal && !flipVertical)
            return bitmap;
        if (bitmap.Length < 54 || bitmap[0] != 'B' || bitmap[1] != 'M' ||
            BinaryPrimitives.ReadUInt16LittleEndian(bitmap.AsSpan(28)) != 24)
            return bitmap;
        var width = BinaryPrimitives.ReadInt32LittleEndian(bitmap.AsSpan(18));
        var signedHeight = BinaryPrimitives.ReadInt32LittleEndian(bitmap.AsSpan(22));
        if (width <= 0 || signedHeight is 0 or int.MinValue)
            throw new InvalidDataException("BMP transform dimensions are invalid.");
        var height = Math.Abs(signedHeight);
        var sourceStride = checked((width * 3 + 3) / 4 * 4);
        if (bitmap.Length < checked(54 + sourceStride * height))
            throw new InvalidDataException("BMP transform pixels are truncated.");
        var targetWidth = quarterTurns % 2 == 0 ? width : height;
        var targetHeight = quarterTurns % 2 == 0 ? height : width;
        var targetStride = checked((targetWidth * 3 + 3) / 4 * 4);
        var output = new byte[checked(54 + targetStride * targetHeight)];
        bitmap.AsSpan(0, 54).CopyTo(output);
        BinaryPrimitives.WriteInt32LittleEndian(output.AsSpan(2), output.Length);
        BinaryPrimitives.WriteInt32LittleEndian(output.AsSpan(18), targetWidth);
        BinaryPrimitives.WriteInt32LittleEndian(output.AsSpan(22), -targetHeight);
        BinaryPrimitives.WriteInt32LittleEndian(output.AsSpan(34), targetStride * targetHeight);
        for (var row = 0; row < height; row++)
            for (var column = 0; column < width; column++)
            {
                var horizontal = flipHorizontal ? width - 1 - column : column;
                var vertical = flipVertical ? height - 1 - row : row;
                var targetColumn = quarterTurns switch
                {
                    1 => height - 1 - vertical,
                    2 => width - 1 - horizontal,
                    3 => vertical,
                    _ => horizontal
                };
                var targetRow = quarterTurns switch
                {
                    1 => horizontal,
                    2 => height - 1 - vertical,
                    3 => width - 1 - horizontal,
                    _ => vertical
                };
                var source = 54 + (signedHeight < 0 ? row : height - 1 - row) * sourceStride + column * 3;
                var target = 54 + targetRow * targetStride + targetColumn * 3;
                bitmap.AsSpan(source, 3).CopyTo(output.AsSpan(target, 3));
            }
        return output;
    }

    public static byte[] NormalizeOrOriginal(byte[] bitmap)
    {
        try { return Normalize(bitmap); }
        catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or
            OverflowException or ArgumentException)
        {
            return bitmap;
        }
    }

    public static byte[] Normalize(byte[] bitmap)
    {
        var dib = DocDibBitmap.ToDib(bitmap);
        var width = BinaryPrimitives.ReadInt32LittleEndian(dib.AsSpan(4));
        var signedHeight = BinaryPrimitives.ReadInt32LittleEndian(dib.AsSpan(8));
        var height = Math.Abs(signedHeight);
        var bits = BinaryPrimitives.ReadUInt16LittleEndian(dib.AsSpan(14));
        var compression = BinaryPrimitives.ReadUInt32LittleEndian(dib.AsSpan(16));
        var colorsUsed = BinaryPrimitives.ReadUInt32LittleEndian(dib.AsSpan(32));
        var paletteSize = bits <= 8 ? checked((int)(colorsUsed == 0 ? 1u << bits : colorsUsed)) : 0;
        var pixelOffset = 40 + paletteSize * 4 + (compression == 3 ? 12 : 0);
        var sourceStride = checked((width * bits + 31) / 32 * 4);
        var destinationStride = checked((width * 3 + 3) / 4 * 4);
        var output = new byte[checked(54 + destinationStride * height)];
        output[0] = (byte)'B';
        output[1] = (byte)'M';
        BinaryPrimitives.WriteInt32LittleEndian(output.AsSpan(2), output.Length);
        BinaryPrimitives.WriteInt32LittleEndian(output.AsSpan(10), 54);
        BinaryPrimitives.WriteInt32LittleEndian(output.AsSpan(14), 40);
        BinaryPrimitives.WriteInt32LittleEndian(output.AsSpan(18), width);
        BinaryPrimitives.WriteInt32LittleEndian(output.AsSpan(22), signedHeight);
        BinaryPrimitives.WriteUInt16LittleEndian(output.AsSpan(26), 1);
        BinaryPrimitives.WriteUInt16LittleEndian(output.AsSpan(28), 24);
        BinaryPrimitives.WriteInt32LittleEndian(output.AsSpan(34), destinationStride * height);
        var rlePixels = compression == 1 ? DecodeRle8(dib, pixelOffset, width, height) : null;
        var rle4Presence = BinaryPrimitives.ReadUInt32LittleEndian(bitmap.AsSpan(30)) == 2
            ? DecodeRle4Presence(bitmap, BinaryPrimitives.ReadInt32LittleEndian(bitmap.AsSpan(10)), width, height)
            : null;
        var redMask = compression == 3 ? BinaryPrimitives.ReadUInt32LittleEndian(dib.AsSpan(40)) : bits == 16 ? 0x7C00u : 0x00FF0000u;
        var greenMask = compression == 3 ? BinaryPrimitives.ReadUInt32LittleEndian(dib.AsSpan(44)) : bits == 16 ? 0x03E0u : 0x0000FF00u;
        var blueMask = compression == 3 ? BinaryPrimitives.ReadUInt32LittleEndian(dib.AsSpan(48)) : bits == 16 ? 0x001Fu : 0x000000FFu;
        for (var row = 0; row < height; row++)
        {
            for (var column = 0; column < width; column++)
            {
                var source = pixelOffset + row * sourceStride;
                var target = 54 + row * destinationStride + column * 3;
                if (bits <= 8)
                {
                    if (rle4Presence != null && !rle4Presence[row * width + column]) continue;
                    var index = rlePixels != null ? rlePixels[row * width + column] :
                        bits == 8 ? dib[source + column] :
                        bits == 4 ? (dib[source + column / 2] >> (column % 2 == 0 ? 4 : 0)) & 15 :
                        (dib[source + column / 8] >> (7 - column % 8)) & 1;
                    if (index >= paletteSize) throw new InvalidDataException("BMP palette index exceeds its color table.");
                    dib.AsSpan(40 + index * 4, 3).CopyTo(output.AsSpan(target, 3));
                }
                else if (bits == 24)
                    dib.AsSpan(source + column * 3, 3).CopyTo(output.AsSpan(target, 3));
                else
                {
                    var value = bits == 16 ? BinaryPrimitives.ReadUInt16LittleEndian(dib.AsSpan(source + column * 2)) :
                        BinaryPrimitives.ReadUInt32LittleEndian(dib.AsSpan(source + column * 4));
                    output[target] = Channel(value, blueMask);
                    output[target + 1] = Channel(value, greenMask);
                    output[target + 2] = Channel(value, redMask);
                }
            }
        }
        return output;
    }

    private static byte Channel(uint value, uint mask)
    {
        var shift = 0;
        while ((mask & 1) == 0) { mask >>= 1; shift++; }
        return (byte)(((value >> shift & mask) * 255 + mask / 2) / mask);
    }

    private static byte[] DecodeRle8(byte[] dib, int offset, int width, int height)
    {
        var pixels = new byte[checked(width * height)];
        var column = 0;
        var row = 0;
        while (offset + 1 < dib.Length)
        {
            var count = dib[offset++];
            var value = dib[offset++];
            if (count > 0)
            {
                for (var index = 0; index < count; index++) Put(value);
            }
            else if (value == 0) { column = 0; row++; }
            else if (value == 1) return pixels;
            else if (value == 2)
            {
                if (offset + 1 >= dib.Length) throw new InvalidDataException("Truncated BMP RLE delta.");
                column += dib[offset++];
                row += dib[offset++];
            }
            else
            {
                if (offset + value + value % 2 > dib.Length) throw new InvalidDataException("Truncated BMP RLE run.");
                for (var index = 0; index < value; index++) Put(dib[offset + index]);
                offset += value + value % 2;
            }
        }
        throw new InvalidDataException("BMP RLE stream has no end marker.");

        void Put(byte color)
        {
            if (column >= width || row >= height) throw new InvalidDataException("BMP RLE pixel exceeds its bounds.");
            pixels[row * width + column++] = color;
        }
    }

    private static bool[] DecodeRle4Presence(byte[] bitmap, int offset, int width, int height)
    {
        var present = new bool[checked(width * height)];
        var column = 0;
        var row = 0;
        while (offset + 1 < bitmap.Length)
        {
            var count = bitmap[offset++];
            var value = bitmap[offset++];
            if (count > 0)
            {
                for (var index = 0; index < count; index++) Mark();
            }
            else if (value == 0) { column = 0; row++; }
            else if (value == 1) return present;
            else if (value == 2)
            {
                if (offset + 1 >= bitmap.Length) throw new InvalidDataException("Truncated BMP RLE delta.");
                column += bitmap[offset++];
                row += bitmap[offset++];
            }
            else
            {
                var packetBytes = (value + 1) / 2;
                if (offset + packetBytes + packetBytes % 2 > bitmap.Length)
                    throw new InvalidDataException("Truncated BMP RLE run.");
                for (var index = 0; index < value; index++) Mark();
                offset += packetBytes + packetBytes % 2;
            }
        }
        throw new InvalidDataException("BMP RLE stream has no end marker.");

        void Mark()
        {
            if (column >= width || row >= height) throw new InvalidDataException("BMP RLE pixel exceeds its bounds.");
            present[row * width + column++] = true;
        }
    }
}
