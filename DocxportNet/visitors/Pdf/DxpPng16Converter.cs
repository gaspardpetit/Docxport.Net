using System.Buffers.Binary;
using System.IO.Compression;

namespace DocxportNet.Visitors.Pdf;

internal static class DxpPng16Converter
{
    public static byte[] ToBmpOrOriginal(byte[] png)
    {
        try { return ToBmp(png); }
        catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or
            OverflowException or ArgumentException)
        {
            return png;
        }
    }

    public static byte[] ToBmp(byte[] png)
    {
        if (png.Length < 33 || !png.AsSpan(0, 8).SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }))
            throw new InvalidDataException("Invalid PNG signature.");
        var width = BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(16));
        var height = BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(20));
        if (width <= 0 || height <= 0 || width > 10000 || height > 10000 ||
            png[24] != 16 || png[25] != 2 || png[26] != 0 || png[27] != 0 || png[28] != 0)
            throw new NotSupportedException("Only non-interlaced 16-bit truecolor PNG is supported.");
        var scanlineBytes = checked(width * 6);
        var expectedLength = checked((scanlineBytes + 1) * height);
        if (expectedLength > 64 * 1024 * 1024)
            throw new NotSupportedException("PNG image exceeds the supported decoded size.");
        using var compressed = new MemoryStream();
        var offset = 8;
        while (offset + 12 <= png.Length)
        {
            var length = BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(offset));
            if (length < 0 || length > png.Length - offset - 12)
                throw new InvalidDataException("PNG chunk exceeds its source data.");
            if (png.AsSpan(offset + 4, 4).SequenceEqual(new byte[] { 73, 68, 65, 84 }))
                compressed.Write(png, offset + 8, length);
            offset += length + 12;
        }
        var encoded = compressed.ToArray();
        if (encoded.Length < 6 || (encoded[0] & 15) != 8)
            throw new InvalidDataException("PNG zlib stream is invalid.");
        var decoded = new byte[expectedLength];
        using (var stream = new DeflateStream(new MemoryStream(encoded, 2, encoded.Length - 6), CompressionMode.Decompress))
        {
            var read = 0;
            while (read < decoded.Length)
            {
                var count = stream.Read(decoded, read, decoded.Length - read);
                if (count == 0) throw new InvalidDataException("PNG image data is incomplete.");
                read += count;
            }
        }
        var stride = checked((width * 3 + 3) / 4 * 4);
        var bitmap = new byte[checked(54 + stride * height)];
        bitmap[0] = (byte)'B';
        bitmap[1] = (byte)'M';
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(2), bitmap.Length);
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(10), 54);
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(14), 40);
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(18), width);
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(22), -height);
        BinaryPrimitives.WriteUInt16LittleEndian(bitmap.AsSpan(26), 1);
        BinaryPrimitives.WriteUInt16LittleEndian(bitmap.AsSpan(28), 24);
        BinaryPrimitives.WriteInt32LittleEndian(bitmap.AsSpan(34), stride * height);
        var previous = new byte[scanlineBytes];
        var current = new byte[scanlineBytes];
        for (var row = 0; row < height; row++)
        {
            var filter = decoded[row * (scanlineBytes + 1)];
            for (var index = 0; index < scanlineBytes; index++)
            {
                var left = index >= 6 ? current[index - 6] : 0;
                var up = previous[index];
                var upperLeft = index >= 6 ? previous[index - 6] : 0;
                var prediction = filter switch
                {
                    0 => 0,
                    1 => left,
                    2 => up,
                    3 => (left + up) / 2,
                    4 => Paeth(left, up, upperLeft),
                    _ => throw new InvalidDataException("Unknown PNG scanline filter.")
                };
                current[index] = unchecked((byte)(decoded[row * (scanlineBytes + 1) + index + 1] + prediction));
            }
            for (var column = 0; column < width; column++)
            {
                var source = column * 6;
                var target = 54 + row * stride + column * 3;
                bitmap[target] = current[source + 4];
                bitmap[target + 1] = current[source + 2];
                bitmap[target + 2] = current[source];
            }
            (previous, current) = (current, previous);
        }
        return bitmap;
    }

    private static int Paeth(int left, int up, int upperLeft)
    {
        var prediction = left + up - upperLeft;
        var leftDistance = Math.Abs(prediction - left);
        var upDistance = Math.Abs(prediction - up);
        var upperLeftDistance = Math.Abs(prediction - upperLeft);
        return leftDistance <= upDistance && leftDistance <= upperLeftDistance ? left :
            upDistance <= upperLeftDistance ? up : upperLeft;
    }
}
