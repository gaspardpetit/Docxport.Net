using System.Buffers.Binary;

namespace DocxportNet.Visitors.Pdf;

internal static class DxpTiffConverter
{
    public static byte[] ToBmpOrOriginal(byte[] tiff)
    {
        try { return ToBmp(tiff); }
        catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or
            OverflowException or ArgumentException)
        {
            return tiff;
        }
    }

    public static byte[] ToBmp(byte[] tiff)
    {
        if (tiff.Length < 8 || tiff[0] != 'I' || tiff[1] != 'I' ||
            BinaryPrimitives.ReadUInt16LittleEndian(tiff.AsSpan(2)) != 42)
            throw new NotSupportedException("Only little-endian classic TIFF is supported.");
        var directory = checked((int)ReadUInt32(4));
        var entryCount = ReadUInt16(directory);
        var entries = new Dictionary<int, (int Type, int Count, int Offset)>();
        for (var index = 0; index < entryCount; index++)
        {
            var at = checked(directory + 2 + index * 12);
            entries[ReadUInt16(at)] = (ReadUInt16(at + 2), checked((int)ReadUInt32(at + 4)), at + 8);
        }
        var width = Value(256);
        var height = Value(257);
        if (width <= 0 || height <= 0 || width > 10000 || height > 10000 ||
            Value(259) != 5 || Value(262) != 2 || Value(277) != 3 ||
            Value(284, 1) != 1 || Value(274, 1) != 1 ||
            Value(317, 1) is not (1 or 2) || Value(278) != height ||
            !entries.TryGetValue(258, out var bits) || bits.Type != 3 || bits.Count != 3)
            throw new NotSupportedException("This TIFF image layout is unsupported.");
        var bitsOffset = checked((int)ReadUInt32(bits.Offset));
        var sampleBits = ReadUInt16(bitsOffset);
        if (sampleBits is not (8 or 16) || ReadUInt16(bitsOffset + 2) != sampleBits ||
            ReadUInt16(bitsOffset + 4) != sampleBits)
            throw new NotSupportedException("Only uniform 8/16-bit RGB TIFF samples are supported.");
        var stripOffset = Value(273);
        var stripLength = Value(279);
        if (stripOffset < 0 || stripLength < 0 || stripOffset > tiff.Length - stripLength)
            throw new InvalidDataException("TIFF strip exceeds the source data.");
        var bytesPerSample = sampleBits / 8;
        var pixelBytes = 3 * bytesPerSample;
        var expectedLength = checked(width * height * pixelBytes);
        if (expectedLength > 64 * 1024 * 1024)
            throw new NotSupportedException("TIFF image exceeds the supported decoded size.");
        var decoded = DecodeLzw(tiff.AsSpan(stripOffset, stripLength), expectedLength);
        if (Value(317, 1) == 2)
        {
            for (var row = 0; row < height; row++)
                for (var column = pixelBytes; column < width * pixelBytes; column += bytesPerSample)
                {
                    var at = row * width * pixelBytes + column;
                    if (sampleBits == 8)
                        decoded[at] = unchecked((byte)(decoded[at] + decoded[at - pixelBytes]));
                    else
                    {
                        var value = unchecked((ushort)(BinaryPrimitives.ReadUInt16LittleEndian(decoded.AsSpan(at)) +
                            BinaryPrimitives.ReadUInt16LittleEndian(decoded.AsSpan(at - pixelBytes))));
                        BinaryPrimitives.WriteUInt16LittleEndian(decoded.AsSpan(at), value);
                    }
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
        for (var row = 0; row < height; row++)
            for (var column = 0; column < width; column++)
            {
                var source = (row * width + column) * pixelBytes;
                var target = 54 + row * stride + column * 3;
                bitmap[target] = Sample(source);
                bitmap[target + 1] = Sample(source + bytesPerSample);
                bitmap[target + 2] = Sample(source + 2 * bytesPerSample);
            }
        return bitmap;

        byte Sample(int offset) => bytesPerSample == 1 ? decoded[offset] :
            (byte)((BinaryPrimitives.ReadUInt16LittleEndian(decoded.AsSpan(offset)) * 255L + 32767) / 65535);

        ushort ReadUInt16(int offset)
        {
            if (offset < 0 || offset > tiff.Length - 2) throw new InvalidDataException("Truncated TIFF entry.");
            return BinaryPrimitives.ReadUInt16LittleEndian(tiff.AsSpan(offset));
        }

        uint ReadUInt32(int offset)
        {
            if (offset < 0 || offset > tiff.Length - 4) throw new InvalidDataException("Truncated TIFF entry.");
            return BinaryPrimitives.ReadUInt32LittleEndian(tiff.AsSpan(offset));
        }

        int Value(int tag, int fallback = -1)
        {
            if (!entries.TryGetValue(tag, out var entry)) return fallback;
            if (entry.Count != 1 || entry.Type is not (3 or 4))
                throw new NotSupportedException("TIFF tag arrays are unsupported here.");
            return entry.Type == 3 ? ReadUInt16(entry.Offset) : checked((int)ReadUInt32(entry.Offset));
        }
    }

    private static byte[] DecodeLzw(ReadOnlySpan<byte> encoded, int expectedLength)
    {
        var dictionary = new byte[4096][];
        for (var index = 0; index < 256; index++) dictionary[index] = [(byte)index];
        var output = new byte[expectedLength];
        var outputOffset = 0;
        var bitOffset = 0;
        var codeSize = 9;
        var nextCode = 258;
        byte[]? previous = null;
        while (bitOffset + codeSize <= encoded.Length * 8)
        {
            var code = 0;
            for (var bit = 0; bit < codeSize; bit++, bitOffset++)
                code = code << 1 | (encoded[bitOffset / 8] >> (7 - bitOffset % 8) & 1);
            if (code == 256)
            {
                codeSize = 9;
                nextCode = 258;
                previous = null;
                continue;
            }
            if (code == 257) break;
            byte[] entry;
            if (code < nextCode && dictionary[code] != null) entry = dictionary[code];
            else if (code == nextCode && previous != null)
            {
                entry = new byte[previous.Length + 1];
                previous.CopyTo(entry, 0);
                entry[previous.Length] = previous[0];
            }
            else throw new InvalidDataException("TIFF LZW code is invalid.");
            if (entry.Length > expectedLength - outputOffset)
                throw new InvalidDataException("TIFF LZW data exceeds its dimensions.");
            entry.CopyTo(output, outputOffset);
            outputOffset += entry.Length;
            if (previous != null && nextCode < dictionary.Length)
            {
                var added = new byte[previous.Length + 1];
                previous.CopyTo(added, 0);
                added[previous.Length] = entry[0];
                dictionary[nextCode++] = added;
                if (nextCode == (1 << codeSize) - 1 && codeSize < 12) codeSize++;
            }
            previous = entry;
        }
        if (outputOffset != expectedLength) throw new InvalidDataException("TIFF LZW data is incomplete.");
        return output;
    }
}
