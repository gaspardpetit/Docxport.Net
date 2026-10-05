using System.Security.Cryptography;

namespace DocxportNet.Doc;

internal static class DocBinaryCompat
{
    public static byte[] Md5(byte[] bytes)
    {
        using var md5 = MD5.Create();
        return md5.ComputeHash(bytes);
    }

    public static byte[] Hex(string value)
    {
        if ((value.Length & 1) != 0)
            throw new FormatException("A hexadecimal byte string must have an even length.");
        var bytes = new byte[value.Length / 2];
        for (var i = 0; i < bytes.Length; i++)
            bytes[i] = Convert.ToByte(value.Substring(i * 2, 2), 16);
        return bytes;
    }

    public static bool IsFinite(double value) =>
        !double.IsNaN(value) && !double.IsInfinity(value);

    public static bool IsTiff(ReadOnlySpan<byte> bytes) =>
        bytes.Length >= 4 &&
        (bytes[0] == 0x49 && bytes[1] == 0x49 && bytes[2] == 0x2A && bytes[3] == 0 ||
         bytes[0] == 0x4D && bytes[1] == 0x4D && bytes[2] == 0 && bytes[3] == 0x2A);

    public static void Write(this Stream stream, byte[] bytes) =>
        stream.Write(bytes, 0, bytes.Length);

    public static void Write(this Stream stream, ReadOnlySpan<byte> bytes)
    {
        var buffer = bytes.ToArray();
        stream.Write(buffer, 0, buffer.Length);
    }

    public static void ReadExactly(this Stream stream, byte[] bytes)
    {
        var offset = 0;
        while (offset < bytes.Length)
        {
            var count = stream.Read(bytes, offset, bytes.Length - offset);
            if (count == 0) throw new EndOfStreamException();
            offset += count;
        }
    }

    public static string ToHexString(byte[] bytes) =>
        BitConverter.ToString(bytes).Replace("-", string.Empty);

    public static bool AllBytesEqual(ReadOnlySpan<byte> bytes, byte value)
    {
        foreach (var item in bytes)
            if (item != value) return false;
        return true;
    }
}
