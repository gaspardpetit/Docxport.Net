using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using OpenMcdf;

namespace DocxportNet.Doc;

/// <summary>Decrypts the three RC4 CryptoAPI streams into a temporary compound file.</summary>
internal static class DocCryptoApiReader
{
    public static MemoryStream Decrypt(Stream input, string password)
    {
        if (password == null) throw new ArgumentNullException(nameof(password));
        input.Position = 0;
        using var source = RootStorage.Open(input, StorageModeFlags.LeaveOpen);
        if (!source.ContainsEntry("WordDocument"))
            throw new InvalidDataException("The compound file has no WordDocument stream.");
        var word = ReadAll(source.OpenStream("WordDocument"));
        if (word.Length < 68 || BinaryPrimitives.ReadUInt16LittleEndian(word) != 0xA5EC)
            throw new InvalidDataException("The WordDocument FIB is invalid.");
        var flags = BinaryPrimitives.ReadUInt16LittleEndian(word.AsSpan(10));
        if ((flags & 0x0100) == 0 || (flags & 0x8000) != 0)
            throw new NotSupportedException("The DOC is not encrypted with RC4 CryptoAPI.");
        var tableName = (flags & 0x0200) != 0 ? "1Table" : "0Table";
        if (!source.ContainsEntry(tableName))
            throw new InvalidDataException("The selected table stream is missing.");
        var table = ReadAll(source.OpenStream(tableName));
        var headerLength = BinaryPrimitives.ReadUInt32LittleEndian(word.AsSpan(14));
        if (headerLength < 72 || headerLength > table.Length)
            throw new InvalidDataException("The CryptoAPI header length is invalid.");
        var major = BinaryPrimitives.ReadUInt16LittleEndian(table);
        var minor = BinaryPrimitives.ReadUInt16LittleEndian(table.AsSpan(2));
        var size = BinaryPrimitives.ReadUInt32LittleEndian(table.AsSpan(8));
        if (major is < 2 or > 4 || minor != 2 || size < 32 ||
            size > headerLength - 72)
            throw new NotSupportedException("The DOC does not have a supported RC4 CryptoAPI header.");
        var keyBits = BinaryPrimitives.ReadUInt32LittleEndian(table.AsSpan(28));
        if (keyBits == 0) keyBits = 40;
        if ((BinaryPrimitives.ReadUInt32LittleEndian(table.AsSpan(4)) & 4) == 0 ||
            (BinaryPrimitives.ReadUInt32LittleEndian(table.AsSpan(12)) & 4) == 0 ||
            BinaryPrimitives.ReadUInt32LittleEndian(table.AsSpan(20)) != 0x6801 ||
            BinaryPrimitives.ReadUInt32LittleEndian(table.AsSpan(24)) != 0x8004 ||
            BinaryPrimitives.ReadUInt32LittleEndian(table.AsSpan(32)) != 1 ||
            keyBits is < 40 or > 128 || keyBits % 8 != 0)
            throw new NotSupportedException("Only SHA-1 RC4 CryptoAPI DOC encryption is supported.");
        var verifier = checked((int)(12 + size));
        if (BinaryPrimitives.ReadUInt32LittleEndian(table.AsSpan(verifier)) != 16 ||
            BinaryPrimitives.ReadUInt32LittleEndian(table.AsSpan(verifier + 36)) != 20)
            throw new InvalidDataException("The CryptoAPI verifier is invalid.");
        var salt = table.AsSpan(verifier + 4, 16).ToArray();
        var seed = Hash(Combine(salt, Encoding.Unicode.GetBytes(password)));
        var encryptedCheck = table.AsSpan(verifier + 20, 16 + 4 + 20).ToArray();
        var encrypted = new byte[36];
        Array.Copy(encryptedCheck, 0, encrypted, 0, 16);
        Array.Copy(encryptedCheck, 20, encrypted, 16, 20);
        Rc4(encrypted, Key(seed, 0, keyBits));
        if (!Hash(encrypted.AsSpan(0, 16).ToArray()).SequenceEqual(encrypted.Skip(16)))
            throw new UnauthorizedAccessException("The DOC password is incorrect.");

        var originalWordPrefix = word.AsSpan(0, 68).ToArray();
        var originalTablePrefix = table.AsSpan(0, checked((int)headerLength)).ToArray();
        TransformBlocks(word, seed, keyBits);
        TransformBlocks(table, seed, keyBits);
        originalWordPrefix.CopyTo(word, 0);
        originalTablePrefix.CopyTo(table, 0);
        BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(10), (ushort)(flags & ~0x0100));
        byte[]? data = null;
        if (source.ContainsEntry("Data"))
        {
            data = ReadAll(source.OpenStream("Data"));
            TransformBlocks(data, seed, keyBits);
        }
        var output = new MemoryStream();
        using (var target = RootStorage.Create(output, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen))
            CopyEntries(source, target, tableName, word, table, data, true);
        output.Position = 0;
        return output;
    }

    private static void CopyEntries(Storage source, Storage target, string tableName,
        byte[] word, byte[] table, byte[]? data, bool root)
    {
        foreach (var entry in source.EnumerateEntries())
        {
            if (entry.Type == EntryType.Storage)
                CopyEntries(source.OpenStorage(entry.Name), target.CreateStorage(entry.Name),
                    tableName, word, table, data, false);
            else
            {
                var bytes = root && entry.Name == "WordDocument" ? word :
                    root && entry.Name == tableName ? table :
                    root && entry.Name == "Data" && data != null ? data : ReadAll(source.OpenStream(entry.Name));
                using var stream = target.CreateStream(entry.Name);
                stream.Write(bytes, 0, bytes.Length);
            }
        }
    }

    private static byte[] ReadAll(Stream stream)
    {
        using (stream)
        using (var copy = new MemoryStream())
        {
            stream.CopyTo(copy);
            return copy.ToArray();
        }
    }

    private static byte[] Hash(byte[] bytes)
    {
        using var sha = SHA1.Create();
        return sha.ComputeHash(bytes);
    }

    private static byte[] Combine(byte[] first, byte[] second)
    {
        var output = new byte[first.Length + second.Length];
        Buffer.BlockCopy(first, 0, output, 0, first.Length);
        Buffer.BlockCopy(second, 0, output, first.Length, second.Length);
        return output;
    }

    private static byte[] Key(byte[] seed, uint block, uint bits)
    {
        var input = new byte[seed.Length + 4];
        Buffer.BlockCopy(seed, 0, input, 0, seed.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(input.AsSpan(seed.Length), block);
        var hash = Hash(input);
        var length = bits == 40 ? 16 : checked((int)bits / 8);
        var key = new byte[length];
        Buffer.BlockCopy(hash, 0, key, 0, bits == 40 ? 5 : length);
        return key;
    }

    private static void TransformBlocks(byte[] bytes, byte[] seed, uint bits)
    {
        for (var cursor = 0; cursor < bytes.Length; cursor += 512)
        {
            var block = new byte[Math.Min(512, bytes.Length - cursor)];
            Buffer.BlockCopy(bytes, cursor, block, 0, block.Length);
            Rc4(block, Key(seed, checked((uint)(cursor / 512)), bits));
            Buffer.BlockCopy(block, 0, bytes, cursor, block.Length);
        }
    }

    private static void Rc4(byte[] bytes, byte[] key)
    {
        var state = new byte[256];
        for (var i = 0; i < 256; i++) state[i] = (byte)i;
        var j = 0;
        for (var i = 0; i < 256; i++)
        {
            j = (j + state[i] + key[i % key.Length]) & 255;
            (state[i], state[j]) = (state[j], state[i]);
        }
        var a = 0;
        j = 0;
        for (var k = 0; k < bytes.Length; k++)
        {
            a = (a + 1) & 255;
            j = (j + state[a]) & 255;
            (state[a], state[j]) = (state[j], state[a]);
            bytes[k] ^= state[(state[a] + state[j]) & 255];
        }
    }
}
