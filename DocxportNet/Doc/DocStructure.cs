using System.Buffers.Binary;
using OpenMcdf;

namespace DocxportNet.Doc;

/// <summary>A location in one compound-file stream. An empty location has no meaningful offset.</summary>
public sealed record DocLocation(string Name, string StreamName, uint Offset, uint Length, int FibIndex, int FibFieldOffset, bool IsRange)
{
    public bool IsPresent => IsRange && Length != 0;
}

/// <summary>The subset of FibBase needed to navigate a binary Word document.</summary>
public sealed record DocFibBase(ushort Identifier, ushort Version, ushort Flags, string TableStreamName)
{
    public uint FcMin { get; init; }
    public uint FcMac { get; init; }
    public bool IsEncrypted => (Flags & 0x0100) != 0;
    public bool IsObfuscated => (Flags & 0x8000) != 0;
}

/// <summary>A logical document-part range in global character positions.</summary>
public sealed record DocPartRange(string Name, uint CpStart, uint CpEnd)
{
    public uint Length => CpEnd - CpStart;
}

/// <summary>A structural node. Offsets are relative to the named compound-file stream.</summary>
public sealed class DocStructureNode
{
    private Lazy<DocParsedBlock>? _payload;

    public DocStructureNode(string kind, string name, string? streamName = null, long? offset = null, long? length = null)
    {
        Kind = kind;
        Name = name;
        StreamName = streamName;
        Offset = offset;
        Length = length;
    }

    public string Kind { get; }
    public string Name { get; }
    public string? StreamName { get; }
    public long? Offset { get; }
    public long? Length { get; }
    public IDictionary<string, string> Attributes { get; } = new SortedDictionary<string, string>(StringComparer.Ordinal);
    public IList<DocStructureNode> Children { get; } = new List<DocStructureNode>();

    /// <summary>Whether this node has a structured payload available on request.</summary>
    public bool HasPayload => _payload != null;

    public bool IsPayloadLoaded => _payload?.IsValueCreated == true;

    /// <summary>Parses this block on first access, then returns the cached result.</summary>
    public DocParsedBlock? Payload => _payload?.Value;

    internal void SetPayloadFactory(Func<DocParsedBlock> parser) =>
        _payload = new Lazy<DocParsedBlock>(parser);
}

public abstract record DocParsedBlock;
public sealed record DocDopVisibility(bool Visible) : DocParsedBlock;
public sealed record DocEnvelopeHeader(Guid Clsid, uint Version) : DocParsedBlock;
public sealed record DocTextPieceContent(uint CpStart, uint CpEnd, string Text) : DocParsedBlock;

/// <summary>The discovered container and FIB directory, without eagerly parsing document content.</summary>
public sealed class DocStructure : IDisposable
{
    private readonly RootStorage _storage;
    private Stream? _ownedInput;
    private bool _disposed;

    internal DocStructure(DocStructureNode root, DocFibBase fibBase, IReadOnlyList<DocLocation> locations,
        IReadOnlyList<DocPartRange> parts,
        RootStorage storage, Stream? ownedInput = null)
    {
        Root = root;
        FibBase = fibBase;
        Locations = locations;
        Parts = parts;
        _storage = storage;
        _ownedInput = ownedInput;
    }

    public DocStructureNode Root { get; }
    public DocFibBase FibBase { get; }
    public IReadOnlyList<DocLocation> Locations { get; }
    public IReadOnlyList<DocPartRange> Parts { get; }
    public DocLocation? FindLocation(string name) => Locations.FirstOrDefault(x => x.Name == name);

    internal void OwnInput(Stream input) => _ownedInput = input;

    internal Stream OpenStream(string name)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DocStructure));
        return _storage.OpenStream(name);
    }

    internal byte[] ReadRange(string streamName, long offset, int length)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DocStructure));
        using var stream = _storage.OpenStream(streamName);
        if (offset < 0 || length < 0 || offset > stream.Length || length > stream.Length - offset)
            throw new InvalidDataException("The block range is outside its stream.");
        stream.Position = offset;
        var bytes = new byte[length];
        var read = 0;
        while (read < length)
        {
            var count = stream.Read(bytes, read, length - read);
            if (count == 0) throw new InvalidDataException("The block ended unexpectedly.");
            read += count;
        }
        return bytes;
    }

    internal DocParsedBlock ReadDopVisibility(string streamName, long offset)
        => new DocDopVisibility((ReadRange(streamName, offset, 1)[0] & 2) != 0);

    internal DocParsedBlock ReadEnvelopeHeader(string streamName, long offset)
    {
        var bytes = ReadRange(streamName, offset, 20);
        return new DocEnvelopeHeader(new Guid(bytes.Take(16).ToArray()),
            BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(16)));
    }

    internal DocParsedBlock ReadTextPiece(uint cpStart, uint cpEnd, long offset, int byteLength, bool compressed)
    {
        var bytes = ReadRange("WordDocument", offset, byteLength);
        string text;
        if (compressed)
        {
            var characters = new char[bytes.Length];
            for (var i = 0; i < bytes.Length; i++)
                characters[i] = DecodeCompressed(bytes[i]);
            text = new string(characters);
        }
        else
        {
            try
            {
                text = new System.Text.UnicodeEncoding(false, false, true).GetString(bytes);
            }
            catch (System.Text.DecoderFallbackException ex)
            {
                throw new InvalidDataException("A UTF-16 text piece contains an invalid character sequence.", ex);
            }
        }
        if (text.Length != cpEnd - cpStart)
            throw new InvalidDataException("The decoded text length does not match its CP range.");
        return new DocTextPieceContent(cpStart, cpEnd, text);
    }

    private static char DecodeCompressed(byte value) => value switch
    {
        0x82 => '\u201A', 0x83 => '\u0192', 0x84 => '\u201E', 0x85 => '\u2026',
        0x86 => '\u2020', 0x87 => '\u2021', 0x88 => '\u02C6', 0x89 => '\u2030',
        0x8A => '\u0160', 0x8B => '\u2039', 0x8C => '\u0152', 0x91 => '\u2018',
        0x92 => '\u2019', 0x93 => '\u201C', 0x94 => '\u201D', 0x95 => '\u2022',
        0x96 => '\u2013', 0x97 => '\u2014', 0x98 => '\u02DC', 0x99 => '\u2122',
        0x9A => '\u0161', 0x9B => '\u203A', 0x9C => '\u0153', 0x9F => '\u0178',
        _ => (char)value
    };

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _storage.Dispose();
        _ownedInput?.Dispose();
    }
}
