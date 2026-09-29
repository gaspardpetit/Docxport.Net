using System.Buffers.Binary;
using OpenMcdf;

namespace DocxportNet.Doc;

/// <summary>Queues focused changes to an indexed binary DOC and applies them when saved.</summary>
/// <remarks>The source bytes and index remain unchanged. Save produces a new compound file.</remarks>
public sealed class DocEditor : IDisposable
{
    private const int EnvelopeFibIndex = 97;
    private const int DopFibIndex = 31;
    private const int DopVisibilityOffset = 504;
    private const int DefaultDopLength = 610;

    private readonly byte[] _source;
    private readonly MemoryStream _input;
    private readonly DocStructure _index;
    private byte[]? _envelopeEdit;
    private bool? _visibilityEdit;
    private bool _disposed;

    private DocEditor(byte[] bytes)
    {
        _source = bytes;
        _input = new MemoryStream(_source, false);
        try { _index = new DocStructureWalker().Read(_input); }
        catch { _input.Dispose(); throw; }
    }

    /// <summary>Copies the input into an immutable source snapshot and indexes its FIB locations.</summary>
    public static DocEditor Open(byte[] bytes)
    {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        return new DocEditor((byte[])bytes.Clone());
    }

    public static DocEditor Open(Stream input)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (!input.CanRead) throw new ArgumentException("The input must be readable.", nameof(input));
        using var copy = new MemoryStream();
        var originalPosition = input.CanSeek ? input.Position : -1;
        try
        {
            if (input.CanSeek) input.Position = 0;
            input.CopyTo(copy);
        }
        finally
        {
            if (originalPosition >= 0) input.Position = originalPosition;
        }
        return new DocEditor(copy.ToArray());
    }

    /// <summary>The read-only structural index of the original DOC.</summary>
    public DocStructure Index { get { ThrowIfDisposed(); return _index; } }

    /// <summary>Replaces the complete envelope, including recipients and attachments.</summary>
    public DocEditor SetEmailEnvelope(DocEmailEnvelope envelope)
    {
        ThrowIfDisposed();
        RequireEnvelopeField();
        _envelopeEdit = DocEmailEnvelopeCodec.Serialize(envelope);
        _visibilityEdit = envelope.Visible;
        return this;
    }

    /// <summary>Removes the envelope reference and hides its header.</summary>
    /// <remarks>The editor clears a validated old envelope block, but this is not secure sanitization.</remarks>
    public DocEditor RemoveEmailEnvelope()
    {
        ThrowIfDisposed();
        RequireEnvelopeField();
        _envelopeEdit = Array.Empty<byte>();
        _visibilityEdit = false;
        return this;
    }

    /// <summary>Changes only the header visibility, preserving any existing envelope bytes.</summary>
    public DocEditor SetEmailEnvelopeVisibility(bool visible)
    {
        ThrowIfDisposed();
        RequireEnvelopeField();
        _visibilityEdit = visible;
        return this;
    }

    /// <summary>Reads supported version 8 envelope settings, including queued changes.</summary>
    /// <remarks>Unknown envelope fields cannot be represented by this settings model; visibility-only edits still preserve them.</remarks>
    public DocEmailEnvelope? ReadEmailEnvelope()
    {
        ThrowIfDisposed();
        var bytes = _envelopeEdit;
        if (bytes == null)
        {
            var location = RequireEnvelopeField();
            if (!location.IsPresent) return null;
            if (location.Length > int.MaxValue)
                throw new NotSupportedException("The envelope is too large to materialize.");
            bytes = _index.ReadRange(location.StreamName, location.Offset, (int)location.Length);
        }
        if (bytes.Length == 0) return null;
        return DocEmailEnvelopeCodec.Deserialize(bytes, ReadVisibility());
    }

    /// <summary>Applies queued edits to a new DOC byte array. Repeated saves leave the source unchanged.</summary>
    public byte[] Save(CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        cancellationToken.ThrowIfCancellationRequested();
        if (_envelopeEdit == null && _visibilityEdit == null) return (byte[])_source.Clone();
        var envelope = RequireEnvelopeField();
        var dop = RequireDopField();
        if (_envelopeEdit != null && envelope.IsPresent)
            ValidateOldEnvelope(envelope);
        if (_visibilityEdit.HasValue && dop.IsPresent && dop.Length < DopVisibilityOffset + 1)
            throw new NotSupportedException("The DOP has no Dop2000 visibility field.");
        if (_visibilityEdit.HasValue && dop.IsPresent)
            ValidateDopVisibility(dop);
        using var output = new MemoryStream();
        output.Write(_source, 0, _source.Length);
        output.Position = 0;
        using (var storage = RootStorage.Open(output, StorageModeFlags.LeaveOpen))
        {
            using var word = storage.OpenStream("WordDocument");
            using var table = storage.OpenStream(_index.FibBase.TableStreamName);
            if (_envelopeEdit != null)
            {
                if (envelope.IsPresent)
                {
                    ClearOldEnvelope(table, envelope, cancellationToken);
                }
                var newOffset = _envelopeEdit.Length == 0 ? 0u : checked((uint)table.Length);
                if (_envelopeEdit.Length != 0)
                {
                    table.Position = table.Length;
                    table.Write(_envelopeEdit, 0, _envelopeEdit.Length);
                }
                WriteFibPair(word, envelope.FibFieldOffset, newOffset, checked((uint)_envelopeEdit.Length));
            }
            if (_visibilityEdit.HasValue)
            {
                uint dopOffset;
                if (dop.IsPresent)
                {
                    dopOffset = dop.Offset;
                }
                else
                {
                    dopOffset = checked((uint)table.Length);
                    table.Position = table.Length;
                    table.Write(new byte[DefaultDopLength], 0, DefaultDopLength);
                    WriteFibPair(word, dop.FibFieldOffset, dopOffset, DefaultDopLength);
                }
                table.Position = checked((long)dopOffset + DopVisibilityOffset);
                var flags = table.ReadByte();
                if (flags < 0) throw new InvalidDataException("The DOP visibility byte is missing.");
                table.Position--;
                table.WriteByte((byte)(_visibilityEdit.Value ? flags | 2 : flags & ~2));
            }
            cancellationToken.ThrowIfCancellationRequested();
            storage.Flush();
        }
        return output.ToArray();
    }

    public void Save(Stream output, CancellationToken cancellationToken = default)
    {
        if (output == null) throw new ArgumentNullException(nameof(output));
        if (!output.CanWrite) throw new ArgumentException("The output must be writable.", nameof(output));
        var bytes = Save(cancellationToken);
        output.Write(bytes, 0, bytes.Length);
    }

    private bool ReadVisibility()
    {
        if (_visibilityEdit.HasValue) return _visibilityEdit.Value;
        var dop = RequireDopField();
        return dop.IsPresent && dop.Length >= DopVisibilityOffset + 1 &&
            (_index.ReadRange(dop.StreamName, checked((long)dop.Offset + DopVisibilityOffset), 1)[0] & 2) != 0;
    }

    private DocLocation RequireEnvelopeField() => RequireField(EnvelopeFibIndex, "email envelope");
    private DocLocation RequireDopField() => RequireField(DopFibIndex, "document properties");

    private DocLocation RequireField(int index, string name)
    {
        if (_index.Locations.Count <= index)
            throw new NotSupportedException($"The DOC FIB has no {name} field.");
        return _index.Locations[index];
    }

    private void ValidateOldEnvelope(DocLocation envelope)
    {
        if (envelope.Length < 20)
            throw new InvalidDataException("The existing envelope header is truncated.");
        var header = _index.ReadRange(envelope.StreamName, envelope.Offset, 20);
        if (new Guid(header.Take(16).ToArray()) != DocEmailEnvelopeCodec.ClassId)
            throw new InvalidDataException("The existing envelope class identifier is invalid.");
        var version = BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(16));
        if (version != 6 && version != 8)
            throw new NotSupportedException("The existing envelope version cannot be replaced or removed.");
        foreach (var location in _index.Locations)
        {
            if (location.FibIndex == EnvelopeFibIndex || !location.IsRange || !location.IsPresent ||
                location.StreamName != envelope.StreamName) continue;
            if ((ulong)envelope.Offset < (ulong)location.Offset + location.Length &&
                (ulong)location.Offset < (ulong)envelope.Offset + envelope.Length)
                throw new InvalidDataException($"The envelope overlaps {location.Name}.");
        }
    }

    private void ValidateDopVisibility(DocLocation dop)
    {
        var visibilityOffset = (ulong)dop.Offset + DopVisibilityOffset;
        foreach (var location in _index.Locations)
        {
            if (location.FibIndex == DopFibIndex || !location.IsRange || !location.IsPresent ||
                location.StreamName != dop.StreamName) continue;
            if ((ulong)location.Offset <= visibilityOffset &&
                visibilityOffset < (ulong)location.Offset + location.Length)
                throw new InvalidDataException($"The DOP visibility byte overlaps {location.Name}.");
        }
    }

    private static void ClearOldEnvelope(Stream table, DocLocation envelope, CancellationToken cancellationToken)
    {
        table.Position = envelope.Offset;
        var zeros = new byte[8192];
        long remaining = envelope.Length;
        while (remaining > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var count = (int)Math.Min(remaining, zeros.Length);
            table.Write(zeros, 0, count);
            remaining -= count;
        }
        if ((long)envelope.Offset + envelope.Length == table.Length)
            table.SetLength(envelope.Offset);
    }

    private static void WriteFibPair(Stream word, int offset, uint valueOffset, uint length)
    {
        var pair = new byte[8];
        BinaryPrimitives.WriteUInt32LittleEndian(pair.AsSpan(), valueOffset);
        BinaryPrimitives.WriteUInt32LittleEndian(pair.AsSpan(4), length);
        word.Position = offset;
        word.Write(pair, 0, pair.Length);
    }

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DocEditor));
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _index.Dispose();
        _input.Dispose();
    }
}
