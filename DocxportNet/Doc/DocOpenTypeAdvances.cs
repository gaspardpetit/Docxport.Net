using System.Buffers.Binary;
using System.Text;

namespace DocxportNet.Doc;

/// <summary>Reads unshaped horizontal glyph advances from a TrueType/OpenType face.</summary>
internal sealed class DocOpenTypeAdvances
{
    private readonly byte[] _data;
    private readonly int _hmtx;
    private readonly int _cmap;
    private readonly int _metricCount;
    private readonly int _glyphCount;
    private readonly int _unitsPerEm;
    private readonly int _cmapFormat;
    private readonly Dictionary<uint, short>? _kernPairs;

    public DocOpenTypeAdvances(byte[] data)
    {
        _data = data ?? throw new ArgumentNullException(nameof(data));
        var face = Tag(0) == "ttcf" ? checked((int)U32(12)) : 0;
        var tableCount = U16(face + 4);
        var tables = new Dictionary<string, (int Offset, int Length)>();
        for (var i = 0; i < tableCount; i++)
        {
            var record = checked(face + 12 + i * 16);
            var name = Tag(record);
            var offset = checked((int)U32(record + 8));
            var length = checked((int)U32(record + 12));
            Check(offset, length);
            tables[name] = (offset, length);
        }
        int Table(string name, int minimum)
        {
            if (!tables.TryGetValue(name, out var table) || table.Length < minimum)
                throw new InvalidDataException($"The OpenType {name} table is missing or truncated.");
            return table.Offset;
        }
        var head = Table("head", 20);
        var hhea = Table("hhea", 36);
        var maxp = Table("maxp", 6);
        _hmtx = Table("hmtx", 4);
        _cmap = Table("cmap", 4);
        _unitsPerEm = U16(head + 18);
        _metricCount = U16(hhea + 34);
        _glyphCount = U16(maxp + 4);
        if (_unitsPerEm == 0 || _metricCount == 0 || _metricCount > _glyphCount ||
            tables["hmtx"].Length < checked(_metricCount * 4 +
                (_glyphCount - _metricCount) * 2))
            throw new InvalidDataException("The OpenType horizontal metrics are invalid.");
        var encodingCount = U16(_cmap + 2);
        if (tables["cmap"].Length < checked(4 + encodingCount * 8))
            throw new InvalidDataException("The OpenType cmap encoding records are truncated.");
        var selected = -1;
        var priority = -1;
        for (var i = 0; i < encodingCount; i++)
        {
            var record = checked(_cmap + 4 + i * 8);
            var platform = U16(record);
            var encoding = U16(record + 2);
            var offset = checked(_cmap + (int)U32(record + 4));
            if (offset < _cmap || offset > _cmap + tables["cmap"].Length - 2)
                throw new InvalidDataException("An OpenType cmap subtable is outside the table.");
            var format = U16(offset);
            var score = format == 12 && platform == 3 && encoding == 10 ? 4 :
                format == 12 && platform == 0 ? 3 :
                format == 4 && platform == 3 && encoding == 1 ? 2 :
                format == 4 && platform == 0 ? 1 : -1;
            if (score <= priority) continue;
            selected = offset;
            priority = score;
        }
        if (selected < 0)
            throw new NotSupportedException("The OpenType face has no Unicode cmap format 4 or 12.");
        _cmapFormat = U16(selected);
        _cmap = selected;
        if (_cmapFormat == 12)
            Check(_cmap, checked(16 + (int)U32(_cmap + 12) * 12));
        else
        {
            var segments = U16(_cmap + 6) / 2;
            Check(_cmap, checked(16 + segments * 8));
        }
        if (tables.TryGetValue("kern", out var kern) && kern.Length >= 18 &&
            U16(kern.Offset) == 0 && U16(kern.Offset + 2) > 0)
        {
            var subtable = kern.Offset + 4;
            var coverage = U16(subtable + 4);
            // The old horizontal format-0 pair table is sufficient for
            // common Latin table labels. Ignore other formats and flags.
            if (U16(subtable) == 0 && (coverage >> 8) == 0 &&
                (coverage & 0x0001) != 0 && (coverage & 0x0006) == 0)
            {
                var count = U16(subtable + 6);
                var end = checked(subtable + 14 + count * 6);
                if (end <= kern.Offset + kern.Length)
                {
                    _kernPairs = new Dictionary<uint, short>(count);
                    for (var i = 0; i < count; i++)
                    {
                        var pair = subtable + 14 + i * 6;
                        var key = ((uint)U16(pair) << 16) | (uint)U16(pair + 2);
                        _kernPairs[key] = unchecked((short)U16(pair + 4));
                    }
                }
            }
        }
    }

    public bool TryMeasure(string text, double fontSizePoints, out double widthPoints,
        bool useKerning = false)
    {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (fontSizePoints <= 0) throw new ArgumentOutOfRangeException(nameof(fontSizePoints));
        long advances = 0;
        var previousGlyph = 0;
        for (var i = 0; i < text.Length; i++)
        {
            int scalar;
            if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length &&
                char.IsLowSurrogate(text[i + 1]))
                scalar = char.ConvertToUtf32(text[i], text[++i]);
            else
                scalar = text[i];
            if (scalar is '\r' or '\n') continue;
            var glyph = Glyph(scalar);
            if (glyph == 0 || glyph >= _glyphCount)
            {
                widthPoints = 0;
                return false; // A font substitution is needed; do not guess its width.
            }
            advances += U16(_hmtx + Math.Min(glyph, _metricCount - 1) * 4);
            if (useKerning && previousGlyph != 0 && _kernPairs != null &&
                _kernPairs.TryGetValue(((uint)previousGlyph << 16) | (uint)glyph,
                    out var adjustment))
                advances += adjustment;
            previousGlyph = glyph;
        }
        widthPoints = advances * fontSizePoints / _unitsPerEm;
        return true;
    }

    private int Glyph(int codepoint)
    {
        if (_cmapFormat == 12)
        {
            var groups = checked((int)U32(_cmap + 12));
            var low = 0;
            var high = groups - 1;
            while (low <= high)
            {
                var mid = low + (high - low) / 2;
                var row = checked(_cmap + 16 + mid * 12);
                var first = U32(row);
                var last = U32(row + 4);
                if (codepoint < first) high = mid - 1;
                else if (codepoint > last) low = mid + 1;
                else return checked((int)(U32(row + 8) + codepoint - first));
            }
            return 0;
        }
        if (codepoint > ushort.MaxValue) return 0;
        var count = U16(_cmap + 6) / 2;
        var ends = _cmap + 14;
        var starts = ends + count * 2 + 2;
        var deltas = starts + count * 2;
        var offsets = deltas + count * 2;
        for (var i = 0; i < count; i++)
        {
            if (codepoint > U16(ends + i * 2)) continue;
            if (codepoint < U16(starts + i * 2)) return 0;
            var delta = U16(deltas + i * 2);
            var offset = U16(offsets + i * 2);
            if (offset == 0) return (codepoint + delta) & 0xFFFF;
            var glyph = U16(offsets + i * 2 + offset +
                (codepoint - U16(starts + i * 2)) * 2);
            return glyph == 0 ? 0 : (glyph + delta) & 0xFFFF;
        }
        return 0;
    }

    private string Tag(int offset)
    {
        Check(offset, 4);
        return Encoding.ASCII.GetString(_data, offset, 4);
    }

    private int U16(int offset)
    {
        Check(offset, 2);
        return BinaryPrimitives.ReadUInt16BigEndian(_data.AsSpan(offset));
    }

    private uint U32(int offset)
    {
        Check(offset, 4);
        return BinaryPrimitives.ReadUInt32BigEndian(_data.AsSpan(offset));
    }

    private void Check(int offset, int length)
    {
        if (offset < 0 || length < 0 || offset > _data.Length - length)
            throw new InvalidDataException("The OpenType font data is truncated.");
    }
}
