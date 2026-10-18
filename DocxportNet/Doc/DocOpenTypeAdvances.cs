using System.Buffers.Binary;
using System.Collections.Concurrent;
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
    private readonly int _glyf;
    private readonly int _glyfLength;
    private readonly int _loca;
    private readonly int _locaFormat;
    private readonly int _math;
    private readonly int _mathLength;
    private readonly Dictionary<uint, short>? _kernPairs;
    private readonly int[][] _gposPairLookups = [];
    private readonly ConcurrentDictionary<uint, int> _pairAdjustments = new();

    public DocOpenTypeAdvances(byte[] data, int faceIndex = 0)
    {
        _data = data ?? throw new ArgumentNullException(nameof(data));
        if (faceIndex < 0) throw new ArgumentOutOfRangeException(nameof(faceIndex));
        var collection = Tag(0) == "ttcf";
        if (collection && faceIndex >= U32(8) || !collection && faceIndex != 0)
            throw new InvalidDataException("The requested OpenType collection face is unavailable.");
        var face = collection ? checked((int)U32(checked(12 + faceIndex * 4))) : 0;
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
        if (tables.TryGetValue("glyf", out var glyf) &&
            tables.TryGetValue("loca", out var loca) && tables["head"].Length >= 54)
        {
            var locaFormat = U16(head + 50);
            if (locaFormat is 0 or 1 &&
                loca.Length >= checked((_glyphCount + 1) * (locaFormat == 0 ? 2 : 4)))
            {
                _glyf = glyf.Offset;
                _glyfLength = glyf.Length;
                _loca = loca.Offset;
                _locaFormat = locaFormat;
            }
        }
        if (tables.TryGetValue("MATH", out var math) && math.Length >= 10)
        {
            _math = math.Offset;
            _mathLength = math.Length;
        }
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
        if (tables.TryGetValue("GPOS", out var positioning) && positioning.Length >= 10)
            _gposPairLookups = ReadPairLookups(positioning.Offset);
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
            if (useKerning && previousGlyph != 0)
                advances += PairAdjustment(previousGlyph, glyph);
            previousGlyph = glyph;
        }
        widthPoints = advances * fontSizePoints / _unitsPerEm;
        return true;
    }

    public bool TryGetGlyphVerticalBounds(int codepoint, out short minimum, out short maximum)
    {
        minimum = maximum = 0;
        if (_loca == 0) return false;
        var glyph = Glyph(codepoint);
        if (glyph == 0 || glyph >= _glyphCount) return false;
        var start = _locaFormat == 0 ? U16(_loca + glyph * 2) * 2 :
            checked((int)U32(_loca + glyph * 4));
        var end = _locaFormat == 0 ? U16(_loca + (glyph + 1) * 2) * 2 :
            checked((int)U32(_loca + (glyph + 1) * 4));
        if (start == end || start < 0 || end > _glyfLength || start > end - 10)
            return false;
        minimum = unchecked((short)U16(_glyf + start + 4));
        maximum = unchecked((short)U16(_glyf + start + 8));
        return true;
    }

    public bool TryGetMathVerticalVariant(int codepoint, out ushort variant)
    {
        variant = 0;
        if (_math == 0) return false;
        var variantsOffset = U16(_math + 8);
        if (variantsOffset == 0 || variantsOffset > _mathLength - 10) return false;
        var variants = _math + variantsOffset;
        var coverageOffset = U16(variants + 2);
        var count = U16(variants + 6);
        if (coverageOffset == 0 || coverageOffset > _mathLength - variantsOffset - 4 ||
            count > (_mathLength - variantsOffset - 10) / 2)
            return false;
        var coverage = CoverageIndex(variants + coverageOffset, Glyph(codepoint));
        if (coverage < 0 || coverage >= count) return false;
        var constructionOffset = U16(variants + 10 + coverage * 2);
        if (constructionOffset == 0 || constructionOffset > _mathLength - variantsOffset - 8)
            return false;
        var construction = variants + constructionOffset;
        var variantCount = U16(construction + 2);
        if (variantCount > (_mathLength - variantsOffset - constructionOffset - 4) / 4)
            return false;
        var baseGlyph = Glyph(codepoint);
        for (var index = 0; index < variantCount; index++)
        {
            var candidate = U16(construction + 4 + index * 4);
            if (candidate == 0 || candidate == baseGlyph || candidate >= _glyphCount) continue;
            variant = (ushort)candidate;
            return true;
        }
        return false;
    }

    public int UnitsPerEm => _unitsPerEm;

    public int GlyphIndex(int codepoint) => Glyph(codepoint);

    public double KerningAdjustment(char previous, char current, double fontSizePoints)
    {
        if (fontSizePoints <= 0) return 0;
        var previousGlyph = Glyph(previous);
        var currentGlyph = Glyph(current);
        return previousGlyph != 0 && currentGlyph != 0
            ? PairAdjustment(previousGlyph, currentGlyph) * fontSizePoints / _unitsPerEm : 0;
    }

    public double LegacyKerningAdjustment(char previous, char current, double fontSizePoints)
    {
        if (fontSizePoints <= 0) return 0;
        if (_kernPairs == null) return KerningAdjustment(previous, current, fontSizePoints);
        var previousGlyph = Glyph(previous);
        var currentGlyph = Glyph(current);
        var key = ((uint)previousGlyph << 16) | (uint)currentGlyph;
        return _kernPairs.TryGetValue(key, out var adjustment)
            ? adjustment * fontSizePoints / _unitsPerEm : 0;
    }


    private int[][] ReadPairLookups(int positioning)
    {
        var scripts = positioning + U16(positioning + 4);
        var features = positioning + U16(positioning + 6);
        var lookups = positioning + U16(positioning + 8);
        var scriptCount = U16(scripts);
        var selectedScript = -1;
        for (var index = 0; index < scriptCount; index++)
        {
            var record = scripts + 2 + index * 6;
            var tag = Tag(record);
            if (tag == "latn" || tag == "DFLT" && selectedScript < 0)
                selectedScript = scripts + U16(record + 4);
            if (tag == "latn") break;
        }
        if (selectedScript < 0) return [];
        var defaultLanguage = U16(selectedScript);
        if (defaultLanguage == 0) return [];
        var language = selectedScript + defaultLanguage;
        var featureCount = U16(features);
        var lookupIndices = new HashSet<int>();
        for (var index = 0; index < U16(language + 4); index++)
        {
            var featureIndex = U16(language + 6 + index * 2);
            if (featureIndex >= featureCount) continue;
            var record = features + 2 + featureIndex * 6;
            if (Tag(record) != "kern") continue;
            var feature = features + U16(record + 4);
            for (var item = 0; item < U16(feature + 2); item++)
                lookupIndices.Add(U16(feature + 4 + item * 2));
        }
        var result = new List<int[]>();
        foreach (var index in lookupIndices.OrderBy(value => value))
        {
            if (index >= U16(lookups)) continue;
            var lookup = lookups + U16(lookups + 2 + index * 2);
            if (U16(lookup) != 2) continue;
            var subtables = new int[U16(lookup + 4)];
            for (var item = 0; item < subtables.Length; item++)
                subtables[item] = lookup + U16(lookup + 6 + item * 2);
            result.Add(subtables);
        }
        return result.ToArray();
    }

    private int PairAdjustment(int first, int second)
    {
        var key = ((uint)first << 16) | (uint)second;
        if (_pairAdjustments.TryGetValue(key, out var cached)) return cached;
        var adjustment = 0;
        foreach (var lookup in _gposPairLookups)
        {
            foreach (var subtable in lookup)
            {
                var coverage = CoverageIndex(subtable + U16(subtable + 2), first);
                if (coverage < 0) continue;
                var firstFormat = U16(subtable + 4);
                var secondFormat = U16(subtable + 6);
                if (U16(subtable) == 1)
                {
                    var count = U16(subtable + 8);
                    if (coverage >= count) continue;
                    var pairSet = subtable + U16(subtable + 10 + coverage * 2);
                    var recordSize = 2 + ValueSize(firstFormat) + ValueSize(secondFormat);
                    var matched = false;
                    for (var index = 0; index < U16(pairSet); index++)
                    {
                        var record = pairSet + 2 + index * recordSize;
                        if (U16(record) != second) continue;
                        adjustment += XAdvance(record + 2, firstFormat);
                        matched = true;
                        break;
                    }
                    if (matched) break;
                }
                else if (U16(subtable) == 2)
                {
                    var firstClass = ClassIndex(subtable + U16(subtable + 8), first);
                    var secondClass = ClassIndex(subtable + U16(subtable + 10), second);
                    var firstCount = U16(subtable + 12);
                    var secondCount = U16(subtable + 14);
                    if (firstClass >= firstCount || secondClass >= secondCount) continue;
                    var record = subtable + 16 + (firstClass * secondCount + secondClass) *
                        (ValueSize(firstFormat) + ValueSize(secondFormat));
                    adjustment += XAdvance(record, firstFormat);
                    break;
                }
            }
        }
        if (_gposPairLookups.Length == 0 && _kernPairs != null &&
            _kernPairs.TryGetValue(key, out var legacy)) adjustment = legacy;
        _pairAdjustments.TryAdd(key, adjustment);
        return adjustment;
    }

    private int CoverageIndex(int coverage, int glyph)
    {
        var count = U16(coverage + 2);
        if (U16(coverage) == 1)
        {
            for (var index = 0; index < count; index++)
                if (U16(coverage + 4 + index * 2) == glyph) return index;
        }
        else if (U16(coverage) == 2)
        {
            for (var index = 0; index < count; index++)
            {
                var record = coverage + 4 + index * 6;
                if (glyph >= U16(record) && glyph <= U16(record + 2))
                    return U16(record + 4) + glyph - U16(record);
            }
        }
        return -1;
    }

    private int ClassIndex(int definition, int glyph)
    {
        if (U16(definition) == 1)
        {
            var count = U16(definition + 4);
            var index = glyph - U16(definition + 2);
            return index >= 0 && index < count ? U16(definition + 6 + index * 2) : 0;
        }
        if (U16(definition) == 2)
            for (var index = 0; index < U16(definition + 2); index++)
            {
                var record = definition + 4 + index * 6;
                if (glyph >= U16(record) && glyph <= U16(record + 2)) return U16(record + 4);
            }
        return 0;
    }

    private static int ValueSize(int format) => PopCount((uint)format) * 2;

    private static int PopCount(uint value)
    {
        int count = 0;
        while (value != 0)
        {
            count++;
            value &= value - 1;
        }
        return count;
    }

    private int XAdvance(int record, int format)
    {
        if ((format & 4) == 0) return 0;
        var offset = PopCount((uint)(format & 3)) * 2;
        return unchecked((short)U16(record + offset));
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
