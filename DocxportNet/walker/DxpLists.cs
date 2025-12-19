using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.api;
using System.Globalization;

namespace DocxportNet.walker;


public sealed class DxpNumberingResolver
{
	private readonly Numbering? _numbering;
	private readonly Dictionary<int, NumberingInstance> _numById = new();
	private readonly Dictionary<int, AbstractNum> _absById = new();
	private readonly Dictionary<string, Style> _styleById = new(StringComparer.Ordinal);

	public DxpNumberingResolver(WordprocessingDocument doc)
	{
		_numbering = doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
		if (_numbering != null)
		{
			foreach (var n in _numbering.Elements<NumberingInstance>())
				if (n.NumberID?.Value != null)
					_numById[n.NumberID.Value] = n;

			foreach (var a in _numbering.Elements<AbstractNum>())
				if (a.AbstractNumberId?.Value != null)
					_absById[a.AbstractNumberId.Value] = a;
		}

		var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
		if (styles != null)
		{
			foreach (var s in styles.Elements<Style>())
			{
				var id = s.StyleId?.Value;
				if (!string.IsNullOrEmpty(id))
					_styleById[id!] = s;
			}
		}
	}

	public (AbstractNum abs, Level lvl, int startAt)? ResolveLevel(int numId, int ilvl)
	{
		if (!_numById.TryGetValue(numId, out var num))
			return null;

		var absId = num.AbstractNumId?.Val?.Value;
		if (absId == null || !_absById.TryGetValue(absId.Value, out var abs))
			return null;

		// IMPORTANT: some abstractNum are just a link (numStyleLink/styleLink) and contain no <w:lvl>
		abs = FollowAbstractNumLinks(abs) ?? abs;

		// clamp + Word-like fallback to level 0 when missing
		if (ilvl < 0)
			ilvl = 0;
		if (ilvl > 8)
			ilvl = 8;

		var lvl = abs.Elements<Level>().FirstOrDefault(x => x.LevelIndex?.Value == ilvl)
			   ?? abs.Elements<Level>().FirstOrDefault(x => x.LevelIndex?.Value == 0)
			   ?? abs.Elements<Level>().OrderBy(x => x.LevelIndex?.Value ?? int.MaxValue).FirstOrDefault();

		if (lvl == null)
			return null;

		// startAt: from override if present, else from lvl.Start
		int startAt = 1;

		var ov = num.Elements<LevelOverride>()
					.FirstOrDefault(o => o.LevelIndex?.Value == ilvl);

		var startOverride = ov?.StartOverrideNumberingValue?.Val?.Value;
		if (startOverride != null)
			startAt = startOverride.Value;
		else if (lvl.StartNumberingValue?.Val?.Value != null)
			startAt = lvl.StartNumberingValue.Val.Value;

		return (abs, lvl, startAt);
	}

	private AbstractNum? FollowAbstractNumLinks(AbstractNum abs)
	{
		// If it already has levels, nothing to do.
		if (abs.Elements<Level>().Any())
			return abs;

		// 1) numStyleLink (abstractNumId may point to a linked style without levels)
		var numStyleLinkId = abs.NumberingStyleLink?.Val?.Value; // maps to <w:numStyleLink>
		if (!string.IsNullOrEmpty(numStyleLinkId))
		{
			var linked = ResolveAbstractFromNumberingStyleId(numStyleLinkId!);
			if (linked != null)
				return linked;
		}

		// 2) styleLink (also common)
		var styleLinkId = abs.StyleLink?.Val?.Value; // maps to <w:styleLink>
		if (!string.IsNullOrEmpty(styleLinkId))
		{
			var linked = ResolveAbstractFromNumberingStyleId(styleLinkId!);
			if (linked != null)
				return linked;
		}

		return null;
	}

	private AbstractNum? ResolveAbstractFromNumberingStyleId(string numberingStyleId)
	{
		// styles.xml: numbering style PatentSpecificationList has numId=2
		if (!_styleById.TryGetValue(numberingStyleId, out var style))
			return null;

		var numId = style.StyleParagraphProperties?
						.NumberingProperties?
						.NumberingId?.Val?.Value;

		if (numId == null || numId.Value == 0)
			return null;

		if (!_numById.TryGetValue(numId.Value, out var num))
			return null;

		var absId = num.AbstractNumId?.Val?.Value;
		if (absId == null)
			return null;

		return _absById.TryGetValue(absId.Value, out var abs) ? abs : null;
	}
}



public sealed class DxpListTracker
{
	private readonly Dictionary<int, int?[]> _counters = new();

	private int?[] GetArr(int numId)
	{
		if (!_counters.TryGetValue(numId, out var arr))
		{
			arr = Enumerable.Repeat<int?>(null, 9).ToArray();
			_counters[numId] = arr;
		}
		return arr;
	}

	public int NextIndex(int numId, int ilvl, int startAt)
	{
		var arr = GetArr(numId);

		for (int d = ilvl + 1; d < arr.Length; d++)
			arr[d] = null;

		arr[ilvl] = arr[ilvl] == null ? startAt : arr[ilvl]!.Value + 1;
		return arr[ilvl]!.Value;
	}

	public int? GetCurrent(int numId, int ilvl)
		=> GetArr(numId)[ilvl];

	public void ClearAll() => _counters.Clear();
}

public static class DxpListMarkerFormatter
{
	private static int? TryGetCustomDecimalWidth(Level lvl)
	{
		// The "custom" numFmt is in mc:AlternateContent and may not be surfaced by the SDK as lvl.NumberingFormat.
		// So parse from the raw XML.
		var xml = lvl.OuterXml;

		// Look for: w:numFmt w:val="custom" w:format="0001, 0002, ..."
		const string key = "w:format=\"";
		var idx = xml.IndexOf("w:numFmt", StringComparison.Ordinal);
		if (idx < 0)
			return null;

		// quick/robust-enough scan
		var customIdx = xml.IndexOf("w:val=\"custom\"", idx, StringComparison.Ordinal);
		if (customIdx < 0)
			return null;

		var fmtIdx = xml.IndexOf(key, customIdx, StringComparison.Ordinal);
		if (fmtIdx < 0)
			return null;

		fmtIdx += key.Length;
		var end = xml.IndexOf('"', fmtIdx);
		if (end < 0)
			return null;

		var format = xml.Substring(fmtIdx, end - fmtIdx); // e.g. "0001, 0002, 0003, ."
														  // Take the first token up to comma/space; "0001" => width 4
		var token = format.Split(',', ' ', ';').FirstOrDefault(t => t.Length > 0);
		if (string.IsNullOrEmpty(token))
			return null;

		// Count leading zeros in first token (0001 => width 4). If no zeros, ignore.
		var width = token.Length;
		if (width <= 1)
			return null;

		// Ensure it's a numeric mask like 0001
		if (token.Any(ch => ch != '0' && ch != '1' && ch != '2' && ch != '3' && ch != '4' && ch != '5' && ch != '6' && ch != '7' && ch != '8' && ch != '9'))
			return null;

		return width;
	}

	public static string? TryBuildMarkerText(
		DxpStyleEffectiveNumPr numPr,
		DxpNumberingResolver nr,
		DxpListTracker tracker)
	{
		var resolved = nr.ResolveLevel(numPr.NumId, numPr.Ilvl);
		if (resolved == null)
			return null;

		var (abs, lvl, startAt) = resolved.Value;

		var fmt = lvl.NumberingFormat?.Val?.Value;     // w:numFmt (may be null)
		var lvlText = lvl.LevelText?.Val?.Value ?? ""; // w:lvlText ("" is meaningful!)

		// IMPORTANT: Word semantics
		// - numFmt="none" means "this level displays no numbering"
		// - lvlText may be "" (show nothing) or "–" (show a dash), etc.
		// => Do NOT invent "%1." here.
			if (fmt == NumberFormatValues.None)
			{
				// Do NOT advance counters for a "none" level; these are usually continuation lines.
				// Behavior can be revisited if documents expect these to advance.
				return string.IsNullOrEmpty(lvlText) ? null : lvlText;
			}

		// If there is no lvlText for a non-none format, do NOT guess.
		if (string.IsNullOrEmpty(lvlText))
			return null;

		// Increment current level (normal list item / bullet / numbered)
		tracker.NextIndex(numPr.NumId, numPr.Ilvl, startAt);

		// If bullet, Word stores the glyph directly in lvlText (and font is in lvl.rPr).
		if (fmt == NumberFormatValues.Bullet)
		{
			var font = GetBulletFont(lvl);
			if (!string.IsNullOrEmpty(font) && lvlText.Length > 0)
				return $"""<span style="font-family:{font}">{lvlText}</span>""";
			return lvlText;
		}

		string result = lvlText;

		for (int i = 0; i < 9; i++)
		{
			var placeholder = "%" + (i + 1).ToString(CultureInfo.InvariantCulture);
			if (result.IndexOf(placeholder, StringComparison.Ordinal) < 0)
				continue;

			var cur = tracker.GetCurrent(numPr.NumId, i);
			if (cur == null)
			{
				// If a placeholder references a level that hasn't been started yet,
				// Word's behavior is nuanced; "0" is safer than crashing.
				cur = 0;
			}

			var refLvl = nr.ResolveLevel(numPr.NumId, i)?.lvl;
			var refFmt = refLvl?.NumberingFormat?.Val?.Value ?? NumberFormatValues.Decimal;

			// Custom "0001" padding (from mc:AlternateContent) if present
			int? customWidth = refLvl != null ? TryGetCustomDecimalWidth(refLvl) : null;

			string formatted = customWidth != null
				? cur.Value.ToString(CultureInfo.InvariantCulture).PadLeft(customWidth.Value, '0')
				: FormatNumber(cur.Value, refFmt);

			result = result.Replace(placeholder, formatted);

		}

		return result;
	}

	private static string FormatNumber(int n, NumberFormatValues fmt)
	{
		if (fmt == NumberFormatValues.Decimal)
			return n.ToString(CultureInfo.InvariantCulture);
		if (fmt == NumberFormatValues.UpperRoman)
			return ToRoman(n).ToUpperInvariant();
		if (fmt == NumberFormatValues.LowerRoman)
			return ToRoman(n).ToLowerInvariant();
		if (fmt == NumberFormatValues.UpperLetter)
			return ToAlpha(n).ToUpperInvariant();
		if (fmt == NumberFormatValues.LowerLetter)
			return ToAlpha(n).ToLowerInvariant();

		return n.ToString(CultureInfo.InvariantCulture);
	}

	private static string ToAlpha(int n)
	{
		// 1 -> A, 26 -> Z, 27 -> AA ...
		if (n <= 0)
			return n.ToString(CultureInfo.InvariantCulture);
		var s = "";
		while (n > 0)
		{
			n--; // 1-based
			s = (char)('A' + n % 26) + s;
			n /= 26;
		}
		return s;
	}

	private static string ToRoman(int number)
	{
		if (number <= 0)
			return number.ToString(CultureInfo.InvariantCulture);
		var map = new (int v, string s)[] {
			(1000,"M"),(900,"CM"),(500,"D"),(400,"CD"),
			(100,"C"),(90,"XC"),(50,"L"),(40,"XL"),
			(10,"X"),(9,"IX"),(5,"V"),(4,"IV"),(1,"I")
		};
		var result = "";
		foreach (var (v, s) in map)
			while (number >= v)
			{ result += s; number -= v; }
		return result;
	}

	private static string? GetBulletFont(Level lvl)
	{
		var fonts = lvl.NumberingSymbolRunProperties?.RunFonts;
		if (fonts == null)
			return null;

		return fonts.Ascii?.Value
			?? fonts.HighAnsi?.Value;
	}
}


public class DxpLists
{
	private DxpNumberingResolver? _num;
	private readonly DxpListTracker _lists = new();

	public DxpLists()
	{
	}

	internal void Init(WordprocessingDocument doc)
	{
		_num = new DxpNumberingResolver(doc);
	}

	internal (string? marker, int? numId, int? iLvl) MaterializeMarker(Paragraph p, IDxpStyleResolver s)
	{
		string? marker = null;
		int? numId = null;
		int? iLvl = null;

		if (_num != null)
		{
			var numPr = s.ResolveEffectiveNumPr(p);
			if (numPr != null)
			{
				marker = DxpListMarkerFormatter.TryBuildMarkerText(numPr, _num, _lists);
				numId = numPr.NumId;
				iLvl = numPr.Ilvl;
			}
		}
		return (marker, numId, iLvl);
	}

	public DxpStyleEffectiveIndentTwips GetIndentation(Paragraph p, IDxpStyleResolver s)
	{
		DxpStyleEffectiveIndentTwips indent = s.GetIndentation(p, _num);
		return indent;
	}
}
