using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.api;
using DocxportNet;
using System.Globalization;

namespace l3ia.lapi.services.documents.docx.convert;


public struct DxpEffectiveRunStyleBuilder
{
	public bool? Bold;
	public bool? Italic;
	public bool? Underline;
	public bool? Strike;
	public bool? DoubleStrike;
	public bool? Superscript;
	public bool? Subscript;
	public bool? AllCaps;
	public bool? SmallCaps;
	public string? FontName;
	public int? FontSizeHalfPoints;

	public DxpStyleEffectiveRunStyle ToImmutable() => new(
		Bold ?? false,
		Italic ?? false,
		Underline ?? false,
		Strike ?? false,
		DoubleStrike ?? false,
		Superscript ?? false,
		Subscript ?? false,
		AllCaps ?? false,
		SmallCaps ?? false,
		FontName,
		FontSizeHalfPoints
	);

	public static DxpEffectiveRunStyleBuilder FromDefaults(RunPropertiesBaseStyle? defaults)
	{
		var acc = new DxpEffectiveRunStyleBuilder();
		ApplyRunPropertiesBaseStyle(defaults, ref acc);
		return acc;
	}

	public static void ApplyStyleRunProperties(StyleRunProperties? rp, ref DxpEffectiveRunStyleBuilder acc)
	{
		if (rp == null)
			return;

		ApplyStyle(
			rp.Bold, rp.Italic, rp.Underline, rp.Strike, rp.DoubleStrike,
			rp.VerticalTextAlignment, rp.RunFonts, rp.FontSize,
			rp.Caps, rp.SmallCaps,
			ref acc
		);
	}

	public static void ApplyRunPropertiesBaseStyle(RunPropertiesBaseStyle? rp, ref DxpEffectiveRunStyleBuilder acc)
	{
		if (rp == null)
			return;

		ApplyStyle(
			rp.Bold, rp.Italic, rp.Underline, rp.Strike, rp.DoubleStrike,
			rp.VerticalTextAlignment, rp.RunFonts, rp.FontSize,
			rp.Caps, rp.SmallCaps,
			ref acc
		);
	}

	public static void ApplyRunProperties(RunProperties? rp, ref DxpEffectiveRunStyleBuilder acc)
	{
		if (rp == null)
			return;

		ApplyStyle(
			rp.Bold, rp.Italic, rp.Underline, rp.Strike, rp.DoubleStrike,
			rp.VerticalTextAlignment, rp.RunFonts, rp.FontSize,
			rp.Caps, rp.SmallCaps,
			ref acc
		);
	}


	private static void ApplyStyle(
		Bold? bold,
		Italic? italic,
		Underline? underline,
		Strike? strike,
		DoubleStrike? doubleStrike,
		VerticalTextAlignment? vAlign,
		RunFonts? fonts,
		FontSize? fontSize,
		Caps? caps,
		SmallCaps? smallCaps,
	ref DxpEffectiveRunStyleBuilder acc)
	{
		if (bold != null)
			acc.Bold = IsOn(bold.Val);
		if (italic != null)
			acc.Italic = IsOn(italic.Val);

		if (underline != null)
			acc.Underline = underline.Val != null && underline.Val != UnderlineValues.None;

		if (strike != null)
			acc.Strike = IsOn(strike.Val);
		if (doubleStrike != null)
			acc.DoubleStrike = IsOn(doubleStrike.Val);

		if (vAlign != null)
		{
			var v = vAlign.Val?.Value;
			acc.Superscript = v == VerticalPositionValues.Superscript;
			acc.Subscript = v == VerticalPositionValues.Subscript;
		}

		if (caps != null)
			acc.AllCaps = IsOn(caps.Val);
		if (smallCaps != null)
			acc.SmallCaps = IsOn(smallCaps.Val);

		if (fonts?.Ascii?.Value != null)
			acc.FontName = fonts.Ascii.Value;
		if (fontSize?.Val?.Value != null && int.TryParse(fontSize.Val.Value, out var hp))
			acc.FontSizeHalfPoints = hp;
	}


	private static bool IsOn(OnOffValue? v)
		=> v == null || v.Value; // in WordprocessingML, missing val often means "on"
}


public sealed class DxpStyleResolver : IDxpStyleResolver
{
	private readonly Styles? _styles;
	private readonly Dictionary<string, Style> _byId;

	private readonly RunPropertiesBaseStyle? _docDefaultRunProps;
	private readonly ParagraphPropertiesBaseStyle? _docDefaultParaProps;

	public DxpStyleEffectiveIndentTwips GetIndentation(
	Paragraph p,
	DxpNumberingResolver? nr = null)
	{
		var acc = new IndentAcc();

		// 1) Document defaults (paragraph defaults)
		ApplyIndentation(_docDefaultParaProps?.Indentation, ref acc);

		// 2) Paragraph style chain (base -> ... -> direct style)
		var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
		foreach (var style in EnumerateStyleChainRaw(pStyleId).Reverse())
		{
			var ind = style.StyleParagraphProperties?.Indentation;
			ApplyIndentation(ind, ref acc);
		}

		// 3) Numbering level indentation (if paragraph is in a list)
		if (nr != null)
		{
			var np = ResolveEffectiveNumPr(p);
			if (np != null)
			{
				var resolved = nr.ResolveLevel(np.NumId, np.Ilvl);
				if (resolved != null)
				{
					var lvlInd = resolved.Value.lvl.PreviousParagraphProperties?.Indentation;
					ApplyIndentation(lvlInd, ref acc);
				}
			}
		}

		// 4) Direct paragraph indentation (highest precedence)
		ApplyIndentation(p.ParagraphProperties?.Indentation, ref acc);

		return acc.ToImmutable();
	}

	// ---------------- helpers ----------------

	private struct IndentAcc
	{
		public int? Left;
		public int? Right;
		public int? FirstLine;
		public int? Hanging;

		public DxpStyleEffectiveIndentTwips ToImmutable() => new(Left, Right, FirstLine, Hanging);
	}

	private static int? ReadTwips(StringValue? v)
	{
		if (v?.Value == null)
			return null;
		return int.TryParse(v.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : null;
	}

	private static void ApplyIndentation(Indentation? ind, ref IndentAcc acc)
	{
		if (ind == null)
			return;

		int? left = ReadTwips(ind.Left) ?? ReadTwips(ind.Start);
		int? right = ReadTwips(ind.Right) ?? ReadTwips(ind.End);

		if (left != null)
			acc.Left = left;
		if (right != null)
			acc.Right = right;

		var firstLine = ReadTwips(ind.FirstLine);
		var hanging = ReadTwips(ind.Hanging);

		if (firstLine != null)
			acc.FirstLine = firstLine;
		if (hanging != null)
			acc.Hanging = hanging;
	}




	public DxpStyleResolver(WordprocessingDocument doc)
	{
		_styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
		_byId = _styles?
			.Elements<Style>()
			.Where(s => s.StyleId?.Value != null)
			.ToDictionary(s => s.StyleId!.Value!, s => s)
			?? new Dictionary<string, Style>();

		var docDefaults = _styles?.DocDefaults;
		_docDefaultRunProps = docDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
		_docDefaultParaProps = docDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
	}

	public DxpStyleInfo? GetStyleInfo(string? styleId)
	{
		if (string.IsNullOrEmpty(styleId))
			return null;
		if (!_byId.TryGetValue(styleId!, out var s))
			return null;

		return new DxpStyleInfo(
			StyleId: styleId!,
			Name: s.StyleName?.Val?.Value,
			Type: s.Type?.Value,
			BasedOnStyleId: s.BasedOn?.Val?.Value
		);
	}

	public IReadOnlyList<DxpStyleInfo> GetStyleChain(string? styleId)
	{
		var result = new List<DxpStyleInfo>();
		var seen = new HashSet<string>(StringComparer.Ordinal);

		var current = styleId;
		while (!string.IsNullOrEmpty(current) && seen.Add(current!))
		{
			var info = GetStyleInfo(current);
			if (info == null)
				break;

			result.Add(info);
			current = info.BasedOnStyleId;
		}

		return result; // [direct, parent, grandparent, ...]
	}

	public IReadOnlyList<DxpStyleInfo> GetParagraphStyleChain(Paragraph p)
	{
		var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
		return GetStyleChain(pStyleId);
	}


	public DxpStyleEffectiveRunStyle ResolveRunStyle(Paragraph p, Run r)
	{
		// 1) Start with doc defaults
		var acc = DxpEffectiveRunStyleBuilder.FromDefaults(_docDefaultRunProps);

		// 2) Apply paragraph style chain (paragraph style's rPr affects runs)
		var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
		ApplyParagraphStyleChainRunProps(pStyleId, ref acc);

		// 3) Apply character style chain (rStyle)
		var rStyleId = r.RunProperties?.RunStyle?.Val?.Value;
		ApplyCharacterStyleChainRunProps(rStyleId, ref acc);

		// 4) Apply direct run formatting (highest precedence)
		DxpEffectiveRunStyleBuilder.ApplyRunProperties(r.RunProperties, ref acc);

		return acc.ToImmutable();
	}

	private void ApplyParagraphStyleChainRunProps(string? styleId, ref DxpEffectiveRunStyleBuilder acc)
	{
		foreach (var style in EnumerateStyleChain(styleId).Reverse())
		{
			DxpEffectiveRunStyleBuilder.ApplyStyleRunProperties(style.StyleRunProperties, ref acc);
			var rp = style.StyleParagraphProperties?.GetFirstChild<RunProperties>();
			DxpEffectiveRunStyleBuilder.ApplyRunProperties(rp, ref acc);
		}
	}

	private void ApplyCharacterStyleChainRunProps(string? styleId, ref DxpEffectiveRunStyleBuilder acc)
	{
		foreach (var style in EnumerateStyleChain(styleId).Reverse())
			DxpEffectiveRunStyleBuilder.ApplyStyleRunProperties(style.StyleRunProperties, ref acc);
	}

	private IEnumerable<Style> EnumerateStyleChain(string? styleId)
	{
		// Walk basedOn chain, starting from styleId, stopping on cycles or missing.
		if (string.IsNullOrEmpty(styleId))
			yield break;

		var seen = new HashSet<string>(StringComparer.Ordinal);

		var current = styleId;
		while (!string.IsNullOrEmpty(current) && seen.Add(current!))
		{
			if (!_byId.TryGetValue(current!, out var style))
				yield break;

			yield return style;

			current = style.BasedOn?.Val?.Value;
		}
	}


	public int? GetOutlineLevel(Paragraph p)
	{
		// Direct formatting on the paragraph (highest precedence)
		var direct = p.ParagraphProperties?.OutlineLevel?.Val?.Value;
		if (direct != null)
			return direct; // 0-based

		// From style chain
		var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
		foreach (var s in EnumerateStyleChainRaw(pStyleId)) // Style objects
		{
			var lvl = s.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
			if (lvl != null)
				return lvl;
		}
		return null;
	}

		// helper: raw style chain (same logic as EnumerateStyleChain)
	private IEnumerable<Style> EnumerateStyleChainRaw(string? styleId)
	{
		if (string.IsNullOrEmpty(styleId))
			yield break;
		var seen = new HashSet<string>(StringComparer.Ordinal);
		var current = styleId;

		while (!string.IsNullOrEmpty(current) && seen.Add(current!))
		{
			if (!_byId.TryGetValue(current!, out var style))
				yield break;
			yield return style;
			current = style.BasedOn?.Val?.Value;
		}
	}

	public int? GetHeadingLevel(Paragraph p)
	{
		// OutlineLvl is 0-based; convert to 1-based heading level
		var outline = GetOutlineLevel(p);
		if (outline is >= 0 and <= 8)
			return outline.Value + 2; // shift by 1 so Title can stay at level 1

		// Fallback: name/id heuristics
		var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
		return GetHeadingLevelFromStyleChain(pStyleId);
	}

	public static int? TryGetHeadingLevelFromStyleNameOrId(DxpStyleInfo s)
	{
		// StyleName examples: "Heading 1", "Heading 2"
		// StyleId examples: "Heading1", "Heading2"
		static int? ParseLevel(string? x)
		{
			if (string.IsNullOrWhiteSpace(x))
				return null;

			x = x?.Trim();

			if (string.IsNullOrWhiteSpace(x))
				return null;

			if (x!.Equals(WordBuiltInStyleId.wdStyleTitle, StringComparison.OrdinalIgnoreCase))
				return 1;

			if (x.Equals(WordBuiltInStyleId.wdStyleHeading1, StringComparison.OrdinalIgnoreCase))
				return 2;
			if (x.Equals(WordBuiltInStyleId.wdStyleHeading2, StringComparison.OrdinalIgnoreCase))
				return 3;
			if (x.Equals(WordBuiltInStyleId.wdStyleHeading3, StringComparison.OrdinalIgnoreCase))
				return 4;
			if (x.Equals(WordBuiltInStyleId.wdStyleHeading4, StringComparison.OrdinalIgnoreCase))
				return 5;
			if (x.Equals(WordBuiltInStyleId.wdStyleHeading5, StringComparison.OrdinalIgnoreCase))
				return 6;
			if (x.Equals(WordBuiltInStyleId.wdStyleHeading6, StringComparison.OrdinalIgnoreCase))
				return 7;
			if (x.Equals(WordBuiltInStyleId.wdStyleHeading7, StringComparison.OrdinalIgnoreCase))
				return 8;
			if (x.Equals(WordBuiltInStyleId.wdStyleHeading8, StringComparison.OrdinalIgnoreCase))
				return 9;
			if (x.Equals(WordBuiltInStyleId.wdStyleHeading9, StringComparison.OrdinalIgnoreCase))
				return 10;

			if (x.StartsWith("Heading ", StringComparison.OrdinalIgnoreCase) &&
				int.TryParse(x.Substring("Heading ".Length), out var n1))
				return n1 + 1;

			if (x.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) &&
				int.TryParse(x.Substring("Heading".Length), out var n2))
				return n2 + 1;

			return null;
		}

		return ParseLevel(s.Name) ?? ParseLevel(s.StyleId);
	}

	public int? GetHeadingLevelFromStyleChain(string? pStyleId)
	{
		foreach (var s in GetStyleChain(pStyleId))
		{
			var lvl = TryGetHeadingLevelFromStyleNameOrId(s);
			if (lvl is >= 1 and <= 9)
				return lvl;
		}
		return null;
	}

	public DxpStyleEffectiveRunStyle GetDefaultRunStyle()
	{
		var acc = DxpEffectiveRunStyleBuilder.FromDefaults(_docDefaultRunProps);
		ApplyParagraphStyleChainRunProps(WordBuiltInStyleId.wdStyleNormal, ref acc);
		return acc.ToImmutable();
	}

	public DxpStyleEffectiveNumPr? ResolveEffectiveNumPr(Paragraph p)
	{
		// 1) direct pPr wins, including explicit "no numbering" (numId=0)
		var directNp = p.ParagraphProperties?.NumberingProperties;
		if (directNp != null)
		{
			var directNumId = directNp.NumberingId?.Val?.Value;
			if (directNumId != null)
			{
				if (directNumId.Value == 0)
					return null; // explicit suppression: do NOT consult styles

				var directIlvl = directNp.NumberingLevelReference?.Val?.Value ?? 0;
				return new DxpStyleEffectiveNumPr(directNumId.Value, directIlvl);
			}

			// If there's a direct numPr but no numId, Word still considers it "direct";
			// in that weird case we fall through to styles.
		}

		// 2) from style chain (closest style wins)
		var pStyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

		int? numId = null;
		int? ilvl = null;

		foreach (var style in EnumerateStyleChainRaw(pStyleId)) // direct -> parent -> ...
		{
			var np = style.StyleParagraphProperties?.NumberingProperties;
			if (np == null)
				continue;

			var sid = np.NumberingId?.Val?.Value;
			if (sid != null)
			{
				// If a style explicitly sets numId=0, treat as "no numbering" too.
				if (sid.Value == 0)
					return null;

				if (numId == null)
					numId = sid.Value;
			}

			if (ilvl == null)
				ilvl = np.NumberingLevelReference?.Val?.Value;

			if (numId != null)
				break;
		}

		if (numId == null || numId.Value == 0)
			return null;

		return new DxpStyleEffectiveNumPr(numId.Value, ilvl ?? 0);
	}
}
