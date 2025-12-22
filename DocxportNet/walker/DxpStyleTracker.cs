using DocxportNet.API;

public sealed class DxpStyleTracker : DxpIStyleTracker
{
	private DxpStyleEffectiveRunStyle? _current;

	// Inner styles (everything except font). Keep a canonical order to stabilize nesting.
	// This makes diffs predictable and as small as possible over time.
	private enum K
	{
		AllCaps,
		SmallCaps,
		Superscript,
		Subscript,
		Bold,
		Italic,
		Underline,
		Strike,
		DoubleStrike
	}

	// Bottom -> top (opened order). We enforce canonical order, so this is stable.
	private readonly List<K> _open = new();

	// Font wrapper state
	private bool _fontOpen;
	private string? _fontName;
	private int? _fontSizeHp;

	public void ApplyStyle(DxpStyleEffectiveRunStyle next, DxpIDocumentContext d, DxpIVisitor v)
	{
		// 0) Ensure font wrapper exists and is outermost
		EnsureFontOutermost(next, d, v);

		// 1) Compute desired inner set from next
		var desired = DesiredInner(next);

		// 2) Apply inner delta with proper nesting and minimal operations
		ApplyInnerDelta(desired, d, v);

		_current = next;
	}

	public void ResetStyle(DxpIDocumentContext d, DxpIVisitor v)
	{
		// Close inner styles (top->bottom)
		for (int i = _open.Count - 1; i >= 0; i--)
			End(_open[i], d, v);
		_open.Clear();

		// Close font wrapper last (outermost)
		if (_fontOpen)
		{
			v.StyleFontEnd(d);
			_fontOpen = false;
		}

		_current = null;
		_fontName = null;
		_fontSizeHp = null;
	}

	// ---------------- Font handling (always outermost) ----------------

	private void EnsureFontOutermost(DxpStyleEffectiveRunStyle next, DxpIDocumentContext d, DxpIVisitor v)
	{
		bool fontChanged =
			!_fontOpen ||
			!string.Equals(_fontName, next.FontName, StringComparison.Ordinal) ||
			_fontSizeHp != next.FontSizeHalfPoints;

		if (!_fontOpen)
		{
			v.StyleFontBegin(new DxpFont(next.FontName, next.FontSizeHalfPoints), d);
			_fontOpen = true;
			_fontName = next.FontName;
			_fontSizeHp = next.FontSizeHalfPoints;
			return;
		}

		if (!fontChanged)
			return;

		// Close all inner tags so font stays outward
		for (int i = _open.Count - 1; i >= 0; i--)
			End(_open[i], d, v);
		_open.Clear();

		// Switch font wrapper
		v.StyleFontEnd(d);
		v.StyleFontBegin(new DxpFont(next.FontName, next.FontSizeHalfPoints), d);

		_fontName = next.FontName;
		_fontSizeHp = next.FontSizeHalfPoints;
	}

	// ---------------- Inner delta (well-nested, minimal) ----------------

	private static IReadOnlyList<K> CanonicalOrder { get; } = new[]
	{
		// Mutually exclusive groups are represented as separate keys; DesiredInner enforces exclusivity.
		K.AllCaps, K.SmallCaps,
		K.Superscript, K.Subscript,
		K.Bold, K.Italic, K.Underline, K.Strike, K.DoubleStrike
	};

	private static HashSet<K> DesiredInner(DxpStyleEffectiveRunStyle s)
	{
		var set = new HashSet<K>();

		// caps mutually exclusive
		if (s.AllCaps)
			set.Add(K.AllCaps);
		else if (s.SmallCaps)
			set.Add(K.SmallCaps);

		// vertical alignment mutually exclusive
		if (s.Superscript)
			set.Add(K.Superscript);
		else if (s.Subscript)
			set.Add(K.Subscript);

		if (s.Bold)
			set.Add(K.Bold);
		if (s.Italic)
			set.Add(K.Italic);
		if (s.Underline)
			set.Add(K.Underline);
		if (s.Strike)
			set.Add(K.Strike);
		if (s.DoubleStrike)
			set.Add(K.DoubleStrike);

		return set;
	}

	private void ApplyInnerDelta(HashSet<K> desired, DxpIDocumentContext d, DxpIVisitor v)
	{
			// Strategy:
			// - Find the longest prefix of currently-open styles (in canonical order) that remains desired.
			// - Close anything above that.
			// - Open any remaining desired styles after that, in canonical order.
			// This yields a minimal well-nested delta given canonical nesting.

			// Ensure _open is in canonical order (it always will be if we only open in canonical order)
			int keep = 0;
		while (keep < _open.Count && desired.Contains(_open[keep]))
			keep++;

		// Close from top down to "keep"
		for (int i = _open.Count - 1; i >= keep; i--)
		{
			End(_open[i], d, v);
			_open.RemoveAt(i);
		}

		// Open any desired styles not yet open, in canonical order
		foreach (var k in CanonicalOrder)
		{
			if (!desired.Contains(k))
				continue;
			if (_open.Contains(k))
				continue; // should only be in the kept prefix, but safe

			Begin(k, d, v);
			_open.Add(k);
		}
	}

	private static void Begin(K k, DxpIDocumentContext d, DxpIVisitor v)
	{
		switch (k)
		{
			case K.Bold:
				v.StyleBoldBegin(d);
				break;
			case K.Italic:
				v.StyleItalicBegin(d);
				break;
			case K.Underline:
				v.StyleUnderlineBegin(d);
				break;
			case K.Strike:
				v.StyleStrikeBegin(d);
				break;
			case K.DoubleStrike:
				v.StyleDoubleStrikeBegin(d);
				break;
			case K.Superscript:
				v.StyleSuperscriptBegin(d);
				break;
			case K.Subscript:
				v.StyleSubscriptBegin(d);
				break;
			case K.AllCaps:
				v.StyleAllCapsBegin(d);
				break;
			case K.SmallCaps:
				v.StyleSmallCapsBegin(d);
				break;
			default:
				throw new ArgumentOutOfRangeException(nameof(k));
		}
	}

	private static void End(K k, DxpIDocumentContext d, DxpIVisitor v)
	{
		switch (k)
		{
			case K.Bold:
				v.StyleBoldEnd(d);
				break;
			case K.Italic:
				v.StyleItalicEnd(d);
				break;
			case K.Underline:
				v.StyleUnderlineEnd(d);
				break;
			case K.Strike:
				v.StyleStrikeEnd(d);
				break;
			case K.DoubleStrike:
				v.StyleDoubleStrikeEnd(d);
				break;
			case K.Superscript:
				v.StyleSuperscriptEnd(d);
				break;
			case K.Subscript:
				v.StyleSubscriptEnd(d);
				break;
			case K.AllCaps:
				v.StyleAllCapsEnd(d);
				break;
			case K.SmallCaps:
				v.StyleSmallCapsEnd(d);
				break;
			default:
				throw new ArgumentOutOfRangeException(nameof(k));
		}
	}
}

