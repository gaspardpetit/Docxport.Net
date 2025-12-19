using DocxportNet.api;

public sealed class DxpStyleTracker : IDxpStyleTracker
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

	public void ApplyStyle(DxpStyleEffectiveRunStyle next, IDxpVisitor v)
	{
		// 0) Ensure font wrapper exists and is outermost
		EnsureFontOutermost(next, v);

		// 1) Compute desired inner set from next
		var desired = DesiredInner(next);

		// 2) Apply inner delta with proper nesting and minimal operations
		ApplyInnerDelta(desired, v);

		_current = next;
	}

	public void ResetStyle(IDxpVisitor v)
	{
		// Close inner styles (top->bottom)
		for (int i = _open.Count - 1; i >= 0; i--)
			End(_open[i], v);
		_open.Clear();

		// Close font wrapper last (outermost)
		if (_fontOpen)
		{
			v.StyleFontEnd();
			_fontOpen = false;
		}

		_current = null;
		_fontName = null;
		_fontSizeHp = null;
	}

	// ---------------- Font handling (always outermost) ----------------

	private void EnsureFontOutermost(DxpStyleEffectiveRunStyle next, IDxpVisitor v)
	{
		bool fontChanged =
			!_fontOpen ||
			!string.Equals(_fontName, next.FontName, StringComparison.Ordinal) ||
			_fontSizeHp != next.FontSizeHalfPoints;

		if (!_fontOpen)
		{
			v.StyleFontBegin(next.FontName, next.FontSizeHalfPoints);
			_fontOpen = true;
			_fontName = next.FontName;
			_fontSizeHp = next.FontSizeHalfPoints;
			return;
		}

		if (!fontChanged)
			return;

		// Close all inner tags so font stays outward
		for (int i = _open.Count - 1; i >= 0; i--)
			End(_open[i], v);
		_open.Clear();

		// Switch font wrapper
		v.StyleFontEnd();
		v.StyleFontBegin(next.FontName, next.FontSizeHalfPoints);

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

	private void ApplyInnerDelta(HashSet<K> desired, IDxpVisitor v)
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
			End(_open[i], v);
			_open.RemoveAt(i);
		}

		// Open any desired styles not yet open, in canonical order
		foreach (var k in CanonicalOrder)
		{
			if (!desired.Contains(k))
				continue;
			if (_open.Contains(k))
				continue; // should only be in the kept prefix, but safe

			Begin(k, v);
			_open.Add(k);
		}
	}

	private static void Begin(K k, IDxpVisitor v)
	{
		switch (k)
		{
			case K.Bold:
				v.StyleBoldBegin();
				break;
			case K.Italic:
				v.StyleItalicBegin();
				break;
			case K.Underline:
				v.StyleUnderlineBegin();
				break;
			case K.Strike:
				v.StyleStrikeBegin();
				break;
			case K.DoubleStrike:
				v.StyleDoubleStrikeBegin();
				break;
			case K.Superscript:
				v.StyleSuperscriptBegin();
				break;
			case K.Subscript:
				v.StyleSubscriptBegin();
				break;
			case K.AllCaps:
				v.StyleAllCapsBegin();
				break;
			case K.SmallCaps:
				v.StyleSmallCapsBegin();
				break;
			default:
				throw new ArgumentOutOfRangeException(nameof(k));
		}
	}

	private static void End(K k, IDxpVisitor v)
	{
		switch (k)
		{
			case K.Bold:
				v.StyleBoldEnd();
				break;
			case K.Italic:
				v.StyleItalicEnd();
				break;
			case K.Underline:
				v.StyleUnderlineEnd();
				break;
			case K.Strike:
				v.StyleStrikeEnd();
				break;
			case K.DoubleStrike:
				v.StyleDoubleStrikeEnd();
				break;
			case K.Superscript:
				v.StyleSuperscriptEnd();
				break;
			case K.Subscript:
				v.StyleSubscriptEnd();
				break;
			case K.AllCaps:
				v.StyleAllCapsEnd();
				break;
			case K.SmallCaps:
				v.StyleSmallCapsEnd();
				break;
			default:
				throw new ArgumentOutOfRangeException(nameof(k));
		}
	}
}

