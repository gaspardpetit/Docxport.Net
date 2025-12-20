using System.Globalization;
using System.Text;
using DocxportNet.symbols;

namespace DocxportNet;

/// <summary>
/// Public helper for translating common symbol fonts (Symbol, Zapf Dingbats, Webdings, Wingdings variants)
/// into their Unicode equivalents.
/// </summary>
public static class DxpFontSymbols
{
	private static readonly Dictionary<string, string[]> _fontTables = new(StringComparer.OrdinalIgnoreCase) {
		["Symbol"] = SymbolEncoding.Table,
		["ZapfDingbats"] = ZapfDingbatsEncoding.Table,
		["Zapf Dingbats"] = ZapfDingbatsEncoding.Table,
		["Webdings"] = WebdingsEncoding.Table,
		["Wingdings"] = WingdingsEncoding.Table,
		["Wingdings 2"] = Wingdings2Map.Table,
		["Wingdings 3"] = Wingdings3Encoding.Table,
	};

	/// <summary>
	/// Translate a symbol-font encoded string into Unicode. Unknown characters are passed through unchanged.
	/// </summary>
	public static string Substitute(string? fontName, string? text, char? replacementForNonPrintable = null)
	{
		if (string.IsNullOrEmpty(text))
			return string.Empty;

		DxpFontSymbolConverter? converter = GetSymbolConverter(fontName);
		if (converter == null)
			return text!;

		var sb = new StringBuilder(text!.Length);
		foreach (char ch in text)
			sb.Append(converter.Substitute(ch, replacementForNonPrintable));

		return sb.ToString();
	}

	/// <summary>
	/// Translate a single symbol-font encoded character into Unicode. Unknown characters are returned unchanged.
	/// </summary>
	public static string Substitute(string? fontName, char ch, char? replacementForNonPrintable = null)
	{
		// Outside 8-bit code pages the font mappings do not apply.
		if (ch > byte.MaxValue)
			return ch.ToString();

		DxpFontSymbolConverter? converter = GetSymbolConverter(fontName);
		if (converter == null)
			return ch.ToString();

		return converter.Substitute(ch, replacementForNonPrintable);
	}

	/// <summary>
	/// Get a converter for the given symbol font, or null if the font is not supported.
	/// Useful for reusing the lookup without re-validating the font name each call.
	/// </summary>
	public static DxpFontSymbolConverter? GetSymbolConverter(string? fontName)
	{
		if (string.IsNullOrEmpty(fontName))
			return null;

		if (_fontTables.TryGetValue(fontName!.Trim(), out string[]? table) == false)
			return null;

		return new DxpFontSymbolConverter(table);
	}
}

public class DxpFontSymbolConverter
{
	private string[] _table;

	/// <summary>
	/// Create a converter bound to a specific symbol font table. Preferred when reusing many lookups.
	/// </summary>
	public DxpFontSymbolConverter(string[] table)
	{
		_table = table;
	}

	/// <summary>
	/// Translate a single symbol-font encoded character into Unicode, optionally replacing non-printable glyphs.
	/// </summary>
	public string Substitute(char ch, char? replacementForNonPrintable)
	{
		// Map using the low byte so private-use/codepage escapes (e.g., U+F0B7) still resolve.
		string sub = _table[(byte)ch];

		if (replacementForNonPrintable != null && IsPrintable(sub) == false)
			return replacementForNonPrintable.ToString()!;

		return sub;
	}

	/// <summary>
	/// Determine whether a string consists only of printable (non-control/non-private) Unicode characters.
	/// </summary>
	public static bool IsPrintable(string value)
	{
		for (int i = 0; i < value.Length; i++)
		{
			int codepoint = char.ConvertToUtf32(value, i);
			if (char.IsSurrogatePair(value, i))
				i++; // skip the low surrogate

			// CharUnicodeInfo lacks an int overload in netstandard2.0; use the string+index overload.
			var cat = CharUnicodeInfo.GetUnicodeCategory(char.ConvertFromUtf32(codepoint), 0);
			if (cat == UnicodeCategory.Control ||
				cat == UnicodeCategory.Format ||
				cat == UnicodeCategory.OtherNotAssigned ||
				cat == UnicodeCategory.Surrogate ||
				cat == UnicodeCategory.PrivateUse)
				return false;
		}
		return true;
	}
}
