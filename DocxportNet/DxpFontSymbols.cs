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

		var sb = new StringBuilder(text!.Length);
		foreach (char ch in text)
			sb.Append(Substitute(fontName, ch, replacementForNonPrintable));
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

		if (string.IsNullOrEmpty(fontName))
			return ch.ToString();

		if (_fontTables.TryGetValue(fontName!.Trim(), out var table) == false)
			return ch.ToString();

		string sub = table[(byte)ch];

		if (replacementForNonPrintable != null && IsPrintable(sub) == false)
			return replacementForNonPrintable.ToString()!;

		return sub;
	}

	private static bool IsPrintable(string value)
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
