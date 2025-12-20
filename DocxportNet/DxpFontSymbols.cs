using System;
using System.Collections.Generic;
using System.Text;
using DocxportNet.symbols;

namespace DocxportNet;

/// <summary>
/// Public helper for translating common symbol fonts (Symbol, Zapf Dingbats, Webdings, Wingdings variants)
/// into their Unicode equivalents.
/// </summary>
public static class DxpFontSymbols
{
	private static readonly Dictionary<string, Func<byte, string?>> _fontTranslators = new(StringComparer.OrdinalIgnoreCase) {
		["Symbol"] = SymbolEncoding.ToUnicode,
		["ZapfDingbats"] = ZapfDingbatsEncoding.ToUnicode,
		["Zapf Dingbats"] = ZapfDingbatsEncoding.ToUnicode,
		["Webdings"] = WebdingsEncoding.ToUnicode,
		["Wingdings"] = WingdingsEncoding.ToUnicode,
		["Wingdings 2"] = Wingdings2Map.ToUnicode,
		["Wingdings 3"] = Wingdings3Encoding.ToUnicode,
	};

	/// <summary>
	/// Translate a symbol-font encoded string into Unicode. Unknown characters are passed through unchanged.
	/// </summary>
	public static string Substitute(string? fontName, string? text)
	{
		if (string.IsNullOrEmpty(text))
			return string.Empty;

		var sb = new StringBuilder(text.Length);
		foreach (char ch in text)
			sb.Append(Substitute(fontName, ch));
		return sb.ToString();
	}

	/// <summary>
	/// Translate a single symbol-font encoded character into Unicode. Unknown characters are returned unchanged.
	/// </summary>
	public static string Substitute(string? fontName, char ch)
	{
		// Outside 8-bit code pages the font mappings do not apply.
		if (ch > byte.MaxValue)
			return ch.ToString();

		if (!string.IsNullOrEmpty(fontName))
		{
			if (_fontTranslators.TryGetValue(fontName.Trim(), out var translator))
			{
				var translated = translator((byte)ch);
				if (!string.IsNullOrEmpty(translated))
					return translated!;
			}
		}

		// Fallback: try generic Symbol and Zapf Dingbats even without a font hint.
		var fallback = TryFallback((byte)ch);
		return !string.IsNullOrEmpty(fallback) ? fallback! : ch.ToString();
	}

	private static string? TryFallback(byte code)
	{
		var symbol = SymbolEncoding.ToUnicode(code);
		if (!string.IsNullOrEmpty(symbol))
			return symbol;

		var zapf = ZapfDingbatsEncoding.ToUnicode(code);
		if (!string.IsNullOrEmpty(zapf))
			return zapf;

		return null;
	}
}
