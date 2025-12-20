using System.Text;

public static class SymbolEncoding
{
	// Key: Symbol Encoding code point (byte)
	// Value: one or more Unicode scalar values (int)
	public static readonly IReadOnlyDictionary<byte, int[]> Map =
		new Dictionary<byte, int[]> {
			[0x20] = new[] { 0x0020, 0x00A0 }, // SPACE, NO-BREAK SPACE
			[0x21] = new[] { 0x0021 }, // !
			[0x22] = new[] { 0x2200 }, // FOR ALL
			[0x23] = new[] { 0x0023 }, // #
			[0x24] = new[] { 0x2203 }, // THERE EXISTS
			[0x25] = new[] { 0x0025 }, // %
			[0x26] = new[] { 0x0026 }, // &
			[0x27] = new[] { 0x220B }, // CONTAINS AS MEMBER
			[0x28] = new[] { 0x0028 }, // (
			[0x29] = new[] { 0x0029 }, // )
			[0x2A] = new[] { 0x2217 }, // ASTERISK OPERATOR
			[0x2B] = new[] { 0x002B }, // +
			[0x2C] = new[] { 0x002C }, // ,
			[0x2D] = new[] { 0x2212 }, // MINUS SIGN
			[0x2E] = new[] { 0x002E }, // .
			[0x2F] = new[] { 0x002F }, // /

			[0x30] = new[] { 0x0030 },
			[0x31] = new[] { 0x0031 },
			[0x32] = new[] { 0x0032 },
			[0x33] = new[] { 0x0033 },
			[0x34] = new[] { 0x0034 },
			[0x35] = new[] { 0x0035 },
			[0x36] = new[] { 0x0036 },
			[0x37] = new[] { 0x0037 },
			[0x38] = new[] { 0x0038 },
			[0x39] = new[] { 0x0039 },

			[0x3A] = new[] { 0x003A }, // :
			[0x3B] = new[] { 0x003B }, // ;
			[0x3C] = new[] { 0x003C }, // <
			[0x3D] = new[] { 0x003D }, // =
			[0x3E] = new[] { 0x003E }, // >
			[0x3F] = new[] { 0x003F }, // ?
			[0x40] = new[] { 0x2245 }, // APPROXIMATELY EQUAL TO

			[0x41] = new[] { 0x0391 }, // Α
			[0x42] = new[] { 0x0392 }, // Β
			[0x43] = new[] { 0x03A7 }, // Χ
			[0x44] = new[] { 0x0394, 0x2206 }, // Δ, INCREMENT
			[0x45] = new[] { 0x0395 }, // Ε
			[0x46] = new[] { 0x03A6 }, // Φ
			[0x47] = new[] { 0x0393 }, // Γ
			[0x48] = new[] { 0x0397 }, // Η
			[0x49] = new[] { 0x0399 }, // Ι
			[0x4A] = new[] { 0x03D1 }, // ϑ (theta symbol)
			[0x4B] = new[] { 0x039A }, // Κ
			[0x4C] = new[] { 0x039B }, // Λ
			[0x4D] = new[] { 0x039C }, // Μ
			[0x4E] = new[] { 0x039D }, // Ν
			[0x4F] = new[] { 0x039F }, // Ο
			[0x50] = new[] { 0x03A0 }, // Π
			[0x51] = new[] { 0x0398 }, // Θ
			[0x52] = new[] { 0x03A1 }, // Ρ
			[0x53] = new[] { 0x03A3 }, // Σ
			[0x54] = new[] { 0x03A4 }, // Τ
			[0x55] = new[] { 0x03A5 }, // Υ
			[0x56] = new[] { 0x03C2 }, // ς (final sigma)
			[0x57] = new[] { 0x03A9, 0x2126 }, // Ω, OHM SIGN
			[0x58] = new[] { 0x039E }, // Ξ
			[0x59] = new[] { 0x03A8 }, // Ψ
			[0x5A] = new[] { 0x0396 }, // Ζ

			[0x5B] = new[] { 0x005B }, // [
			[0x5C] = new[] { 0x2234 }, // THEREFORE
			[0x5D] = new[] { 0x005D }, // ]
			[0x5E] = new[] { 0x22A5 }, // UP TACK
			[0x5F] = new[] { 0x005F }, // _
			[0x60] = new[] { 0xF8E5 }, // RADICAL EXTENDER (CUS)

			[0x61] = new[] { 0x03B1 }, // α
			[0x62] = new[] { 0x03B2 }, // β
			[0x63] = new[] { 0x03C7 }, // χ
			[0x64] = new[] { 0x03B4 }, // δ
			[0x65] = new[] { 0x03B5 }, // ε
			[0x66] = new[] { 0x03C6 }, // φ
			[0x67] = new[] { 0x03B3 }, // γ
			[0x68] = new[] { 0x03B7 }, // η
			[0x69] = new[] { 0x03B9 }, // ι
			[0x6A] = new[] { 0x03D5 }, // ϕ (phi symbol)
			[0x6B] = new[] { 0x03BA }, // κ
			[0x6C] = new[] { 0x03BB }, // λ
			[0x6D] = new[] { 0x00B5, 0x03BC }, // MICRO SIGN, μ
			[0x6E] = new[] { 0x03BD }, // ν
			[0x6F] = new[] { 0x03BF }, // ο
			[0x70] = new[] { 0x03C0 }, // π
			[0x71] = new[] { 0x03B8 }, // θ
			[0x72] = new[] { 0x03C1 }, // ρ
			[0x73] = new[] { 0x03C3 }, // σ
			[0x74] = new[] { 0x03C4 }, // τ
			[0x75] = new[] { 0x03C5 }, // υ
			[0x76] = new[] { 0x03D6 }, // ϖ (pi symbol)  (note: file says "omega1")
			[0x77] = new[] { 0x03C9 }, // ω
			[0x78] = new[] { 0x03BE }, // ξ
			[0x79] = new[] { 0x03C8 }, // ψ
			[0x7A] = new[] { 0x03B6 }, // ζ

			[0x7B] = new[] { 0x007B }, // {
			[0x7C] = new[] { 0x007C }, // |
			[0x7D] = new[] { 0x007D }, // }
			[0x7E] = new[] { 0x223C }, // TILDE OPERATOR

			[0xA0] = new[] { 0x20AC }, // €
			[0xA1] = new[] { 0x03D2 }, // ϒ
			[0xA2] = new[] { 0x2032 }, // ′
			[0xA3] = new[] { 0x2264 }, // ≤
			[0xA4] = new[] { 0x2044, 0x2215 }, // FRACTION SLASH, DIVISION SLASH
			[0xA5] = new[] { 0x221E }, // ∞
			[0xA6] = new[] { 0x0192 }, // ƒ
			[0xA7] = new[] { 0x2663 }, // ♣
			[0xA8] = new[] { 0x2666 }, // ♦
			[0xA9] = new[] { 0x2665 }, // ♥
			[0xAA] = new[] { 0x2660 }, // ♠
			[0xAB] = new[] { 0x2194 }, // ↔
			[0xAC] = new[] { 0x2190 }, // ←
			[0xAD] = new[] { 0x2191 }, // ↑
			[0xAE] = new[] { 0x2192 }, // →
			[0xAF] = new[] { 0x2193 }, // ↓

			[0xB0] = new[] { 0x00B0 }, // °
			[0xB1] = new[] { 0x00B1 }, // ±
			[0xB2] = new[] { 0x2033 }, // ″
			[0xB3] = new[] { 0x2265 }, // ≥
			[0xB4] = new[] { 0x00D7 }, // ×
			[0xB5] = new[] { 0x221D }, // ∝
			[0xB6] = new[] { 0x2202 }, // ∂
			[0xB7] = new[] { 0x2022 }, // •
			[0xB8] = new[] { 0x00F7 }, // ÷
			[0xB9] = new[] { 0x2260 }, // ≠
			[0xBA] = new[] { 0x2261 }, // ≡
			[0xBB] = new[] { 0x2248 }, // ≈
			[0xBC] = new[] { 0x2026 }, // …
			[0xBD] = new[] { 0xF8E6 }, // VERTICAL ARROW EXTENDER (CUS)
			[0xBE] = new[] { 0xF8E7 }, // HORIZONTAL ARROW EXTENDER (CUS)
			[0xBF] = new[] { 0x21B5 }, // ↵

			[0xC0] = new[] { 0x2135 }, // ℵ
			[0xC1] = new[] { 0x2111 }, // ℑ
			[0xC2] = new[] { 0x211C }, // ℜ
			[0xC3] = new[] { 0x2118 }, // ℘
			[0xC4] = new[] { 0x2297 }, // ⊗
			[0xC5] = new[] { 0x2295 }, // ⊕
			[0xC6] = new[] { 0x2205 }, // ∅
			[0xC7] = new[] { 0x2229 }, // ∩
			[0xC8] = new[] { 0x222A }, // ∪
			[0xC9] = new[] { 0x2283 }, // ⊃
			[0xCA] = new[] { 0x2287 }, // ⊇
			[0xCB] = new[] { 0x2284 }, // ⊄
			[0xCC] = new[] { 0x2282 }, // ⊂
			[0xCD] = new[] { 0x2286 }, // ⊆
			[0xCE] = new[] { 0x2208 }, // ∈
			[0xCF] = new[] { 0x2209 }, // ∉

			[0xD0] = new[] { 0x2220 }, // ∠
			[0xD1] = new[] { 0x2207 }, // ∇
			[0xD2] = new[] { 0xF6DA }, // REGISTERED SIGN SERIF (CUS)
			[0xD3] = new[] { 0xF6D9 }, // COPYRIGHT SIGN SERIF (CUS)
			[0xD4] = new[] { 0xF6DB }, // TRADE MARK SIGN SERIF (CUS)
			[0xD5] = new[] { 0x220F }, // ∏
			[0xD6] = new[] { 0x221A }, // √
			[0xD7] = new[] { 0x22C5 }, // ⋅
			[0xD8] = new[] { 0x00AC }, // ¬
			[0xD9] = new[] { 0x2227 }, // ∧
			[0xDA] = new[] { 0x2228 }, // ∨
			[0xDB] = new[] { 0x21D4 }, // ⇔
			[0xDC] = new[] { 0x21D0 }, // ⇐
			[0xDD] = new[] { 0x21D1 }, // ⇑
			[0xDE] = new[] { 0x21D2 }, // ⇒
			[0xDF] = new[] { 0x21D3 }, // ⇓

			[0xE0] = new[] { 0x25CA }, // ◊
			[0xE1] = new[] { 0x2329 }, // 〈
			[0xE2] = new[] { 0xF8E8 }, // REGISTERED SIGN SANS SERIF (CUS)
			[0xE3] = new[] { 0xF8E9 }, // COPYRIGHT SIGN SANS SERIF (CUS)
			[0xE4] = new[] { 0xF8EA }, // TRADE MARK SIGN SANS SERIF (CUS)
			[0xE5] = new[] { 0x2211 }, // ∑
			[0xE6] = new[] { 0xF8EB }, // LEFT PAREN TOP (CUS)
			[0xE7] = new[] { 0xF8EC }, // LEFT PAREN EXTENDER (CUS)
			[0xE8] = new[] { 0xF8ED }, // LEFT PAREN BOTTOM (CUS)
			[0xE9] = new[] { 0xF8EE }, // LEFT SQUARE BRACKET TOP (CUS)
			[0xEA] = new[] { 0xF8EF }, // LEFT SQUARE BRACKET EXTENDER (CUS)
			[0xEB] = new[] { 0xF8F0 }, // LEFT SQUARE BRACKET BOTTOM (CUS)
			[0xEC] = new[] { 0xF8F1 }, // LEFT CURLY BRACKET TOP (CUS)
			[0xED] = new[] { 0xF8F2 }, // LEFT CURLY BRACKET MID (CUS)
			[0xEE] = new[] { 0xF8F3 }, // LEFT CURLY BRACKET BOTTOM (CUS)
			[0xEF] = new[] { 0xF8F4 }, // CURLY BRACKET EXTENDER (CUS)

			[0xF1] = new[] { 0x232A }, // 〉
			[0xF2] = new[] { 0x222B }, // ∫
			[0xF3] = new[] { 0x2320 }, // ⌠
			[0xF4] = new[] { 0xF8F5 }, // INTEGRAL EXTENDER (CUS)
			[0xF5] = new[] { 0x2321 }, // ⌡
			[0xF6] = new[] { 0xF8F6 }, // RIGHT PAREN TOP (CUS)
			[0xF7] = new[] { 0xF8F7 }, // RIGHT PAREN EXTENDER (CUS)
			[0xF8] = new[] { 0xF8F8 }, // RIGHT PAREN BOTTOM (CUS)
			[0xF9] = new[] { 0xF8F9 }, // RIGHT SQUARE BRACKET TOP (CUS)
			[0xFA] = new[] { 0xF8FA }, // RIGHT SQUARE BRACKET EXTENDER (CUS)
			[0xFB] = new[] { 0xF8FB }, // RIGHT SQUARE BRACKET BOTTOM (CUS)
			[0xFC] = new[] { 0xF8FC }, // RIGHT CURLY BRACKET TOP (CUS)
			[0xFD] = new[] { 0xF8FD }, // RIGHT CURLY BRACKET MID (CUS)
			[0xFE] = new[] { 0xF8FE }, // RIGHT CURLY BRACKET BOTTOM (CUS)
		};

	public static string? ToUnicode(byte symbolCode)
	{
		if (!Map.TryGetValue(symbolCode, out var cps) || cps.Length == 0)
			return null;

		return char.ConvertFromUtf32(cps[0]);
	}

	public static byte[]? ToUtf8Bytes(byte symbolCode)
	{
		var s = ToUnicode(symbolCode);
		return s is null ? null : Encoding.UTF8.GetBytes(s);
	}
}
