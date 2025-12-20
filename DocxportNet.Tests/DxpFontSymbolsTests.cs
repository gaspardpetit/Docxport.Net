using DocxportNet;
using Xunit;

namespace DocxportNet.Tests;

public class DxpFontSymbolsTests
{
	[Fact]
	public void Wingdings_substitutes_string_codes()
	{
		// 0x41 ('A') -> ✌, 0x42 ('B') -> 👌 in the Wingdings mapping
		var result = DxpFontSymbols.Substitute("Wingdings", "\u0041\u0042");
		Assert.Equal("✌👌", result);
	}

	[Fact]
	public void Symbol_bullet_maps_to_unicode_bullet()
	{
		var bullet = DxpFontSymbols.Substitute("Symbol", (char)0xB7);
		Assert.Equal("•", bullet);
	}

	[Fact]
	public void ZapfDingbats_pointing_hand_maps()
	{
		var hand = DxpFontSymbols.Substitute("Zapf Dingbats", (char)0x2A);
		Assert.Equal("☛", hand);
	}

	[Fact]
	public void Webdings_cat_maps_to_unicode_cat()
	{
		var cat = DxpFontSymbols.Substitute("Webdings", (char)0xF6);
		Assert.Equal("🐈", cat);
	}

	[Fact]
	public void Wingdings2_left_point_maps()
	{
		var hand = DxpFontSymbols.Substitute("Wingdings 2", (char)0x42);
		Assert.Equal("👈", hand);
	}

	[Fact]
	public void Wingdings3_arrow_maps()
	{
		var arrows = DxpFontSymbols.Substitute("Wingdings 3", "\u0030\u0031");
		Assert.Equal("⭽⭤", arrows);
	}
}
