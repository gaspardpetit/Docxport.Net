![NuGet Version](https://img.shields.io/nuget/v/DocxportNet)
[![Build & Tests](https://github.com/gaspardpetit/Docxport.Net/actions/workflows/run-tests.yml/badge.svg)](https://github.com/gaspardpetit/Docxport.Net/actions/workflows/run-tests.yml)

# Docxport.Net

Docxport.Net is a .NET library for walking DOCX documents and exporting them to friendly formats. Today it focuses on Markdown (rich and plain), with full handling of:

- Tracked changes (accept/reject/inline views)
- Lists with proper markers and indentation
- Tables
- Comments and threaded replies
- Headers and footers
- Images/drawings
- Bookmarks, hyperlinks, fields, and more

It also ships a small utility that translates legacy “symbol fonts” (Symbol, Zapf Dingbats, Webdings, Wingdings) into modern Unicode.

## Supported symbol fonts

- Symbol
- Zapf Dingbats
- Webdings
- Wingdings
- Wingdings 2
- Wingdings 3

## Symbol font to Unicode

Use `DxpFontSymbols` when you encounter text that was encoded with a symbol font and you want plain Unicode output.

```csharp
// Convert a whole string using the font hint.
// 0x41 ('A') -> ✌, 0x42 ('B') -> 👌 in the Wingdings table.
string text = DxpFontSymbols.Substitute("Wingdings", "\u0041\u0042"); // => "✌👌"
```

```csharp
// Convert a single character; falls back to the original if unmapped
string bullet = DxpFontSymbols.Substitute("Symbol", (char)0xB7); // → "•"
```

Unknown characters are returned unchanged. You can optionally supply a replacement character for non-printable glyphs:

```csharp
// Replace non-printable/control entries with '?'
string safe = DxpFontSymbols.Substitute("Symbol", "\u0001\u00B7", '?'); // => "?•"
```

### Reusing a converter

If you need to translate many strings from the same font or probe whether a font is supported, reuse a converter instance:

```csharp
var converter = DxpFontSymbols.GetSymbolConverter("Webdings");
if (converter != null)
{
    string cat = converter.Substitute((char)0xF6, '?'); // 🐈
    string arrows = converter.Substitute((char)0x3C);   // ↔
}
```

### Common mappings

- Symbol bullet: `DxpFontSymbols.Substitute("Symbol", (char)0xB7)` → `•`
- Webdings cat: `DxpFontSymbols.Substitute("Webdings", (char)0xF6)` → `🐈`
- Wingdings peace/ok: `DxpFontSymbols.Substitute("Wingdings", "\u0041\u0042")` → `✌👌`
- Wingdings 2 left point: `DxpFontSymbols.Substitute("Wingdings 2", (char)0x42)` → `👈`
- Wingdings 3 arrows: `DxpFontSymbols.Substitute("Wingdings 3", "\u0030\u0031")` → `⭽⭤`
