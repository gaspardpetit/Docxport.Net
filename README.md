# Docxport.Net

Docxport.Net is a .NET library for walking DOCX documents and emitting friendly formats. It also ships a small utility that translates legacy “symbol fonts” (Symbol, Zapf Dingbats, Webdings, Wingdings) into modern Unicode.

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

Unknown characters are returned unchanged; the helper also tries a Symbol/Zapf Dingbats fallback when no font name is provided.

### Common mappings

- Symbol bullet: `DxpFontSymbols.Substitute("Symbol", (char)0xB7)` → `•`
- Webdings cat: `DxpFontSymbols.Substitute("Webdings", (char)0xF6)` → `🐈`
- Wingdings peace/ok: `DxpFontSymbols.Substitute("Wingdings", "\u0041\u0042")` → `✌👌`
- Wingdings 2 left point: `DxpFontSymbols.Substitute("Wingdings 2", (char)0x42)` → `👈`
- Wingdings 3 arrows: `DxpFontSymbols.Substitute("Wingdings 3", "\u0030\u0031")` → `⭽⭤`
