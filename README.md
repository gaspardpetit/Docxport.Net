![NuGet Version](https://img.shields.io/nuget/v/DocxportNet)
[![Build & Tests](https://github.com/gaspardpetit/Docxport.Net/actions/workflows/run-tests.yml/badge.svg)](https://github.com/gaspardpetit/Docxport.Net/actions/workflows/run-tests.yml)

# Docxport.Net

Docxport.Net is a .NET library for walking DOCX documents and exporting them to friendly formats. Today it focuses on Markdown (rich and plain), with full handling of:

- Tracked changes (accept/reject/inline views)
- Lists with proper markers and indentation
- Tables (including merged cells / spans)
- Comments and threaded replies
- Headers and footers
- Images/drawings
- Bookmarks, hyperlinks, fields, and more

## Support overview

**Output formats**
- Markdown (rich + plain)
- HTML (rich + plain)
- Plain text

**Document features**
- Tracked changes: accept/reject/inline/split modes
- Comments: threads + replies
- Lists: markers + indentation
- Tables: cell spanning (row/col), borders/background/vertical-align (incl. `tblStyle` presets; theme colors/advanced border patterns are limited)
- Headers/footers, bookmarks, hyperlinks, fields
- Images/drawings (best-effort)

## Why this exists

Most DOCX “save as text” pipelines lose important fidelity: strikethroughs and deletions disappear, list markers collapse to bullets or vanish (pandoc), comments/track changes are missing, images are dropped, and tools like LibreOffice/Interop require UI or platform-specific installs. Docxport.Net walks the OOXML directly, headlessly, and emits Markdown/HTML/plain text while preserving tracked changes, comments, list markers, images, headers/footers, and other semantics.

## Quick start: DOCX → Markdown

### Command line

```bash
# Option A: Install the CLI as a .NET tool (recommended)
dotnet tool install -g DocxportNet.Cli
docxport my-doc.docx -o my-doc.md --tracked=accept

# Option B: Download a self-contained binary from GitHub Releases
# (pick your OS/arch zip, extract it, then run):
./docxport my-doc.docx -o my-doc.md --tracked=accept

# From source (development):
git clone https://github.com/gaspardpetit/Docxport.Net.git
dotnet run --project DocxportNet.Cli -- my-doc.docx -o my-doc.md --tracked=accept
```

### NuGet + code

Install: `dotnet add package DocxportNet`

```csharp
using DocxportNet;
using DocxportNet.Visitors.Markdown;

string docxPath = "my-doc.docx";
var visitor = new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.RICH);
string markdown = DxpExport.ExportToString(docxPath, visitor);

File.WriteAllText(Path.ChangeExtension(docxPath, ".md"), markdown);
```

## Tracked changes

Visitors can emit different views of tracked changes:

- Accept changes (default): `DxpTrackedChangeMode.AcceptChanges`
- Reject changes: `DxpTrackedChangeMode.RejectChanges`
- Inline markup (insert/delete): `DxpTrackedChangeMode.InlineChanges`
- Split accept/reject panes: `DxpTrackedChangeMode.SplitChanges`

Pick the mode on the visitor config, e.g.:

```csharp
var config = DxpMarkdownVisitorConfig.RICH with { TrackedChangeMode = DxpTrackedChangeMode.RejectChanges };
var rejectVisitor = new DxpMarkdownVisitor(config);
string rejected = DxpExport.ExportToString(docxPath, rejectVisitor);
```

## Visitors and options

**Markdown**  
- Presets: `DxpMarkdownVisitorConfig.CreateRichConfig()` (styled) and `CreatePlainConfig()` (minimal).  
- Options cover images, inline styling, rich tables, comments formatting, custom properties, and tracked change mode.

**HTML**  
- Preset: `DxpHtmlVisitorConfig.CreateRichConfig()` (styled) and `CreatePlainConfig()` (minimal).
- Options cover inline styles, colors/backgrounds, table borders, document colors, headers/footers, comments mode, custom properties, timeline, and tracked change mode.

**Plain text**  
- Presets: `DxpPlainTextVisitorConfig.CreateAcceptConfig()` and `CreateRejectConfig()` (choose tracked change handling).  
- Focused on readable text output with list markers, comments, and basic structure.

`DxpExport` has overloads for DOCX file paths, in-memory bytes, or an already-open `WordprocessingDocument`, and can return a `string`, a `byte[]`, write straight to a file path, or just drive a visitor that collects data.

### CLI

A ready-to-use console app lives in `DocxportNet.Cli` and is published as a .NET tool. Example:

```bash
dotnet tool install -g DocxportNet.Cli
docxport my.docx -o my.md --tracked=accept
```

Self-contained binaries for Windows/Linux/macOS are also published on GitHub Releases.

```bash
dotnet run --project DocxportNet.Cli -- my.docx -o my.md --tracked=accept
# or, using a self-contained binary:
./docxport my.docx --format=html --tracked=inline
```

Options: `--format=markdown|html|text`, `--tracked=accept|reject|inline|split` (text uses accept/reject), `--plain` (plain markdown), `-o, --output=path` (infers format from extension when `--format` is omitted), `--vars=path` (JSON/INI DOCVARIABLEs), `-D name=value` (repeatable overrides).

## Custom visitors

You can write your own `DxpIVisitor` to extract specific content. Example: collect all comments.

```csharp
using DocxportNet.API;
using DocxportNet.Visitors;
using DocumentFormat.OpenXml.Wordprocessing;

public sealed class CommentCollector : DxpVisitor
{
	public List<(string Author, string Text)> Comments { get; } = new();

	public override void VisitComment(Comment c, DxpIDocumentContext d)
	{
		Comments.Add((c.Author?.Value ?? "Unknown", c.InnerText));
	}
}

var collector = new CommentCollector();
DxpExport.Export("my-doc.docx", collector);
foreach (var (author, text) in collector.Comments)
	Console.WriteLine($"{author}: {text}");
```

It also ships a small utility that translates legacy “symbol fonts” (Symbol, Zapf Dingbats, Webdings, Wingdings) into modern Unicode.

## Field evaluation

Docxport.Net ships a standalone field evaluator for Word field codes. It can be used on its own or alongside the walker/exporter. For the exact language and field specifications, see [docs/word-fields.md](docs/word-fields.md).

High‑level support includes:
- Core built‑ins: `DATE`/`TIME`/`CREATEDATE`/`SAVEDATE`/`PRINTDATE`, `SET`/`REF`, `DOCPROPERTY`/`DOCVARIABLE`, `MERGEFIELD`, `SEQ`, `IF`/`COMPARE`/`NEXTIF`/`SKIPIF`, `ASK`
- Formulas: arithmetic, comparisons, functions (`SUM`/`AVERAGE`/etc.), nested fields, table cell references (via a resolver)
- Formatting switches: `\*` text transforms, `\#` numeric pictures, `\@` date/time pictures
- Locale‑aware formatting and list‑separator handling
- Number‑to‑words languages: English, French, German, Spanish, Italian, Portuguese, Danish, Japanese, Thai, Chinese (Simplified)

Minimal setup (standalone):

```csharp
using DocxportNet.Fields;

var eval = new DxpFieldEval();
eval.Context.SetBookmark("Total", "123.45");
eval.Context.SetDocumentPropertyValue("Title", new DxpFieldValue("My Doc"));

var date = await eval.EvalAsync(new DxpFieldInstruction("DATE \\@ \"yyyy-MM-dd\""));
var cond = await eval.EvalAsync(new DxpFieldInstruction("IF Total > 100 \"Big\" \"Small\""));
var formula = await eval.EvalAsync(new DxpFieldInstruction("= Total * 2 \\# \"$#,##0.00\""));
var title = await eval.EvalAsync(new DxpFieldInstruction("DOCPROPERTY Title \\* Upper"));

// date.Text    -> "2026-02-07" (uses default NowProvider)
// cond.Text    -> "Big"
// formula.Text -> "$246.90"
// title.Text   -> "MY DOC"
```

Minimal setup (with exporter + middleware):

```csharp
using DocxportNet;
using DocxportNet.Fields;
using DocxportNet.Visitors.PlainText;

var visitor = new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig());
// DxpExport wraps the visitor with field-eval middleware automatically.
string text = DxpExport.ExportToString("my-doc.docx", visitor);
```

Providing DOCVARIABLEs:

```csharp
using DocxportNet.Fields;

var eval = new DxpFieldEval();
eval.Context.SetDocVariable("ClientName", "Contoso Ltd.");

var result = await eval.EvalAsync(new DxpFieldInstruction("DOCVARIABLE ClientName \\* Upper"));
// result.Text -> "CONTOSO LTD."
```

If you need dynamic resolution (DB/HTTP/etc.), register a resolver on the context instead of pre‑seeding values.

## Gaps and contributions welcome

| Area | Current gap |
| --- | --- |
| List markers | Word supports exotic textual markers (“forty-two” in various languages); current formatter covers numeric/roman/alpha and supports text‑number formats via the shared number‑to‑words registry, but not every Word numbering style. |
| Shapes/SmartArt | Complex shapes/SmartArt/OLE rely on preview images if present; true vector or OLE rendering is not implemented. |
| Charts | Charts are emitted via available previews or placeholders; data-driven re-rendering is not implemented. |
| Math/Fields | Field evaluation exists, but complex math/field result rendering in exporters isn’t deeply covered by fixtures. |
| Tables (complex) | Table styles are partially supported (borders/background/vertical align, incl. `tblStyle` presets), but theme color resolution and advanced border patterns are still limited; nested/edge-case tables beyond supplied samples may need additional handling. |

Contributions that improve any of these areas are very welcome.

## Supported symbol fonts

Symbols typically used in list markers are automatically converted to their unicode equivalent. Supported fonts include:

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
