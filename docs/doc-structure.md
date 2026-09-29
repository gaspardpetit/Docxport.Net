# Binary DOC structure navigation

The structure-by-structure implementation inventory is in
[doc-format-coverage.md](doc-format-coverage.md).

`DocxportNet.Doc` is the first structural layer for legacy Word `.doc` files. It reads
the compound-file directory, the `WordDocument` file information block (FIB),
the selected `0Table` or `1Table` stream, and the FIB's location directory. It
reads `FibRgLw97` character counts to expose the global CP ranges for the main
document, footnotes, headers, comments, endnotes, and textboxes. It
recognizes the document settings, text piece table, style sheet, sections,
header/footer references, character and paragraph formatting, font table, and email
envelope locations. The walker enters Unicode version 8 envelope fields,
recipient collections, properties, attachments, and introduction text when the
visitor enters `MsoEnvelope`. The envelope header is read when its location is
entered; text and attachment bytes remain lazy.

Applications use the same export surface for DOC and DOCX input. `DxpExport`
recognizes binary DOC by its compound-file signature, builds the basic DOCX
projection internally, then invokes the requested DOCX visitor:

```csharp
using DocxportNet;
using DocxportNet.Visitors.PlainText;

var text = DxpExport.ExportToString("letter.doc",
    new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()));
var projectedDocx = DxpDocxExport.Export(File.ReadAllBytes("letter.doc"));
```

For the direct DOC-to-DOCX projection, without the subsequent DOCX walk, use
`DxpDocToDocx.Project`. It returns both the DOCX bytes and coverage details.
`DxpExport` uses this same projection whenever its input is a binary DOC:

```csharp
var projection = DxpDocToDocx.Project("letter.doc");
File.WriteAllBytes("letter.projected.docx", projection.DocxBytes);
var coverage = projection.Coverage;
```

The browser API accepts DOC bytes through `createDocxport().projectDocx` for
the direct projection, or through `export` and `resolveDocx` for the subsequent
DOCX walk. The low-level walker, index, and projector remain available
when callers need ranges, lazy payloads, or projection coverage details.

`DxpDocExport.Export` writes a plain-text binary DOC from DOCX or DOC input.
It uses the existing DOCX walker and `DxpDocVisitor`; DOC input first passes
through the basic projection. The CLI supports `--format=doc`, and browser
callers can use `createDocxport().exportDoc(bytes)`. The writer emits the main
story as Unicode text, paragraph marks, tabs, and line/page breaks. It does
not preserve formatting, tables, images, headers, footers, or other stories.
The writer constructs the compound file, Word 2002 FIB, fixed-index style
slots, section table, piece table, and formatting page references directly.
The read-side walker indexes these structures, including the stylesheet and
section boundaries. No DOC template is embedded.

`DocEditor` is a separate edit-and-save surface for existing DOC bytes. It
keeps the original indexed bytes unchanged while `SetEmailEnvelope`,
`SetEmailEnvelopeVisibility`, and `RemoveEmailEnvelope` queue edits. `Save`
flattens those edits into a new DOC byte array. An envelope replacement is
serialized into the selected table stream and its FIB location is updated;
the visibility bit is changed in the DOP. If no DOP exists, the editor creates
one for the visibility field. The editor preserves unrelated streams and
checks an old envelope's header and FIB range overlaps before clearing it.
Removing an envelope is not a secure sanitization operation.

```csharp
using DocxportNet.Doc;

using var editor = DocEditor.Open(docBytes);
editor.SetEmailEnvelope(new DocEmailEnvelope
{
    Subject = "Report",
    To = new[] { new DocEmailAddress("client@example.com", "Client") },
    Attachments = new[] { new DocEmailAttachment("report.pdf", pdfBytes) }
});
byte[] editedDoc = editor.Save();
```

`ReadEmailEnvelope` materializes the supported Unicode version 8 settings on
request. Visibility-only edits leave unknown envelope payloads untouched.
Browser callers can pass DOC bytes and an edit list to
`createDocxport().editDocEnvelope(bytes, edits)`; attachment content is passed
as `Uint8Array` or `ArrayBuffer`.

```csharp
using DocxportNet.Doc;

var walker = new DocStructureWalker();
using var doc = walker.Accept("letter.doc", new DocStructurePrintVisitor(Console.Out));
```

For nested XML output, use `DocStructureXmlVisitor` and dispose it after walking:

```csharp
using var output = File.CreateText("letter.structure.xml");
using var xml = new DocStructureXmlVisitor(output);
using var doc = new DocStructureWalker().Accept("letter.doc", xml);
```

For a reusable logical text index, use `DocTextIndexWalker`. It traverses the
piece table and records `Pcd` node references, CP ranges, WordDocument byte
ranges, and encoding without decoding text. On demand, it also indexes
stylesheet entries, sections, and character/paragraph formatting-page references:

```csharp
using var index = new DocTextIndexWalker().Index("letter.doc");
foreach (var span in index.GetPartSpans("Main"))
    Console.Write(span.Text); // Decodes the underlying piece on first access.
foreach (var page in index.FormattingPages)
    Console.WriteLine($"{page.Node.Kind}: FC {page.FcStart}..{page.FcEnd}");
```

The index owns the open `DocStructure` and exposes it through `Structure` for
future projections. `Pieces` are in logical CP order, even if their byte ranges
are elsewhere in `WordDocument`. `GetPartSpans` clips pieces to the selected
document part, so a piece crossing a part boundary contributes only the
appropriate text. Decoded piece text is cached on its underlying node. Dispose
the index after its spans are no longer needed; a caller-supplied input stream
remains open.

Accessing `FormattingPages`, `Styles`, or `Sections` builds only that view of
the index. `FormattingPages` records the FC range and page number without reading the
512-byte FKP page. Accessing a page's `Runs` parses its FC boundaries and
locates `Chpx` or `PapxInFkp` property blocks; their modifier bytes are still
uninterpreted. `Styles` exposes length-prefixed `LPStd` ranges, while `Sections`
exposes `Sed` ranges. The low-level structure walker emits the same nodes when
its visitor enters their parents. This supplies the read-side structure needed
by the current plain-text DOC writer, while logical paragraph/run assembly
from CP and FC ranges remains future work.

`DocToDocxProjector` turns an index into a minimal in-memory DOCX package:

```csharp
using var index = new DocTextIndexWalker().Index("letter.doc");
var projection = new DocToDocxProjector().Project(index);
File.WriteAllBytes("letter.projected.docx", projection.DocxBytes);
```

The first projection emits the main story as paragraphs, runs, tabs, and line
breaks. It keeps paragraph state across text pieces, even when their physical
byte ranges are out of order. Cell marks become tabs and section marks become
paragraph boundaries. XML-invalid control characters, including field markers,
are omitted. `Coverage` lists the projected main CP range, deferred document
parts, FIB locations, and compound streams, plus counts of omitted and
approximated characters. The DOCX can be passed to `DxpExport.ExportToString`
with `DxpPlainTextVisitor` for the current rough TXT export. The projection
does not yet reconstruct fields, tables, formatting, revisions, or notes.

XML elements use each node's kind as the tag name, with `name`, stream, offset,
and length attributes. Compound-file names are escaped as `\uXXXX` where needed,
because control characters in real stream names are not legal in XML text.
Each `Pcd` has a `Text` child containing its decoded characters in escaped
diagnostic form. The XML visitor requests that text as it visits each piece;
other visitors can leave pieces unmaterialized.

The walker calls `IDocStructureVisitor.Enter` in depth-first order. It disposes
the returned scope after visiting the node's children, including when a child
throws. Returning `null` skips the children. This follows the scope pattern of
the DOCX walker. A node has a kind, name, optional stream-relative offset and length, attributes, and child
nodes. `DocStructure.FindLocation("EmailEnvelope")` returns a typed FIB location
including the byte position of its offset/length fields in `WordDocument`.

Navigation reads the FIB fields and, when the visitor enters `TextPieceTable`,
the CLX piece-table metadata. It emits `Prc`, `Pcdt`, `PlcPcd`, and `Pcd` nodes.
Each `Pcd` has a logical CP range plus the referenced byte range and encoding in
`WordDocument`. The walker labels a piece with its document part, or adds
`PartSpan` children when a piece crosses part boundaries. It validates the final
CP against the FIB counts, including the extra terminal paragraph mark when
other document parts are present. Skipping `TextPieceTable` leaves its CLX
content unread.
Entering `Sections` reads the `PlcfSed` index and emits one `Sed` per section,
with its main-document CP range. When a descriptor references section properties,
it has a `Sepx` child giving the byte range in `WordDocument`; the property
bytes and individual formatting instructions are not parsed. Skipping `Sections`
leaves that index unread. Section CPs are positions in the logical text, while
`Sed` and `Sepx` offsets refer to physical streams.
Text bytes are decoded only when `Payload` is accessed on a `Pcd`, yielding
`DocTextPieceContent` with its CP range and string. The returned nodes are
location descriptors, not byte arrays. Parsable nodes expose `HasPayload`,
`IsPayloadLoaded`, and `Payload`; results are cached. Other current results are
`DocDopVisibility`, `DocEnvelopeHeader`, `DocEnvelopeText`, and
`DocEnvelopeBytes`. The XML visitor emits decoded envelope strings but leaves
attachment data unmaterialized. Returning `null` at `EmailEnvelope` or
`MsoEnvelope` skips its nested parsing. Unsupported recipient property types
are exposed as an opaque remainder. Dispose the returned `DocStructure` when finished; it holds
the compound file open so lazy materialization can read the original range.
For a caller-supplied stream, disposing `DocStructure` leaves that stream open.

This layer discovers locations and decodes raw document text on request. It does
not yet interpret paragraph, field, or table control characters or styles.
Envelope edits use the separate `DocEditor` surface described above. FIB entries without a recognized semantic
name are emitted as `FibEntryN`, so later parsers can be added without changing
the container/FIB navigation model. The reader rejects invalid compound files,
unsupported pre-Word 97 files, encrypted files, missing selected table streams,
and out-of-range locations needed for the envelope path.

Format references: [MS-CFB](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/50708a61-81d9-49c8-ab9c-43c98a795242),
[MS-DOC FIB](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/9aeaa2e7-4a45-468e-ab13-3f6193eb9394),
[MS-DOC FibRgLw97](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/37713d3c-a0c8-40f5-821f-bc9622c7de48),
[MS-DOC Document Parts](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/5f0c4329-8718-4d67-8cc7-60d8968c5127),
[MS-DOC FibRgFcLcb2000](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/265bca68-c4ef-4a03-8517-61d7e79850eb),
[MS-DOC Dop2000](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/fb0ce92e-36c9-4060-a37e-45708860d997),
[MS-DOC Clx](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/bad26767-b575-44d3-9da3-96378d56ce14),
[MS-DOC PlcPcd](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/1caae71f-35c4-49d7-adf0-af5fc766331c),
[MS-DOC PlcfSed](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/68a959a6-11e0-4f3e-9a99-76ca8cc4dddc),
[MS-DOC Sed](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/ae1ec5fc-e3c0-4e27-956e-9ceedc41cc2a),
[MS-DOC Sepx](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/1cc8d6e1-17e2-4667-99b6-39b2c70a0ebe),
and [MS-OSHARED MsoEnvelope](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/52f15393-0e4f-4bd9-a521-dcb8f0a869ef).
