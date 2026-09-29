# Binary DOC format coverage

## Coverage score

Scores use the **423 named-structure rows** below. Each Yes or Full counts as
1, Partial as 0.5, and No as 0; the score is the sum divided by 423. The
area-level and referenced-format tables are excluded from this denominator.

| Dimension | Yes / Full | Partial | No | Weighted score | Coverage |
| --- | ---: | ---: | ---: | ---: | ---: |
| Navigation | 37 | 0 | 386 | 37 / 423 | **8.75%** |
| Parsing | 0 | 34 | 389 | 17 / 423 | **4.02%** |

This is a structural inventory of **[MS-DOC] v20250819** (580 pages). It covers
every numbered structure in sections 2.5, 2.7, 2.8, and 2.9, plus STTB in
2.2.4. The other normative section 2 headings are tracked by area below. It is
a checklist of the current `DocxportNet.Doc` implementation, not a claim that
every structure is present in every `.doc` file. Sources:
[the surveyed 2025 specification](https://officeprotocoldocs-f5hpbjgea6b8gneq.b02.azurefd.net/files/MS-DOC/%5bMS-DOC%5d-250819.docx)
and the [current published edition](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/ccd7b486-7881-484c-a137-51170af7cc22).
The 2026-02-17 edition has the same 486 section 2 headings as the surveyed
2025 edition; the inventory remains pinned to the 2025 page numbers.

**Navigation = Yes** means the walker exposes a named compound-file entry,
stream-relative range, or logical CP range for that structure, possibly through
a named FIB location.
**Parsing = Partial** means at least one field is interpreted, but other fields,
child records, or referenced semantics remain uninterpreted. **Full** would mean
all specified fields of that structure are interpreted; this inventory makes no
full-coverage claim yet. **No** means there is no dedicated navigation or parser
for that named structure. A parent FIB range can still contain an unlisted child.
These statuses concern binary `.doc` reading; the DOCX projection has a separate
scope below.

Of the 423 named rows below, 34 have partial parsing, 3 have a known range but
no parser, and 386 have neither dedicated navigation nor parsing. These are
structure counts, not a measure of how much document content can be exported.

The inventory is based on the code in
[DocStructureWalker.cs](../DocxportNet/Doc/DocStructureWalker.cs),
[DocClxNavigator.cs](../DocxportNet/Doc/DocClxNavigator.cs),
[DocSectionNavigator.cs](../DocxportNet/Doc/DocSectionNavigator.cs), and
[DocFormattingNavigator.cs](../DocxportNet/Doc/DocFormattingNavigator.cs),
[DocStyleSheetNavigator.cs](../DocxportNet/Doc/DocStyleSheetNavigator.cs), and
[DocTextIndexWalker.cs](../DocxportNet/Doc/DocTextIndexWalker.cs).

## Document content and projection

| Area | Navigation | Parsing | Current limit |
| --- | --- | --- | --- |
| Compound-file streams | Yes | Partial | OpenMcdf reads the container; this library lists streams but interprets only WordDocument and the selected table stream. |
| Main document CP range | Yes | Partial | Text is decoded in logical piece order; paragraph marks and basic controls are interpreted only by the DOCX projector. |
| Footnotes, headers, comments, endnotes, textboxes | Yes | Partial | CP ranges and raw text can be requested; their own structures and DOCX parts are not projected. |
| Paragraphs and runs | Partial | Partial | FKP pages expose physical character and paragraph FC ranges. Logical paragraphs and runs are not yet assembled from those ranges and the text piece table. |
| Stylesheet and formatting pages | Yes | Partial | Style headers, style byte ranges, BTE entries, FKP run boundaries, and property byte ranges are indexed. Style and property modifiers remain uninterpreted. |
| Tables, fields, images, lists | No | No | Some FIB ranges are located; their semantic structures are not parsed. |
| DOCX projection | — | Partial | Projects main text, paragraph marks, tabs, and line breaks. Reports deferred parts, locations, streams, and approximated or omitted characters. |
| Binary `.doc` writing | — | Partial | A plain-text writer builds a Word 2002 FIB, fixed-index stylesheet slots, one-section PlcfSed, CLX, BTE indexes, and FKP pages without a template. It does not preserve source formatting or other stories. |
| Binary `.doc` editing | — | Partial | `DocEditor` indexes source bytes, queues envelope replacement/visibility/removal, and saves a new compound file while preserving unrelated streams. General structure editing is not available. |

The five single-property-modifier families in section 2.6 are tracked below.
`Pcd.prm` and `Sepx.grpprl` remain raw references or byte ranges. Individual
Sprm values inside those families are not enumerated as separate rows.

The DOCX projection code is in
[DocToDocxProjector.cs](../DocxportNet/Doc/DocToDocxProjector.cs).
Format dependencies such as [MS-CFB], [MS-ODRAW], [MS-OSHARED], [MS-OLEPS],
and [MS-OVBA] are outside this MS-DOC table; they need separate inventories if
we later implement their contents. The envelope navigator covers the supported
Unicode body as described below; the DOP reader exposes one visibility flag.

| Referenced structure | Navigation | Parsing | Current limit |
| --- | --- | --- | --- |
| MS-CFB container | Yes | Partial | OpenMcdf handles storage and stream access; this library emits their names and lengths. |
| MS-OSHARED email envelope | Yes | Partial | Walks Unicode version 8 scalar fields, string ranges, recipient collections/properties, and attachment ranges. String and attachment payloads are lazy; some recipient property types and version 6 bodies remain opaque. |
| MS-ODRAW drawings | No | No | Embedded drawing structures are not located individually. |
| MS-OLEPS property streams | Yes | No | Compound-file streams are listed; property sets are not decoded. |
| MS-OVBA macros | Yes | No | Compound-file storages are listed; VBA contents are not decoded. |

## Other normative section 2 areas

These entries complete the numbered section 2 survey outside the named binary
structures below. Sections 2.4 and 2.6 describe algorithms or modifier
families, so their rows are feature coverage rather than individual records.

| Spec section | Area | Navigation | Parsing | Details |
| --- | --- | --- | --- | --- |
| 2.1.1 | WordDocument stream | Yes | Partial | Opens stream, reads FIB and referenced text; other contents remain uninterpreted. |
| 2.1.2 | 0Table / 1Table stream | Yes | Partial | Selects stream from FIB; reads CLX and PlcfSed metadata. |
| 2.1.3 | Data stream | Yes | No | Listed as a compound-file stream when present. |
| 2.1.4 | ObjectPool storage | Yes | No | Storage tree is listed when present. |
| 2.1.4.1 | ObjInfo stream | Yes | No | Listed by path when present; contents are not parsed. |
| 2.1.4.2 | Print stream | Yes | No | Listed by path when present; contents are not parsed. |
| 2.1.4.3 | EPrint stream | Yes | No | Listed by path when present; contents are not parsed. |
| 2.1.5 | Custom XML Data storage | Yes | No | Storage tree is listed when present. |
| 2.1.6 | Summary Information stream | Yes | No | Listed by path; property sets are not parsed. |
| 2.1.7 | Document Summary Information stream | Yes | No | Listed by path; property sets are not parsed. |
| 2.1.8 | Encryption stream | Yes | No | Entry may be listed, but encrypted DOC files are rejected before walking. |
| 2.1.9 | Macros storage | Yes | No | Storage tree is listed when present. |
| 2.1.10 | XML Signatures storage | Yes | No | Storage tree is listed when present. |
| 2.1.11 | Signatures stream | Yes | No | Listed by path when present. |
| 2.1.12 | IRM Data Space storage | Yes | No | Storage tree is listed when present. |
| 2.1.13 | Protected Content stream | Yes | No | Listed by path when present. |
| 2.2.1 | Character Position (CP) | Yes | Partial | Indexes CP ranges; does not resolve all anchored objects and formatting. |
| 2.2.2 | PLC | Yes | Partial | Reads PlcPcd and PlcfSed; no generic PLC parser. |
| 2.2.3 | Valid Selection | No | No | Selection validation is not implemented. |
| 2.2.5 | Property Storage | No | No | Property instructions are not applied. |
| 2.2.5.1 | Sprm | No | No | Modifier operands are not decoded. |
| 2.2.5.2 | Prl | No | No | Property records are not decoded. |
| 2.2.6.1–3 | XOR, RC4, RC4 CryptoAPI encryption | No | No | Password-protected DOC files are rejected. |
| 2.3.1 | Main Document | Yes | Partial | Main CP range and text are indexed; basic paragraphs are projected heuristically. |
| 2.3.2–7 | Footnotes, headers, comments, endnotes, textboxes | Yes | Partial | Part CP ranges and raw text are available; individual stories are not parsed. |
| 2.4.1 | Retrieving Text | Yes | Partial | Piece offsets and encodings are decoded on demand; other character semantics remain raw. |
| 2.4.2 | Determining Paragraph Boundaries | No | No | FKP-based boundary algorithm is not implemented; projection splits on text marks. |
| 2.4.3–5 | Tables, cell boundaries, row boundaries | No | No | Table property indexes and FKPs are not parsed. |
| 2.4.6.1–6 | Applying paragraph, character, list, and style properties | No | No | Property modifier application is not implemented. |
| 2.4.7 | VtHyperlink application data | No | No | Hyperlink application data is not parsed. |
| 2.5.14 | Determining the nFib | Yes | Partial | Reads the base version and rejects pre-Word 97 files; does not implement the full version algorithm. |
| 2.5.15 | How to read the FIB | Yes | Partial | Reads enough FIB groups and pairs for current navigation; later-version groups remain uninterpreted. |
| 2.6.1–5 | Character, paragraph, table, section, picture modifiers | No | No | No modifier family is parsed or applied. |

## Named structures

The source page is included to make each row easy to find in the supplied
v20250819 specification. Unmarked rows are unimplemented, not verified absent
from sample files.

### Common structure

| Spec section | Structure | Navigation | Parsing | Details |
| --- | --- | --- | --- | --- |
| 2.2.4 | `STTB` | No | No | — (p. 33) |

### File information block

| Spec section | Structure | Navigation | Parsing | Details |
| --- | --- | --- | --- | --- |
| 2.5.1 | `Fib` | Yes | Partial | FIB boundaries and counts are read; later-version groups are not modeled. (p. 55) |
| 2.5.2 | `FibBase` | Yes | Partial | Reads identifier, version, flags, encryption bits, selected table stream, and fcMin/fcMac text bounds. (p. 57) |
| 2.5.3 | `FibRgW97` | No | No | — (p. 59) |
| 2.5.4 | `FibRgLw97` | Yes | Partial | Reads document-part character counts and cbMac; other fields remain raw. (p. 60) |
| 2.5.5 | `FibRgFcLcb` | Yes | Partial | Reads offset/length pairs and records their FIB field positions. (p. 62) |
| 2.5.6 | `FibRgFcLcb97` | Yes | Partial | Selected pairs have semantic names; most entries remain generic FibEntryN. (p. 62) |
| 2.5.7 | `FibRgFcLcb2000` | Yes | Partial | Generic offset/length pairs are read; version-specific fields are not named. (p. 82) |
| 2.5.8 | `FibRgFcLcb2002` | Yes | Partial | Generic offset/length pairs are read; version-specific fields are not named. (p. 85) |
| 2.5.9 | `FibRgFcLcb2003` | Yes | Partial | Generic offset/length pairs are read; version-specific fields are not named. (p. 92) |
| 2.5.10 | `FibRgFcLcb2007` | Yes | Partial | Generic offset/length pairs are read; version-specific fields are not named. (p. 99) |
| 2.5.11 | `FibRgCswNew` | No | No | — (p. 102) |
| 2.5.12 | `FibRgCswNewData2000` | No | No | — (p. 103) |
| 2.5.13 | `FibRgCswNewData2007` | No | No | — (p. 103) |

### Document properties

| Spec section | Structure | Navigation | Parsing | Details |
| --- | --- | --- | --- | --- |
| 2.7.1 | `Dop` | Yes | Partial | FIB locates the DOP; only one visibility flag is interpreted. (p. 148) |
| 2.7.2 | `DopBase` | No | No | — (p. 149) |
| 2.7.3 | `Dop95` | No | No | — (p. 155) |
| 2.7.4 | `Dop97` | No | No | — (p. 156) |
| 2.7.5 | `Dop2000` | Yes | Partial | Reads only the envelope-visibility byte at DOP offset + 504. (p. 159) |
| 2.7.6 | `Dop2002` | No | No | — (p. 163) |
| 2.7.7 | `Dop2003` | No | No | — (p. 166) |
| 2.7.8 | `Dop2007` | No | No | — (p. 168) |
| 2.7.9 | `Dop2010` | No | No | — (p. 170) |
| 2.7.10 | `Dop2013` | No | No | — (p. 170) |
| 2.7.11 | `Copts60` | No | No | — (p. 171) |
| 2.7.12 | `Copts80` | No | No | — (p. 172) |
| 2.7.13 | `Copts` | No | No | — (p. 173) |
| 2.7.14 | `Asumyi` | No | No | — (p. 176) |
| 2.7.15 | `Dogrid` | No | No | — (p. 177) |
| 2.7.16 | `DopTypography` | No | No | — (p. 178) |
| 2.7.17 | `DopMth` | No | No | — (p. 180) |

### PLC structures

| Spec section | Structure | Navigation | Parsing | Details |
| --- | --- | --- | --- | --- |
| 2.8.1 | `Plcbkf` | No | No | — (p. 182) |
| 2.8.2 | `Plcbkfd` | No | No | — (p. 183) |
| 2.8.3 | `Plcbkl` | No | No | — (p. 184) |
| 2.8.4 | `Plcbkld` | No | No | — (p. 184) |
| 2.8.5 | `PlcBteChpx` | Yes | Partial | Reads FC ranges and referenced character-formatting page numbers. (p. 185) |
| 2.8.6 | `PlcBtePapx` | Yes | Partial | Reads FC ranges and referenced paragraph-formatting page numbers. (p. 185) |
| 2.8.7 | `PlcfandRef` | No | No | — (p. 186) |
| 2.8.8 | `PlcfandTxt` | No | No | — (p. 186) |
| 2.8.9 | `PlcfAsumy` | No | No | — (p. 187) |
| 2.8.10 | `Plcfbkf` | No | No | — (p. 187) |
| 2.8.11 | `Plcfbkfd` | No | No | — (p. 188) |
| 2.8.12 | `Plcfbkl` | No | No | — (p. 189) |
| 2.8.13 | `Plcfbkld` | No | No | — (p. 189) |
| 2.8.14 | `Plcfcookie` | No | No | — (p. 189) |
| 2.8.15 | `PlcfcookieOld` | No | No | — (p. 190) |
| 2.8.16 | `PlcfendRef` | No | No | — (p. 190) |
| 2.8.17 | `PlcfendTxt` | No | No | — (p. 191) |
| 2.8.18 | `Plcffactoid` | No | No | — (p. 191) |
| 2.8.19 | `PlcffndRef` | No | No | — (p. 192) |
| 2.8.20 | `PlcffndTxt` | No | No | — (p. 192) |
| 2.8.21 | `Plcfgram` | No | No | — (p. 193) |
| 2.8.22 | `Plcfhdd` | Yes | No | FIB exposes its byte range as HeadersAndFooters; no PLC entries are read. (p. 193) |
| 2.8.23 | `PlcfHdrtxbxTxt` | No | No | — (p. 194) |
| 2.8.24 | `Plcflad` | No | No | — (p. 194) |
| 2.8.25 | `Plcfld` | No | No | — (p. 195) |
| 2.8.26 | `PlcfSed` | Yes | Partial | Reads section CP boundaries and Sed records; does not apply section properties. (p. 196) |
| 2.8.27 | `PlcfSpa` | No | No | — (p. 196) |
| 2.8.28 | `Plcfspl` | No | No | — (p. 197) |
| 2.8.29 | `PlcfTch` | No | No | — (p. 197) |
| 2.8.30 | `PlcfTxbxBkd` | No | No | — (p. 198) |
| 2.8.31 | `PlcfTxbxHdrBkd` | No | No | — (p. 199) |
| 2.8.32 | `PlcftxbxTxt` | No | No | — (p. 199) |
| 2.8.33 | `Plcfuim` | No | No | — (p. 200) |
| 2.8.34 | `PlcfWKB` | No | No | — (p. 200) |
| 2.8.35 | `PlcPcd` | Yes | Partial | Reads CP boundaries and piece descriptors; does not interpret PRM formatting. (p. 201) |

### Basic types

| Spec section | Structure | Navigation | Parsing | Details |
| --- | --- | --- | --- | --- |
| 2.9.1 | `Acd` | No | No | — (p. 202) |
| 2.9.2 | `Afd` | No | No | — (p. 203) |
| 2.9.3 | `ASUMY` | No | No | — (p. 204) |
| 2.9.4 | `ATNBE` | No | No | — (p. 204) |
| 2.9.5 | `AtrdExtra` | No | No | — (p. 204) |
| 2.9.6 | `ATRDPost10` | No | No | — (p. 205) |
| 2.9.7 | `ATRDPre10` | No | No | — (p. 205) |
| 2.9.8 | `BKC` | No | No | — (p. 206) |
| 2.9.9 | `BKF` | No | No | — (p. 207) |
| 2.9.10 | `BKFD` | No | No | — (p. 207) |
| 2.9.11 | `BKL` | No | No | — (p. 208) |
| 2.9.12 | `BKLD` | No | No | — (p. 208) |
| 2.9.13 | `BlockSel` | No | No | — (p. 209) |
| 2.9.14 | `Bool16` | No | No | — (p. 209) |
| 2.9.15 | `Bool8` | No | No | — (p. 209) |
| 2.9.16 | `Brc` | No | No | — (p. 209) |
| 2.9.17 | `Brc80` | No | No | — (p. 210) |
| 2.9.18 | `Brc80MayBeNil` | No | No | — (p. 210) |
| 2.9.19 | `BrcCvOperand` | No | No | — (p. 210) |
| 2.9.20 | `BrcMayBeNil` | No | No | — (p. 211) |
| 2.9.21 | `BrcOperand` | No | No | — (p. 211) |
| 2.9.22 | `BrcType` | No | No | — (p. 211) |
| 2.9.23 | `BxPap` | Yes | Partial | Reads the PAPX offset; reserved paragraph-height bytes remain raw. (p. 218) |
| 2.9.24 | `CAPI` | No | No | — (p. 218) |
| 2.9.25 | `CDB` | No | No | — (p. 219) |
| 2.9.26 | `CellHideMarkOperand` | No | No | — (p. 220) |
| 2.9.27 | `CellRangeFitText` | No | No | — (p. 220) |
| 2.9.28 | `CellRangeNoWrap` | No | No | — (p. 220) |
| 2.9.29 | `CellRangeTextFlow` | No | No | — (p. 221) |
| 2.9.30 | `CellRangeVertAlign` | No | No | — (p. 221) |
| 2.9.31 | `CFitTextOperand` | No | No | — (p. 221) |
| 2.9.32 | `Chpx` | Yes | Partial | Locates length-prefixed direct-property bytes; modifiers remain raw. (p. 222) |
| 2.9.33 | `ChpxFkp` | Yes | Partial | Reads run FC boundaries and offsets to character-property blocks on demand. (p. 222) |
| 2.9.34 | `Cid` | No | No | — (p. 223) |
| 2.9.35 | `CidAllocated` | No | No | — (p. 223) |
| 2.9.36 | `CidFci` | No | No | — (p. 223) |
| 2.9.37 | `CidMacro` | No | No | — (p. 226) |
| 2.9.38 | `Clx` | Yes | Partial | Indexes Prc and Pcdt records; Prc property instructions remain raw. (p. 227) |
| 2.9.39 | `CMajorityOperand` | No | No | — (p. 227) |
| 2.9.40 | `Cmt` | No | No | — (p. 227) |
| 2.9.41 | `CNFOperand` | No | No | — (p. 228) |
| 2.9.42 | `CNS` | No | No | — (p. 228) |
| 2.9.43 | `COLORREF` | No | No | — (p. 229) |
| 2.9.44 | `COSL` | No | No | — (p. 229) |
| 2.9.45 | `CSSA` | No | No | — (p. 230) |
| 2.9.46 | `CSSAOperand` | No | No | — (p. 230) |
| 2.9.47 | `CSymbolOperand` | No | No | — (p. 231) |
| 2.9.48 | `CTB` | No | No | — (p. 231) |
| 2.9.49 | `CTBWRAPPER` | No | No | — (p. 233) |
| 2.9.50 | `Customization` | No | No | — (p. 233) |
| 2.9.51 | `DCS` | No | No | — (p. 234) |
| 2.9.52 | `DefTableShd80Operand` | No | No | — (p. 235) |
| 2.9.53 | `DefTableShdOperand` | No | No | — (p. 235) |
| 2.9.54 | `DispFldRmOperand` | No | No | — (p. 235) |
| 2.9.55 | `Dofr` | No | No | — (p. 236) |
| 2.9.56 | `DofrFsn` | No | No | — (p. 236) |
| 2.9.57 | `DofrFsnFnm` | No | No | — (p. 237) |
| 2.9.58 | `DofrFsnName` | No | No | — (p. 238) |
| 2.9.59 | `DofrFsnp` | No | No | — (p. 238) |
| 2.9.60 | `DofrFsnSpbd` | No | No | — (p. 238) |
| 2.9.61 | `Dofrh` | No | No | — (p. 239) |
| 2.9.62 | `DofrRglstsf` | No | No | — (p. 239) |
| 2.9.63 | `Dofrt` | No | No | — (p. 240) |
| 2.9.64 | `DPCID` | No | No | — (p. 240) |
| 2.9.65 | `DTTM` | No | No | — (p. 241) |
| 2.9.66 | `FACTOIDINFO` | No | No | — (p. 242) |
| 2.9.67 | `FactoidSpls` | No | No | — (p. 242) |
| 2.9.68 | `FarEastLayoutOperand` | No | No | — (p. 242) |
| 2.9.69 | `Fatl` | No | No | — (p. 243) |
| 2.9.70 | `FBKF` | No | No | — (p. 244) |
| 2.9.71 | `FBKFD` | No | No | — (p. 244) |
| 2.9.72 | `FBKLD` | No | No | — (p. 244) |
| 2.9.73 | `FcCompressed` | Yes | Partial | Decodes compressed flag and text offset through Pcd; no separate node. (p. 245) |
| 2.9.74 | `FCCT` | No | No | — (p. 246) |
| 2.9.75 | `Fci` | No | No | — (p. 247) |
| 2.9.76 | `FCKS` | No | No | — (p. 315) |
| 2.9.77 | `FCKSOLD` | No | No | — (p. 316) |
| 2.9.78 | `FFData` | No | No | — (p. 317) |
| 2.9.79 | `FFDataBits` | No | No | — (p. 319) |
| 2.9.80 | `FFID` | No | No | — (p. 320) |
| 2.9.81 | `FFM` | No | No | — (p. 321) |
| 2.9.82 | `FFN` | No | No | — (p. 321) |
| 2.9.83 | `FieldMapBase` | No | No | — (p. 323) |
| 2.9.84 | `FieldMapDataItem` | No | No | — (p. 323) |
| 2.9.85 | `FieldMapInfo` | No | No | — (p. 324) |
| 2.9.86 | `FieldMapTerminator` | No | No | — (p. 325) |
| 2.9.87 | `FilterDataItem` | No | No | — (p. 325) |
| 2.9.88 | `Fld` | No | No | — (p. 326) |
| 2.9.89 | `fldch` | No | No | — (p. 326) |
| 2.9.90 | `flt` | No | No | — (p. 326) |
| 2.9.91 | `FNFB` | No | No | — (p. 329) |
| 2.9.92 | `FNIF` | No | No | — (p. 330) |
| 2.9.93 | `FNPI` | No | No | — (p. 330) |
| 2.9.94 | `FOBJH` | No | No | — (p. 331) |
| 2.9.95 | `FrameTextFlowOperand` | No | No | — (p. 331) |
| 2.9.96 | `FSDAP` | No | No | — (p. 332) |
| 2.9.97 | `Fsnk` | No | No | — (p. 332) |
| 2.9.98 | `Fssd` | No | No | — (p. 332) |
| 2.9.99 | `FssUnits` | No | No | — (p. 333) |
| 2.9.100 | `FTO` | No | No | — (p. 333) |
| 2.9.101 | `Fts` | No | No | — (p. 333) |
| 2.9.102 | `FtsWWidth_Indent` | No | No | — (p. 334) |
| 2.9.103 | `FtsWWidth_Table` | No | No | — (p. 334) |
| 2.9.104 | `FtsWWidth_TablePart` | No | No | — (p. 335) |
| 2.9.105 | `FTXBXNonReusable` | No | No | — (p. 335) |
| 2.9.106 | `FTXBXS` | No | No | — (p. 336) |
| 2.9.107 | `FTXBXSReusable` | No | No | — (p. 337) |
| 2.9.108 | `GOSL` | No | No | — (p. 337) |
| 2.9.109 | `GrammarSpls` | No | No | — (p. 338) |
| 2.9.110 | `grffldEnd` | No | No | — (p. 338) |
| 2.9.111 | `grfhic` | No | No | — (p. 339) |
| 2.9.112 | `GRFSTD` | No | No | — (p. 340) |
| 2.9.113 | `GrLPUpxSw` | No | No | — (p. 341) |
| 2.9.114 | `GrpPrlAndIstd` | No | No | — (p. 341) |
| 2.9.115 | `HFD` | No | No | — (p. 341) |
| 2.9.116 | `HFDBits` | No | No | — (p. 342) |
| 2.9.117 | `Hplxsdr` | No | No | — (p. 342) |
| 2.9.118 | `HresiOperand` | No | No | — (p. 343) |
| 2.9.119 | `Ico` | No | No | — (p. 343) |
| 2.9.120 | `IDPCI` | No | No | — (p. 344) |
| 2.9.121 | `Ipat` | No | No | — (p. 345) |
| 2.9.122 | `IScrollType` | No | No | — (p. 349) |
| 2.9.123 | `ItcFirstLim` | No | No | — (p. 349) |
| 2.9.124 | `Kcm` | No | No | — (p. 349) |
| 2.9.125 | `Kme` | No | No | — (p. 350) |
| 2.9.126 | `Kt` | No | No | — (p. 350) |
| 2.9.127 | `Kul` | No | No | — (p. 351) |
| 2.9.128 | `LadSpls` | No | No | — (p. 351) |
| 2.9.129 | `LBCOperand` | No | No | — (p. 352) |
| 2.9.130 | `LEGOXTR_V11` | No | No | — (p. 352) |
| 2.9.131 | `LFO` | No | No | — (p. 353) |
| 2.9.132 | `LFOData` | No | No | — (p. 354) |
| 2.9.133 | `LFOLVL` | No | No | — (p. 354) |
| 2.9.134 | `LID` | No | No | — (p. 355) |
| 2.9.135 | `LPStd` | Yes | Partial | Reads each style's length and byte range; style bodies remain raw. (p. 355) |
| 2.9.136 | `LPStshi` | Yes | Partial | Reads stylesheet header size and range. (p. 355) |
| 2.9.137 | `LPStshiGrpPrl` | No | No | — (p. 355) |
| 2.9.138 | `LPUpxChpx` | No | No | — (p. 356) |
| 2.9.139 | `LPUpxChpxRM` | No | No | — (p. 356) |
| 2.9.140 | `LPUpxPapx` | No | No | — (p. 356) |
| 2.9.141 | `LPUpxPapxRM` | No | No | — (p. 357) |
| 2.9.142 | `LPUpxRm` | No | No | — (p. 357) |
| 2.9.143 | `LPUpxTapx` | No | No | — (p. 357) |
| 2.9.144 | `LPXCharBuffer9` | No | No | — (p. 358) |
| 2.9.145 | `LSD` | No | No | — (p. 358) |
| 2.9.146 | `LSPD` | No | No | — (p. 359) |
| 2.9.147 | `LSTF` | No | No | — (p. 359) |
| 2.9.148 | `Lstsf` | No | No | — (p. 360) |
| 2.9.149 | `LVL` | No | No | — (p. 360) |
| 2.9.150 | `LVLF` | No | No | — (p. 361) |
| 2.9.151 | `MacroName` | No | No | — (p. 363) |
| 2.9.152 | `MacroNames` | No | No | — (p. 364) |
| 2.9.153 | `MathPrOperand` | No | No | — (p. 364) |
| 2.9.154 | `Mcd` | No | No | — (p. 364) |
| 2.9.155 | `MDP` | No | No | — (p. 365) |
| 2.9.156 | `MFPF` | No | No | — (p. 365) |
| 2.9.157 | `NilBrc` | No | No | — (p. 366) |
| 2.9.158 | `NilPICFAndBinData` | No | No | — (p. 366) |
| 2.9.159 | `NumRM` | No | No | — (p. 367) |
| 2.9.160 | `NumRMOperand` | No | No | — (p. 369) |
| 2.9.161 | `OcxInfo` | No | No | — (p. 369) |
| 2.9.162 | `ODSOPropertyBase` | No | No | — (p. 370) |
| 2.9.163 | `ODSOPropertyLarge` | No | No | — (p. 372) |
| 2.9.164 | `ODSOPropertyStandard` | No | No | — (p. 372) |
| 2.9.165 | `ODT` | No | No | — (p. 372) |
| 2.9.166 | `ODTPersist1` | No | No | — (p. 373) |
| 2.9.167 | `ODTPersist2` | No | No | — (p. 374) |
| 2.9.168 | `OfficeArtClientAnchor` | No | No | — (p. 375) |
| 2.9.169 | `OfficeArtClientData` | No | No | — (p. 375) |
| 2.9.170 | `OfficeArtClientTextbox` | No | No | — (p. 375) |
| 2.9.171 | `OfficeArtContent` | No | No | — (p. 376) |
| 2.9.172 | `OfficeArtWordDrawing` | No | No | — (p. 376) |
| 2.9.173 | `PANOSE` | No | No | — (p. 377) |
| 2.9.174 | `PapxFkp` | Yes | Partial | Reads paragraph FC boundaries and BxPap references on demand. (p. 381) |
| 2.9.175 | `PapxInFkp` | Yes | Partial | Locates length-prefixed direct-property bytes; modifiers remain raw. (p. 382) |
| 2.9.176 | `PbiGrfOperand` | No | No | — (p. 382) |
| 2.9.177 | `Pcd` | Yes | Partial | Reads CP range, flags, file offset, encoding, and raw PRM; no PRM semantics. (p. 383) |
| 2.9.178 | `Pcdt` | Yes | Partial | Reads container length and locates its PlcPcd child. (p. 383) |
| 2.9.179 | `PChgTabsAdd` | No | No | — (p. 384) |
| 2.9.180 | `PChgTabsDel` | No | No | — (p. 384) |
| 2.9.181 | `PChgTabsDelClose` | No | No | — (p. 384) |
| 2.9.182 | `PChgTabsOperand` | No | No | — (p. 385) |
| 2.9.183 | `PChgTabsPapxOperand` | No | No | — (p. 386) |
| 2.9.184 | `PgbApplyTo` | No | No | — (p. 386) |
| 2.9.185 | `PgbOffsetFrom` | No | No | — (p. 386) |
| 2.9.186 | `PgbPageDepth` | No | No | — (p. 386) |
| 2.9.187 | `PGPArray` | No | No | — (p. 387) |
| 2.9.188 | `PGPInfo` | No | No | — (p. 387) |
| 2.9.189 | `PGPOptions` | No | No | — (p. 388) |
| 2.9.190 | `PICF` | No | No | — (p. 389) |
| 2.9.191 | `PICF_Shape` | No | No | — (p. 390) |
| 2.9.192 | `PICFAndOfficeArtData` | No | No | — (p. 390) |
| 2.9.193 | `PICMID` | No | No | — (p. 391) |
| 2.9.194 | `PlcfGlsy` | No | No | — (p. 393) |
| 2.9.195 | `PlfAcd` | No | No | — (p. 393) |
| 2.9.196 | `PlfCosl` | No | No | — (p. 393) |
| 2.9.197 | `PlfGosl` | No | No | — (p. 394) |
| 2.9.198 | `PlfguidUim` | No | No | — (p. 394) |
| 2.9.199 | `PlfKme` | No | No | — (p. 395) |
| 2.9.200 | `PlfLfo` | No | No | — (p. 395) |
| 2.9.201 | `PlfLst` | No | No | — (p. 395) |
| 2.9.202 | `PlfMcd` | No | No | — (p. 396) |
| 2.9.203 | `PLRSID` | No | No | — (p. 396) |
| 2.9.204 | `Pmfs` | No | No | — (p. 397) |
| 2.9.205 | `Pms` | No | No | — (p. 399) |
| 2.9.206 | `PnFkpChpx` | Yes | Partial | Resolves the 22-bit page number to a WordDocument byte range. (p. 401) |
| 2.9.207 | `PnFkpPapx` | Yes | Partial | Resolves the 22-bit page number to a WordDocument byte range. (p. 401) |
| 2.9.208 | `PositionCodeOperand` | No | No | — (p. 401) |
| 2.9.209 | `Prc` | Yes | Partial | Reads property-block length and range; does not parse its instructions. (p. 402) |
| 2.9.210 | `PrcData` | No | No | — (p. 402) |
| 2.9.211 | `PrDrvr` | No | No | — (p. 402) |
| 2.9.212 | `PrEnvLand` | No | No | — (p. 403) |
| 2.9.213 | `PrEnvPort` | No | No | — (p. 403) |
| 2.9.214 | `Prm` | No | No | — (p. 403) |
| 2.9.215 | `Prm0` | No | No | — (p. 403) |
| 2.9.216 | `Prm1` | No | No | — (p. 405) |
| 2.9.217 | `PropRMark` | No | No | — (p. 405) |
| 2.9.218 | `PropRMarkOperand` | No | No | — (p. 406) |
| 2.9.219 | `ProtectionType` | No | No | — (p. 406) |
| 2.9.220 | `PRTI` | No | No | — (p. 406) |
| 2.9.221 | `PTIstdInfoOperand` | No | No | — (p. 407) |
| 2.9.222 | `Rca` | No | No | — (p. 407) |
| 2.9.223 | `RecipientBase` | No | No | — (p. 408) |
| 2.9.224 | `RecipientDataItem` | No | No | — (p. 408) |
| 2.9.225 | `RecipientInfo` | No | No | — (p. 409) |
| 2.9.226 | `RecipientTerminator` | No | No | — (p. 410) |
| 2.9.227 | `Rfs` | No | No | — (p. 410) |
| 2.9.228 | `RgCdb` | No | No | — (p. 411) |
| 2.9.229 | `RgxOcxInfo` | No | No | — (p. 411) |
| 2.9.230 | `RmdThreading` | No | No | — (p. 412) |
| 2.9.231 | `Rnc` | No | No | — (p. 416) |
| 2.9.232 | `RouteSlip` | No | No | — (p. 417) |
| 2.9.233 | `RouteSlipInfo` | No | No | — (p. 418) |
| 2.9.234 | `RouteSlipProtectionEnum` | No | No | — (p. 419) |
| 2.9.235 | `SBkcOperand` | No | No | — (p. 419) |
| 2.9.236 | `SBOrientationOperand` | No | No | — (p. 419) |
| 2.9.237 | `SClmOperand` | No | No | — (p. 419) |
| 2.9.238 | `SDmBinOperand` | No | No | — (p. 420) |
| 2.9.239 | `SDTI` | No | No | — (p. 420) |
| 2.9.240 | `SDTT` | No | No | — (p. 421) |
| 2.9.241 | `SDxaColSpacingOperand` | No | No | — (p. 421) |
| 2.9.242 | `SDxaColWidthOperand` | No | No | — (p. 421) |
| 2.9.243 | `Sed` | Yes | Partial | Reads CP range and fcSepx; undefined fields are ignored. (p. 422) |
| 2.9.244 | `Selsf` | No | No | — (p. 422) |
| 2.9.245 | `Sepx` | Yes | Partial | Reads cb to bound the Sepx range; does not decode grpprl. (p. 424) |
| 2.9.246 | `SFpcOperand` | No | No | — (p. 425) |
| 2.9.247 | `Shd` | No | No | — (p. 425) |
| 2.9.248 | `Shd80` | No | No | — (p. 426) |
| 2.9.249 | `SHDOperand` | No | No | — (p. 427) |
| 2.9.250 | `SLncOperand` | No | No | — (p. 427) |
| 2.9.251 | `SmartTagData` | No | No | — (p. 427) |
| 2.9.252 | `SortColumnAndDirection` | No | No | — (p. 428) |
| 2.9.253 | `Spa` | No | No | — (p. 428) |
| 2.9.254 | `SpellingSpls` | No | No | — (p. 430) |
| 2.9.255 | `SPgbPropOperand` | No | No | — (p. 431) |
| 2.9.256 | `SPLS` | No | No | — (p. 431) |
| 2.9.257 | `SPPOperand` | No | No | — (p. 432) |
| 2.9.258 | `STD` | Yes | No | Style definition byte ranges are exposed but not parsed. (p. 432) |
| 2.9.259 | `Stdf` | No | No | — (p. 433) |
| 2.9.260 | `StdfBase` | No | No | — (p. 433) |
| 2.9.261 | `StdfPost2000` | No | No | — (p. 435) |
| 2.9.262 | `StdfPost2000OrNone` | No | No | — (p. 436) |
| 2.9.263 | `StkCharGRLPUPX` | No | No | — (p. 436) |
| 2.9.264 | `StkCharLPUpxGrLPUpxRM` | No | No | — (p. 437) |
| 2.9.265 | `StkCharUpxGrLPUpxRM` | No | No | — (p. 437) |
| 2.9.266 | `StkListGRLPUPX` | No | No | — (p. 437) |
| 2.9.267 | `StkParaGRLPUPX` | No | No | — (p. 438) |
| 2.9.268 | `StkParaLPUpxGrLPUpxRM` | No | No | — (p. 438) |
| 2.9.269 | `StkParaUpxGrLPUpxRM` | No | No | — (p. 439) |
| 2.9.270 | `StkTableGRLPUPX` | No | No | — (p. 439) |
| 2.9.271 | `STSH` | Yes | Partial | Reads the style count and locates all LPStd entries. (p. 440) |
| 2.9.272 | `STSHI` | Yes | Partial | Exposes the stylesheet information range and Stshif child. (p. 441) |
| 2.9.273 | `STSHIB` | No | No | — (p. 441) |
| 2.9.274 | `Stshif` | Yes | Partial | Reads cstd, cbSTDBaseInFile, flags, and fixed-index style count; other defaults remain raw. The plain-text writer serializes its minimal header. (p. 442) |
| 2.9.275 | `StshiLsd` | No | No | — (p. 443) |
| 2.9.276 | `SttbfAssoc` | No | No | — (p. 443) |
| 2.9.277 | `SttbfAtnBkmk` | No | No | — (p. 444) |
| 2.9.278 | `SttbfAutoCaption` | No | No | — (p. 445) |
| 2.9.279 | `SttbfBkmk` | No | No | — (p. 446) |
| 2.9.280 | `SttbfBkmkBPRepairs` | No | No | — (p. 450) |
| 2.9.281 | `SttbfBkmkFactoid` | No | No | — (p. 451) |
| 2.9.282 | `SttbfBkmkFcc` | No | No | — (p. 452) |
| 2.9.283 | `SttbfBkmkProt` | No | No | — (p. 453) |
| 2.9.284 | `SttbfBkmkSdt` | No | No | — (p. 454) |
| 2.9.285 | `SttbfCaption` | No | No | — (p. 455) |
| 2.9.286 | `SttbfFfn` | Yes | No | FIB exposes its FontTable byte range; font records are not read. (p. 456) |
| 2.9.287 | `SttbfGlsy` | No | No | — (p. 456) |
| 2.9.288 | `SttbFnm` | No | No | — (p. 457) |
| 2.9.289 | `SttbfRfs` | No | No | — (p. 458) |
| 2.9.290 | `SttbfRMark` | No | No | — (p. 459) |
| 2.9.291 | `SttbGlsyStyle` | No | No | — (p. 460) |
| 2.9.292 | `SttbListNames` | No | No | — (p. 461) |
| 2.9.293 | `SttbProtUser` | No | No | — (p. 462) |
| 2.9.294 | `SttbRgtplc` | No | No | — (p. 463) |
| 2.9.295 | `SttbSavedBy` | No | No | — (p. 464) |
| 2.9.296 | `SttbTtmbd` | No | No | — (p. 464) |
| 2.9.297 | `SttbW6` | No | No | — (p. 465) |
| 2.9.298 | `StwUser` | No | No | — (p. 465) |
| 2.9.299 | `Sty` | No | No | — (p. 467) |
| 2.9.300 | `TabJC` | No | No | — (p. 467) |
| 2.9.301 | `TabLC` | No | No | — (p. 467) |
| 2.9.302 | `TableBordersOperand` | No | No | — (p. 468) |
| 2.9.303 | `TableBordersOperand80` | No | No | — (p. 469) |
| 2.9.304 | `TableBrc80Operand` | No | No | — (p. 470) |
| 2.9.305 | `TableBrcOperand` | No | No | — (p. 470) |
| 2.9.306 | `TableCellWidthOperand` | No | No | — (p. 471) |
| 2.9.307 | `TableSel` | No | No | — (p. 471) |
| 2.9.308 | `TableShadeOperand` | No | No | — (p. 472) |
| 2.9.309 | `TBC` | No | No | — (p. 472) |
| 2.9.310 | `TBD` | No | No | — (p. 473) |
| 2.9.311 | `TBDelta` | No | No | — (p. 473) |
| 2.9.312 | `Tbkd` | No | No | — (p. 475) |
| 2.9.313 | `TC80` | No | No | — (p. 475) |
| 2.9.314 | `TCellBrcTypeOperand` | No | No | — (p. 476) |
| 2.9.315 | `Tcg` | No | No | — (p. 476) |
| 2.9.316 | `Tcg255` | No | No | — (p. 477) |
| 2.9.317 | `TCGRF` | No | No | — (p. 477) |
| 2.9.318 | `TcgSttbf` | No | No | — (p. 478) |
| 2.9.319 | `TcgSttbfCore` | No | No | — (p. 479) |
| 2.9.320 | `Tch` | No | No | — (p. 479) |
| 2.9.321 | `TDefTableOperand` | No | No | — (p. 480) |
| 2.9.322 | `TDxaColOperand` | No | No | — (p. 480) |
| 2.9.323 | `TextFlow` | No | No | — (p. 481) |
| 2.9.324 | `TInsertOperand` | No | No | — (p. 481) |
| 2.9.325 | `TIQ` | No | No | — (p. 481) |
| 2.9.326 | `TLP` | No | No | — (p. 482) |
| 2.9.327 | `ToggleOperand` | No | No | — (p. 482) |
| 2.9.328 | `Tplc` | No | No | — (p. 483) |
| 2.9.329 | `TplcBuildIn` | No | No | — (p. 483) |
| 2.9.330 | `TplcUser` | No | No | — (p. 484) |
| 2.9.331 | `Ttmbd` | No | No | — (p. 484) |
| 2.9.332 | `UFEL` | No | No | — (p. 485) |
| 2.9.333 | `UID` | No | No | — (p. 486) |
| 2.9.334 | `UidSel` | No | No | — (p. 486) |
| 2.9.335 | `UIM` | No | No | — (p. 487) |
| 2.9.336 | `UpxChpx` | No | No | — (p. 487) |
| 2.9.337 | `UPXPadding` | No | No | — (p. 488) |
| 2.9.338 | `UpxPapx` | No | No | — (p. 488) |
| 2.9.339 | `UpxRm` | No | No | — (p. 490) |
| 2.9.340 | `UpxTapx` | No | No | — (p. 490) |
| 2.9.341 | `VerticalAlign` | No | No | — (p. 492) |
| 2.9.342 | `VerticalMergeFlag` | No | No | — (p. 492) |
| 2.9.343 | `VertMergeOperand` | No | No | — (p. 492) |
| 2.9.344 | `Vjc` | No | No | — (p. 493) |
| 2.9.345 | `WHeightAbs` | No | No | — (p. 493) |
| 2.9.346 | `WKB` | No | No | — (p. 493) |
| 2.9.347 | `Wpms` | No | No | — (p. 494) |
| 2.9.348 | `Wpmsdt` | No | No | — (p. 495) |
| 2.9.349 | `XAS` | No | No | — (p. 495) |
| 2.9.350 | `XAS_nonNeg` | No | No | — (p. 495) |
| 2.9.351 | `XAS_plusOne` | No | No | — (p. 496) |
| 2.9.352 | `XSDR` | No | No | — (p. 496) |
| 2.9.353 | `Xst` | No | No | — (p. 496) |
| 2.9.354 | `Xstz` | No | No | — (p. 497) |
| 2.9.355 | `YAS` | No | No | — (p. 497) |
| 2.9.356 | `YAS_nonNeg` | No | No | — (p. 497) |
| 2.9.357 | `YAS_plusOne` | No | No | — (p. 497) |
