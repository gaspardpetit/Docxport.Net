using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.api;
using System.Xml.Linq;

namespace DocxportNet.walker;

public sealed class SectionLayout
{
	public PageSize? PageSize;
	public PageMargin? PageMargin;
	public Columns? Columns;
	public DocGrid? DocGrid;
	public PageBorders? PageBorders;
	public LineNumberType? LineNumbers;
	public TextDirection? TextDirection;            // w:textDirection (if used)
	public VerticalTextAlignment? VerticalJustification;
	public FootnoteProperties? FootnoteProperties;
	public EndnoteProperties? EndnoteProperties;
}


public class DxpWalker
{
	private MainDocumentPart? _main;
	private OpenXmlPart? _currentPart;

	private long? _CurrentFootnoteId = null; // null => not inside a note

	private struct FieldFrame
	{
		public bool SeenSeparate;
		public IDisposable? ResultScope;
	}

	private readonly Stack<FieldFrame> _fieldStack = new();



	private DxpComments _comments = new DxpComments();
	private DxpDrawings _drawings = new DxpDrawings();
	private DxpTables _tables = new DxpTables();
	private DxpLists _lists = new DxpLists();
	private DxpFootnotes _footnotes = new DxpFootnotes();
	private DocxEndnotes _endnotes = new DocxEndnotes();
	private readonly HashSet<string> _referencedAnchors = new HashSet<string>(StringComparer.Ordinal);

	private IDxpStyleTracker _style = new DxpStyleTracker();

	public void Accept(string docxPath, IDxpVisitor v)
	{
		using var doc = WordprocessingDocument.Open(docxPath, false);
		Accept(doc, v);
	}

	public sealed record CustomFileProperty(string Name, string? Type, object? Value);


	public void Accept(WordprocessingDocument doc, IDxpVisitor v)
	{
		_referencedAnchors.Clear();
		if (doc.MainDocumentPart == null)
			return;

		if (v is visitors.DxpMarkdownVisitor mdv)
			mdv.SetReferencedAnchors(_referencedAnchors);

		_main = doc.MainDocumentPart;
		_currentPart = _main;

		_lists.Init(doc);
		_footnotes.Init(doc.MainDocumentPart);
		_endnotes.Init(doc.MainDocumentPart);
		_comments.Init(doc.MainDocumentPart);

		IDxpStyleResolver s = new DxpStyleResolver(doc);

		// New: document-wide settings/background
		var settings = doc.MainDocumentPart.DocumentSettingsPart?.Settings;
		if (settings != null)
			v.VisitDocumentSettings(settings, s);

		var background = doc.MainDocumentPart.Document?.DocumentBackground;
		if (background != null)
			v.VisitDocumentBackground(background, s);

		var core = doc.PackageProperties;
		v.VisitCoreFileProperties(core);

		var customPart = doc.CustomFilePropertiesPart;
		var props = customPart?.Properties;

		if (props != null)
		{
			IEnumerable<CustomFileProperty> custom = props
				.Elements<CustomDocumentProperty>()
				.Select(p => new CustomFileProperty(
					p.Name?.Value ?? string.Empty,
					p.FirstChild?.LocalName,  // e.g., "vt:lpwstr", "vt:bool", "vt:filetime", etc.
					p.FirstChild?.InnerText   // the string form of the value
				));

			v.VisitCustomFileProperties(custom);
		}

		var body = doc.MainDocumentPart?.Document?.Body
			?? throw new InvalidOperationException("DOCX has no main document body.");

		// Walk the main story then section-anchored headers/footers (see #2)
		WalkBody(body, s, v);
		// Remove the global “walk all headers/footers” to avoid duplicates (see #2)

		// Footnotes/Endnotes
		foreach (var fn in _footnotes.GetFootnotes())
			WalkFootnote(fn.Item1, fn.Item2, fn.Item3, s, v);
		foreach (var en in _endnotes.GetEndnotes())
			WalkEndnote(en.Item1, en.Item2, en.Item3, s, v);

		VisitBibliographyIfPresent(doc, s, v);

		_currentPart = null;
		_main = null;
	}

	private void VisitBibliographyIfPresent(WordprocessingDocument doc, IDxpStyleResolver s, IDxpVisitor v)
	{
		// There is no MainDocumentPart.BibliographyPart property in the SDK.
			// Fetch the part via the container API.
			var bibPart = doc.MainDocumentPart?
				.GetPartsOfType<CustomXmlPart>()
				.FirstOrDefault(p =>
					string.Equals(p.RelationshipType,
						"http://schemas.openxmlformats.org/officeDocument/2006/relationships/bibliography",
						StringComparison.OrdinalIgnoreCase));

			if (bibPart != null)
			{
				XDocument sourcesXml;
				using (var strm = bibPart.GetStream(FileMode.Open, FileAccess.Read))
					sourcesXml = XDocument.Load(strm); // root is b:Sources (OOXML bibliography)

				// Visitor can accept the CustomXmlPart and/or XDocument for bibliography sources.
				v.VisitBibliographySources(bibPart, sourcesXml);
			}
		}

	private void WalkHeader(Header header, IDxpStyleResolver s, IDxpVisitor v)
	{
		// Header allows the same block-level content universe as Body.
		foreach (var child in header.ChildElements)
		{
			WalkBlock(child, s, v);
		}
		_style.ResetStyle(v);
	}

	private void WalkFooter(Footer footer, IDxpStyleResolver s, IDxpVisitor v)
	{
		// Footer allows the same block-level content universe as Body.
		foreach (var child in footer.ChildElements)
		{
			WalkBlock(child, s, v);
		}
		_style.ResetStyle(v);
	}


	private void WalkBody(Body body, IDxpStyleResolver s, IDxpVisitor v)
	{
		var lastSectPr = body.Descendants<SectionProperties>().LastOrDefault();
		var firstSectPr = body.Descendants<SectionProperties>().FirstOrDefault();

		var sections = BuildSections(body);

		if (firstSectPr != null && v is visitors.DxpMarkdownVisitor mdv)
		{
			mdv.SetDefaultSectionLayout(CreateSectionLayout(firstSectPr));
		}

		using (v.VisitBodyBegin(body, s))
		{
			if (sections.Count == 0)
			{
				foreach (var child in body.ChildElements)
				{
					WalkBlock(child, s, v, lastSectPr);
				}
				return;
			}

			foreach (var section in sections)
			{
				bool isLast = ReferenceEquals(section.Properties, lastSectPr);
				EmitSectionStart(section.Properties, s, v, isLast);

				foreach (var child in section.Blocks)
				{
					WalkBlock(child, s, v, lastSectPr);
				}

				EmitSectionEnd(section.Properties, s, v, isLast);
			}
		}
	}

	private sealed record SectionSlice(SectionProperties Properties, IReadOnlyList<OpenXmlElement> Blocks);

	private static SectionProperties? ExtractSectionProperties(OpenXmlElement block, out bool includeBlock)
	{
		includeBlock = true;

		if (block is SectionProperties sp)
		{
			includeBlock = false;
			return sp;
		}

		if (block is Paragraph p)
		{
			var pp = p.GetFirstChild<ParagraphProperties>();
			var paragraphSectPr = pp?.GetFirstChild<SectionProperties>();
			return paragraphSectPr;
		}

		return null;
	}

	private List<SectionSlice> BuildSections(Body body)
	{
		var sections = new List<SectionSlice>();
		var sectPrs = body.Descendants<SectionProperties>().ToList();
		if (sectPrs.Count == 0)
			return sections;

		int sectIndex = 0;
		var currentSectPr = sectPrs[sectIndex];
		var currentBlocks = new List<OpenXmlElement>();

		foreach (var child in body.ChildElements)
		{
			bool include = true;
			var sp = ExtractSectionProperties(child, out include);

			if (include)
				currentBlocks.Add(child);

			if (sp != null)
			{
				var props = currentSectPr ?? sp;
				sections.Add(new SectionSlice(props, currentBlocks.ToList()));
				currentBlocks.Clear();

				sectIndex++;
				currentSectPr = sectIndex < sectPrs.Count ? sectPrs[sectIndex] : null;
			}
		}

		if (currentBlocks.Count > 0 && currentSectPr != null)
			sections.Add(new SectionSlice(currentSectPr, currentBlocks));

		return sections;
	}

	private void EmitSectionStart(SectionProperties sp, IDxpStyleResolver s, IDxpVisitor v, bool isLastSection)
	{
		bool emitSectionContent = v is visitors.DxpMarkdownVisitor mdv
			? mdv.EmitSectionHeadersFooters
			: true;

		if (emitSectionContent)
		{
			RenderSectionHeaders(sp, s, v);
			if (v is visitors.DxpMarkdownVisitor mdv2)
				mdv2.BeginSectionBody();
		}
	}

	private void EmitSectionEnd(SectionProperties sp, IDxpStyleResolver s, IDxpVisitor v, bool isLastSection)
	{
		bool emitSectionContent = v is visitors.DxpMarkdownVisitor mdv
			? mdv.EmitSectionHeadersFooters
			: true;

		if (emitSectionContent)
		{
			if (v is visitors.DxpMarkdownVisitor mdv2)
				mdv2.EndSectionBody();
			RenderSectionFooters(sp, s, v);
		}

		if (!isLastSection)
		{
			v.VisitSectionProperties(sp, s);

			var layout = CreateSectionLayout(sp);
			v.VisitSectionLayout(sp, layout, s);
		}
	}

	private void WalkBlock(OpenXmlElement block, IDxpStyleResolver s, IDxpVisitor v, SectionProperties? lastSectPr = null)
	{
		using (v.VisitBlockBegin(block, s))
		{
			switch (block)
			{
				case Paragraph p:
					WalkParagraph(p, s, v);
					break;

				case Table t:
					WalkTable(t, s, v);
					break;

				case BookmarkStart bs:
					{
						var name = bs.Name?.Value;

						if (v is visitors.DxpMarkdownVisitor mdv &&
							!mdv.EmitUnreferencedBookmarks &&
							!string.IsNullOrEmpty(name) &&
							!_referencedAnchors.Contains(name!))
						{
							_style.ResetStyle(v);
							return;
						}
						v.VisitBookmarkStart(bs, s);
						_style.ResetStyle(v);
						return;
					}
				case BookmarkEnd be:
					v.VisitBookmarkEnd(be, s);
					_style.ResetStyle(v);
					return;

				case SectionProperties sp:
					WalkSectionProperties(sp, (WordprocessingDocument)_main?.OpenXmlPackage!, s, v, ReferenceEquals(sp, lastSectPr));
					_style.ResetStyle(v);
					break;

				case SdtBlock sdt:
					WalkSdtBlock(sdt, s, v);
					break;

				case CustomXmlBlock cx:
					WalkCustomXmlBlock(cx, s, v);
					break;

				case AltChunk ac:
					v.VisitAltChunk(ac, s);
					break;

				case ContentPart cp:
					v.VisitContentPart(cp, s);
					break;

				// Anchors / permissions / proofing
				case CommentRangeStart crs:
				{
					string id = crs.Id?.Value ?? string.Empty;
					TryEmitInlineComment(id, s, v);
					break;
				}
				case CommentRangeEnd:
					// nothing to do for inline-at-start policy
					break;
				case PermStart ps:
					v.VisitPermStart(ps, s);
					break;
				case PermEnd pe:
					v.VisitPermEnd(pe, s);
					break;
				case ProofError per:
					v.VisitProofError(per, s);
					break;

				// Tracked-change markers at block scope
				case Inserted ins:
					WalkInserted(ins, s, v);
					break;
				case Deleted del:
					WalkDeleted(del, s, v);
					break;

				// Move ranges (location containers)
				case MoveFromRangeStart mfrs:
					v.VisitMoveFromRangeStart(mfrs, s);
					break;
				case MoveFromRangeEnd mfre:
					v.VisitMoveFromRangeEnd(mfre, s);
					break;
				case MoveToRangeStart mtrs:
					v.VisitMoveToRangeStart(mtrs, s);
					break;
				case MoveToRangeEnd mtre:
					v.VisitMoveToRangeEnd(mtre, s);
					break;

				// customXml range markup (start/end) + Office 2010 conflict ranges
				case CustomXmlInsRangeStart cxInsS:
					v.VisitCustomXmlInsRangeStart(cxInsS, s);
					break;
				case CustomXmlInsRangeEnd cxInsE:
					v.VisitCustomXmlInsRangeEnd(cxInsE, s);
					break;
				case CustomXmlDelRangeStart cxDelS:
					v.VisitCustomXmlDelRangeStart(cxDelS, s);
					break;
				case CustomXmlDelRangeEnd cxDelE:
					v.VisitCustomXmlDelRangeEnd(cxDelE, s);
					break;
				case CustomXmlMoveFromRangeStart cxMfS:
					v.VisitCustomXmlMoveFromRangeStart(cxMfS, s);
					break;
				case CustomXmlMoveFromRangeEnd cxMfE:
					v.VisitCustomXmlMoveFromRangeEnd(cxMfE, s);
					break;
				case CustomXmlMoveToRangeStart cxMtS:
					v.VisitCustomXmlMoveToRangeStart(cxMtS, s);
					break;
				case CustomXmlMoveToRangeEnd cxMtE:
					v.VisitCustomXmlMoveToRangeEnd(cxMtE, s);
					break;

				case DocumentFormat.OpenXml.Office2010.Word.ConflictInsertion cIns:
					v.VisitConflictInsertion(cIns, s);
					break; // w14:conflictIns
				case DocumentFormat.OpenXml.Office2010.Word.ConflictDeletion cDel:
					v.VisitConflictDeletion(cDel, s);
					break; // w14:conflictDel
				case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeStart cxCis:
					v.VisitCustomXmlConflictInsertionRangeStart(cxCis, s);
					break;
				case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeEnd cxCie:
					v.VisitCustomXmlConflictInsertionRangeEnd(cxCie, s);
					break;
				case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeStart cxCds:
					v.VisitCustomXmlConflictDeletionRangeStart(cxCds, s);
					break;
				case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeEnd cxCde:
					v.VisitCustomXmlConflictDeletionRangeEnd(cxCde, s);
					break;

				default:
					ForwardUnknown("Block", block, s, v);
					break;
			}
		}
	}

	private void WalkSectionProperties(SectionProperties sp, WordprocessingDocument doc, IDxpStyleResolver s, IDxpVisitor v, bool isLastSection)
	{
			// 1) Snapshot the layout bits to expose
		var layout = CreateSectionLayout(sp);
		if (!isLastSection)
			v.VisitSectionLayout(sp, layout, s);


		bool emitSectionContent = v is visitors.DxpMarkdownVisitor mdv
			? mdv.EmitSectionHeadersFooters
			: true;

		if (emitSectionContent)
		{
			RenderSectionHeaders(sp, s, v);
			RenderSectionFooters(sp, s, v);
		}
	}

	private void RenderSectionHeaders(SectionProperties sp, IDxpStyleResolver s, IDxpVisitor v)
	{
		var headerRef = PickHeaderFooterReference(sp.Elements<HeaderReference>(), sp, preferFirst: true);
		if (headerRef == null)
			return;

		RenderHeaderReference(headerRef, s, v);
	}

	private void RenderSectionFooters(SectionProperties sp, IDxpStyleResolver s, IDxpVisitor v)
	{
		var footerRef = PickHeaderFooterReference(sp.Elements<FooterReference>(), sp, preferFirst: false);
		if (footerRef == null)
			return;

		RenderFooterReference(footerRef, s, v);
	}

	private HeaderReference? PickHeaderFooterReference(IEnumerable<HeaderReference> refs, SectionProperties sp, bool preferFirst)
	{
		return PickReference(refs, sp, preferFirst, r => r.Type?.Value);
	}

	private FooterReference? PickHeaderFooterReference(IEnumerable<FooterReference> refs, SectionProperties sp, bool preferFirst)
	{
		return PickReference(refs, sp, preferFirst, r => r.Type?.Value);
	}

	private T? PickReference<T>(IEnumerable<T> refs, SectionProperties sp, bool preferFirst, Func<T, HeaderFooterValues?> typeSelector)
		where T : class
	{
		var list = refs?.ToList();
		if (list == null || list.Count == 0)
			return null;

		bool useFirst = preferFirst && sp.GetFirstChild<TitlePage>() != null;

		var ordered = useFirst
			? new[] { HeaderFooterValues.First, HeaderFooterValues.Default, HeaderFooterValues.Even }
			: new[] { HeaderFooterValues.Default, HeaderFooterValues.First, HeaderFooterValues.Even };

		foreach (var target in ordered)
		{
			var match = list.FirstOrDefault(r => NormalizeHeaderFooterType(typeSelector(r)) == target);
			if (match != null)
				return match;
		}

		return list.FirstOrDefault();
	}

	private static HeaderFooterValues NormalizeHeaderFooterType(HeaderFooterValues? type)
	{
		return type ?? HeaderFooterValues.Default;
	}

	private void RenderHeaderReference(HeaderReference hr, IDxpStyleResolver s, IDxpVisitor v)
	{
		var relId = hr.Id?.Value;
		if (string.IsNullOrEmpty(relId))
			return;

		if (_main?.GetPartById(relId!) is HeaderPart part && part.Header is Header hdr)
		{
			var kind = hr.Type?.Value ?? HeaderFooterValues.Default;
			using (PushCurrentPart(part))
			using (v.VisitSectionHeaderBegin(hdr, kind, s))
			{
				foreach (var child in hdr.ChildElements)
					WalkBlock(child, s, v);
			}
			_style.ResetStyle(v);
		}
	}

	private void RenderFooterReference(FooterReference fr, IDxpStyleResolver s, IDxpVisitor v)
	{
		var relId = fr.Id?.Value;
		if (string.IsNullOrEmpty(relId))
			return;

		if (_main?.GetPartById(relId!) is FooterPart part && part.Footer is Footer ftr)
		{
			var kind = fr.Type?.Value ?? HeaderFooterValues.Default;
			using (PushCurrentPart(part))
			using (v.VisitSectionFooterBegin(ftr, kind, s))
			{
				foreach (var child in ftr.ChildElements)
					WalkBlock(child, s, v);
			}
			_style.ResetStyle(v);
		}
	}

	private static SectionLayout CreateSectionLayout(SectionProperties sp)
	{
		return new SectionLayout
		{
			PageSize = sp.GetFirstChild<PageSize>(),
			PageMargin = sp.GetFirstChild<PageMargin>(),
			Columns = sp.GetFirstChild<Columns>(),
			DocGrid = sp.GetFirstChild<DocGrid>(),
			PageBorders = sp.GetFirstChild<PageBorders>(),
			LineNumbers = sp.GetFirstChild<LineNumberType>(),
			TextDirection = sp.GetFirstChild<TextDirection>(),
			VerticalJustification = sp.GetFirstChild<VerticalTextAlignment>(),
			FootnoteProperties = sp.GetFirstChild<FootnoteProperties>(),
			EndnoteProperties = sp.GetFirstChild<EndnoteProperties>(),
		};
	}


	private void WalkTable(Table tbl, IDxpStyleResolver s, IDxpVisitor v)
	{

		DxpTableModel model = _tables.BuildTableModel(tbl);

		// Surface table properties before opening tag so visitors can style the table start.
		var tblPr = tbl.GetFirstChild<TableProperties>();
		if (tblPr != null)
			v.VisitTableProperties(tblPr, s);

		using (v.VisitTableBegin(tbl, model, s))
		{
			// Table-level props/grid/anchors
			foreach (OpenXmlElement child in tbl.ChildElements)
			{
				switch (child)
				{
					case TableProperties:
						// Already surfaced above.
						break;
					case TableGrid grid:
						v.VisitTableGrid(grid, s); // columns & default widths
						break;

					case TableRow tr:
						WalkTableRow(tr, s, v);
						break;

					case SdtRow sdtRow:
						using (v.VisitSdtRowBegin(sdtRow, s))
						{
							var content = sdtRow.SdtContentRow;
							if (content != null)
							{
								foreach (var inner in content.Elements<TableRow>())
									WalkTableRow(inner, s, v);
							}
						}
						break;

					case CustomXmlRow cxRow:
						using (v.VisitCustomXmlRowBegin(cxRow, s))
						{
							foreach (var inner in cxRow.Elements<TableRow>())
								WalkTableRow(inner, s, v);
						}
						break;

					// Anchors / tracked ranges allowed under w:tbl
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, s);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, s);
						break;
					case CommentRangeStart crs:
						v.VisitCommentRangeStart(crs, s);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, s);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, s);
						break;
					case ProofError perr:
						v.VisitProofError(perr, s);
						break;
					case InsertedRun ins:
						WalkInsertedRun(ins, s, v);
						break;
					case DeletedRun del:
						WalkDeletedRun(del, s, v);
						break;
					case MoveFromRun mfr:
						v.VisitMoveFromRun(mfr, s);
						break;
					case MoveToRun mtr:
						v.VisitMoveToRun(mtr, s);
						break;

						// customXml range start/end + Office 2010 conflict ranges also valid here
						// (handled the same way as other dispatchers)

					// Rare but permitted under w:tbl (SDK lists these)
					case DocumentFormat.OpenXml.Math.OfficeMath m:
						v.VisitOMath(m, s);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph mp:
						v.VisitOMathParagraph(mp, s);
						break;

					default:
						ForwardUnknown("Table", child, s, v);
						break;
				}
			}
			_style.ResetStyle(v);
		}
	}

	private void WalkTableRow(TableRow tr, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (v.VisitTableRowBegin(tr, s))
		{
			foreach (var child in tr.ChildElements)
			{
				switch (child)
				{
					case TableCell tc:
						WalkTableCell(tc, s, v);
						break;

					case TableRowProperties trp:
						v.VisitTableRowProperties(trp, s);
						break;

					case SdtCell sdtCell:
					{
						using (v.VisitSdtCellBegin(sdtCell, s))
						{
							// <w:sdtCell><w:sdtContent><w:tc>…</w:tc></w:sdtContent></w:sdtCell>
							var content = sdtCell.SdtContentCell;
							if (content != null)
							{
								foreach (var inner in content.Elements<TableCell>())
									WalkTableCell(inner, s, v);
							}
						}
						break;
					}

					case CustomXmlCell cxCell:
					{
						using (v.VisitCustomXmlCellBegin(cxCell, s))
						{
							// <w:customXmlCell> may contain one or more <w:tc>
							foreach (var inner in cxCell.Elements<TableCell>())
								WalkTableCell(inner, s, v);
						}
						break;
					}

						// Keep other row-level cases (e.g., trPr) or forward unknowns:
					default:
						ForwardUnknown("TableRow child", child, s, v);
						break;
				}
			}
		}
	}




	private void WalkTableCell(TableCell tc, IDxpStyleResolver s, IDxpVisitor v)
	{
			// When a grid/merge model is computed, call VisitTableCellLayout here with resolved spans.
		using (v.VisitTableCellBegin(tc, s))
		{

			bool sawBlock = false; // enforce CT_Tc rule: at least one block-level element

			foreach (var child in tc.ChildElements)
			{
				switch (child)
				{
					case TableCellProperties tcp:
						v.VisitTableCellProperties(tcp, s);
						break;

					// ---- Block-level content inside a cell (EG_BlockLevelElts) ----
					case Paragraph p:
						sawBlock = true;
						WalkBlock(p, s, v);
						break;

					case Table t:
						sawBlock = true;
						WalkBlock(t, s, v);
						break;

					case SdtBlock sdt:
						sawBlock = true;
						WalkSdtBlock(sdt, s, v);
						break;

					case CustomXmlBlock cx:
						sawBlock = true;
						WalkCustomXmlBlock(cx, s, v);
						break;

					case AltChunk ac:
						sawBlock = true;
						v.VisitAltChunk(ac, s);
						break;

					// ---- Math (allowed directly under tc) ----
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						sawBlock = true;
						v.VisitOMathParagraph(oMathPara, s);
						break;
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						sawBlock = true;
						v.VisitOMath(oMath, s);
						break;

					// ---- Anchors / permissions / proofing (directly allowed in tc) ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, s);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, s);
						break;
					case CommentRangeStart crs:
						v.VisitCommentRangeStart(crs, s);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, s);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, s);
						break;
					case ProofError perr:
						v.VisitProofError(perr, s);
						break;

					// ---- Tracked-change containers and range start/end (directly allowed) ----
					case Inserted ins:
						WalkInserted(ins, s, v);
						break;
					case Deleted del:
						WalkDeleted(del, s, v);
						break;
					case MoveFromRun mfr:
						v.VisitMoveFromRun(mfr, s);
						break;
					case MoveToRun mtr:
						v.VisitMoveToRun(mtr, s);
						break;
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, s);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, s);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, s);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, s);
						break;

					// ---- customXml range markup (start/end) ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, s);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, s);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, s);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, s);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, s);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, s);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, s);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, s);
						break;

					default:
						ForwardUnknown("TableCell", child, s, v);
						break;
				}
			}

			if (!sawBlock)
				throw new InvalidOperationException("w:tc must contain at least one block-level element (p/tbl/sdt/customXml/altChunk/math).");
		}
	}


	private void WalkParagraph(Paragraph p, IDxpStyleResolver s, IDxpVisitor v)
	{
		if (!HasRenderableParagraphContent(p))
			return;

		(string? marker, int? numId, int? iLvl) = _lists.MaterializeMarker(p, s);
		DxpStyleEffectiveIndentTwips indent = _lists.GetIndentation(p, s);

		using (v.VisitParagraphBegin(p, s, marker, numId, iLvl, indent))
		{
			int fieldDepthAtEnter = _fieldStack.Count;

			foreach (var child in p.ChildElements)
			{
				switch (child)
				{
					case ProofError pe:
						v.VisitProofError(pe, s);
						break;

					case DeletedRun dr:
						WalkDeletedRun(dr, s, v);
						break;

					case InsertedRun ir:
						WalkInsertedRun(ir, s, v);
						break;

					case ParagraphProperties pp:
						v.VisitParagraphProperties(pp, s);
						break;

					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, s);
						break;

					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, s);
						break;

					case Run r:
						WalkRun(r, s, v);
						break;

					case Hyperlink link:
						WalkHyperlink(link, s, v);
						break;

					case CommentRangeStart crs:
					{
						string id = crs.Id?.Value ?? string.Empty;
						TryEmitInlineComment(id, s, v);
						break;
					}
					case CommentRangeEnd:
						// nothing to do for inline-at-start policy
						break;

						case CustomXmlRun cxr:
							WalkCustomXmlRun(cxr, s, v);
							break; // w:customXml (inline custom XML container).

						case SimpleField fld:
							WalkSimpleField(fld, s, v);
							break; // w:fldSimple (simple field; contains runs/hyperlinks).

						case SdtRun sdtRun:
							WalkSdtRun(sdtRun, s, v);
							break; // w:sdt (run-level SDT).

						case PermStart ps:
							v.VisitPermStart(ps, s);
							break; // w:permStart (editing permission range start).
						case PermEnd pe:
							v.VisitPermEnd(pe, s);
							break; // w:permEnd (editing permission range end).

						case MoveFromRangeStart mfrs:
							v.VisitMoveFromRangeStart(mfrs, s);
							break; // w:moveFromRangeStart (tracked move-out start).
						case MoveFromRangeEnd mfre:
							v.VisitMoveFromRangeEnd(mfre, s);
							break; // w:moveFromRangeEnd (tracked move-out end).
						case MoveToRangeStart mtrs:
							v.VisitMoveToRangeStart(mtrs, s);
							break; // w:moveToRangeStart (tracked move-in start).
						case MoveToRangeEnd mtre:
							v.VisitMoveToRangeEnd(mtre, s);
							break; // w:moveToRangeEnd (tracked move-in end).

						case CustomXmlInsRangeStart cxInsS:
							v.VisitCustomXmlInsRangeStart(cxInsS, s);
							break; // w:customXmlInsRangeStart (customXml insert range start).
						case CustomXmlInsRangeEnd cxInsE:
							v.VisitCustomXmlInsRangeEnd(cxInsE, s);
							break; // w:customXmlInsRangeEnd (customXml insert range end).
						case CustomXmlDelRangeStart cxDelS:
							v.VisitCustomXmlDelRangeStart(cxDelS, s);
							break; // w:customXmlDelRangeStart (customXml delete range start).
						case CustomXmlDelRangeEnd cxDelE:
							v.VisitCustomXmlDelRangeEnd(cxDelE, s);
							break; // w:customXmlDelRangeEnd (customXml delete range end).
						case CustomXmlMoveFromRangeStart cxMfS:
							v.VisitCustomXmlMoveFromRangeStart(cxMfS, s);
							break; // w:customXmlMoveFromRangeStart (customXml move-from start).
						case CustomXmlMoveFromRangeEnd cxMfE:
							v.VisitCustomXmlMoveFromRangeEnd(cxMfE, s);
							break; // w:customXmlMoveFromRangeEnd (customXml move-from end).
						case CustomXmlMoveToRangeStart cxMtS:
							v.VisitCustomXmlMoveToRangeStart(cxMtS, s);
							break; // w:customXmlMoveToRangeStart (customXml move-to start).
						case CustomXmlMoveToRangeEnd cxMtE:
							v.VisitCustomXmlMoveToRangeEnd(cxMtE, s);
							break; // w:customXmlMoveToRangeEnd (customXml move-to end).

						case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeStart cxCis:
							v.VisitCustomXmlConflictInsertionRangeStart(cxCis, s);
							break; // w14:customXmlConflictInsRangeStart (Office 2010 conflict insert start).
						case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeEnd cxCie:
							v.VisitCustomXmlConflictInsertionRangeEnd(cxCie, s);
							break; // w14:customXmlConflictInsRangeEnd (Office 2010 conflict insert end).
						case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeStart cxCds:
							v.VisitCustomXmlConflictDeletionRangeStart(cxCds, s);
							break; // w14:customXmlConflictDelRangeStart (Office 2010 conflict delete start).
						case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeEnd cxCde:
							v.VisitCustomXmlConflictDeletionRangeEnd(cxCde, s);
							break; // w14:customXmlConflictDelRangeEnd (Office 2010 conflict delete end).

						case MoveFromRun mfr:
							v.VisitMoveFromRun(mfr, s);
							break; // w:moveFrom (run container for moved-out text).
						case MoveToRun mtr:
							v.VisitMoveToRun(mtr, s);
							break; // w:moveTo (run container for moved-in text).

						case ContentPart cp:
							v.VisitContentPart(cp, s);
							break; // w:contentPart (external content reference; Office 2010+).

						case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
							v.VisitOMathParagraph(oMathPara, s);
							break; // m:oMathPara (display math paragraph).
						case DocumentFormat.OpenXml.Math.OfficeMath oMath:
							v.VisitOMath(oMath, s);
							break; // m:oMath (inline math object).

						case DocumentFormat.OpenXml.Math.Accent mAccent:
							v.VisitOMathElement(mAccent, s);
							break; // m:acc (math accent – direct child allowed).
						case DocumentFormat.OpenXml.Math.Bar mBar:
							v.VisitOMathElement(mBar, s);
							break; // m:bar (math bar).
						case DocumentFormat.OpenXml.Math.Box mBox:
							v.VisitOMathElement(mBox, s);
							break; // m:box (math box).
						case DocumentFormat.OpenXml.Math.BorderBox mBorderBox:
							v.VisitOMathElement(mBorderBox, s);
							break; // m:borderBox (math border box).
						case DocumentFormat.OpenXml.Math.Delimiter mDelim:
							v.VisitOMathElement(mDelim, s);
							break; // m:d (delimiter).
						case DocumentFormat.OpenXml.Math.EquationArray mEqArr:
							v.VisitOMathElement(mEqArr, s);
							break; // m:eqArr (equation array).
						case DocumentFormat.OpenXml.Math.Fraction mFrac:
							v.VisitOMathElement(mFrac, s);
							break; // m:f (fraction).
						case DocumentFormat.OpenXml.Math.MathFunction mFunc:
							v.VisitOMathElement(mFunc, s);
							break; // m:func (function).
						case DocumentFormat.OpenXml.Math.GroupChar mGroupChr:
							v.VisitOMathElement(mGroupChr, s);
							break; // m:groupChr (group character).
						case DocumentFormat.OpenXml.Math.LimitLower mLimLow:
							v.VisitOMathElement(mLimLow, s);
							break; // m:limLow (lower limit).
						case DocumentFormat.OpenXml.Math.LimitUpper mLimUpp:
							v.VisitOMathElement(mLimUpp, s);
							break; // m:limUpp (upper limit).
						case DocumentFormat.OpenXml.Math.Matrix mMat:
							v.VisitOMathElement(mMat, s);
							break; // m:m (matrix).
						case DocumentFormat.OpenXml.Math.Nary mNary:
							v.VisitOMathElement(mNary, s);
							break; // m:nary (n-ary operator).
						case DocumentFormat.OpenXml.Math.Phantom mPhant:
							v.VisitOMathElement(mPhant, s);
							break; // m:phantom (phantom).
						case DocumentFormat.OpenXml.Math.Radical mRad:
							v.VisitOMathElement(mRad, s);
							break; // m:rad (radical).
						case DocumentFormat.OpenXml.Math.PreSubSuper mPreSubSup:
							v.VisitOMathElement(mPreSubSup, s);
							break; // m:preSubSup (presub/superscript).
						case DocumentFormat.OpenXml.Math.Subscript mSub:
							v.VisitOMathElement(mSub, s);
							break; // m:s (subscript).
						case DocumentFormat.OpenXml.Math.SubSuperscript mSubSup:
							v.VisitOMathElement(mSubSup, s);
							break; // m:sSub (sub-superscript).
						case DocumentFormat.OpenXml.Math.Superscript mSup:
							v.VisitOMathElement(mSup, s);
							break; // m:sup (superscript).
						case DocumentFormat.OpenXml.Math.Run mMathRun:
							v.VisitOMathRun(mMathRun, s);
							break; // m:r (math run).

						case BidirectionalOverride bdo:
							WalkBidirectionalOverride(bdo, s, v);
							break; // w:bdo (Bidi override; Office 2010+).
						case BidirectionalEmbedding bdi:
							WalkBidirectionalEmbedding(bdi, s, v);
							break; // w:dir (Bidi embedding; Office 2010+).

						case SubDocumentReference subDoc:
							v.VisitSubDocumentReference(subDoc, s);
							break; // w:subDoc (subdocument anchor).

					case AlternateContent ac:
						WalkAlternateContent(ac, s, v);
						break; // mc:AlternateContent (compat wrapper around inline kids). (MC allows it here)

					// Comment anchors / tracked change containers will show up as elements too.
					// These currently throw to surface unhandled cases.
				default:
					ForwardUnknown("Paragraph", child, s, v);
					break;
			}
			}

			// on exit: if a producer emitted unterminated fields, clean them up
			while (_fieldStack.Count > fieldDepthAtEnter)
			{
				var frame = _fieldStack.Pop();
				frame.ResultScope?.Dispose();
				// no FieldChar 'end' to pass to the visitor in this broken case
			}
			_style.ResetStyle(v);
		}
	}

	private static bool HasRenderableParagraphContent(Paragraph p)
	{
		// Render if there is any non-empty text, drawings, breaks, or tabs.
		if (p.Descendants<Text>().Any(t => !string.IsNullOrEmpty(t.Text)))
			return true;
		if (p.Descendants<Drawing>().Any())
			return true;
		if (p.Descendants<Break>().Any() || p.Descendants<CarriageReturn>().Any() || p.Descendants<TabChar>().Any())
			return true;
		return false;
	}

	private void WalkBidirectionalEmbedding(BidirectionalEmbedding bdi, IDxpStyleResolver s, IDxpVisitor v)
	{
		// w:dir (Bidirectional Embedding). Attribute w:val is the embedding direction (e.g., "rtl"/"ltr").
		// Children per CT_DirContentRun include inline content, math, tracked ranges, customXml ranges, and nesting. 
		// Ref: SDK "BidirectionalEmbedding" child list & examples. 
		using (v.VisitBidirectionalEmbeddingBegin(bdi, s))
		{
			foreach (var child in bdi.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content ----
					case Run r:
						WalkRun(r, s, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, s, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, s, v);
						break;
					case SimpleField fld:
						WalkSimpleField(fld, s, v);
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, s, v);
						break;

					// ---- Range anchors / bookmarks / comments / permissions ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, s);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, s);
						break;
					case CommentRangeStart crs:
						v.VisitCommentRangeStart(crs, s);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, s);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, s);
						break;
					case ProofError perr:
						v.VisitProofError(perr, s);
						break;

					// ---- Move ranges (location containers) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, s);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, s);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, s);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, s);
						break;

					// ---- customXml ranges (start/end) + Office 2010 conflict ranges ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, s);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, s);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, s);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, s);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, s);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, s);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, s);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, s);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeStart cxCis:
						v.VisitCustomXmlConflictInsertionRangeStart(cxCis, s);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeEnd cxCie:
						v.VisitCustomXmlConflictInsertionRangeEnd(cxCie, s);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeStart cxCds:
						v.VisitCustomXmlConflictDeletionRangeStart(cxCds, s);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeEnd cxCde:
						v.VisitCustomXmlConflictDeletionRangeEnd(cxCde, s);
						break;

					// ---- Tracked-change run containers ----
					case InsertedRun insRun:
						WalkInsertedRun(insRun, s, v);
						break;
					case DeletedRun delRun:
						WalkDeletedRun(delRun, s, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, s);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, s);
						break;

					// ---- Office Math (inline & display) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, s);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, s);
						break;

					// ---- Nesting: bdo/dir/subDoc are allowed inside w:dir ----
					case BidirectionalOverride bdo:
						WalkBidirectionalOverride(bdo, s, v);
						break;
					case BidirectionalEmbedding nestedDir:
						WalkBidirectionalEmbedding(nestedDir, s, v);
						break;
					case SubDocumentReference subDoc:
						v.VisitSubDocumentReference(subDoc, s);
						break;

					// ---- MC wrapper (commonly appears anywhere inline) ----
					case AlternateContent ac:
						WalkAlternateContent(ac, s, v);
						break;

					default:
						ForwardUnknown("BidirectionalEmbedding", child, s, v);
						break;
				}
			}
		}
	}

	private void WalkBidirectionalOverride(BidirectionalOverride bdo, IDxpStyleResolver s, IDxpVisitor v)
	{
		// w:bdo (BiDi override). Direction via w:val (e.g., "rtl"/"ltr").
		using (v.VisitBidirectionalOverrideBegin(bdo, s))
		{
			foreach (var child in bdo.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content ----
					case Run r:
						WalkRun(r, s, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, s, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, s, v);
						break;
					case SimpleField fld:
						WalkSimpleField(fld, s, v);
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, s, v);
						break;
					// Not found in SDK
					case OpenXmlUnknownElement smart
						when smart.LocalName == "smartTag" && smart.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main":
					{
						WalkSmartTagRun(smart, s, v);
						break;
					}

					// ---- Range anchors / bookmarks / comments / permissions ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, s);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, s);
						break;
					case CommentRangeStart crs:
						v.VisitCommentRangeStart(crs, s);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, s);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, s);
						break;
					case ProofError perr:
						v.VisitProofError(perr, s);
						break;

					// ---- Move locations (range start/end) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, s);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, s);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, s);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, s);
						break;

					// ---- customXml ranges (start/end) ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, s);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, s);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, s);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, s);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, s);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, s);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, s);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, s);
						break;

					// ---- Tracked-change run containers ----
					case InsertedRun insRun:
						WalkInsertedRun(insRun, s, v);
						break;
					case DeletedRun delRun:
						WalkDeletedRun(delRun, s, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, s);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, s);
						break;

					// ---- Office Math (inline & display) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, s);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, s);
						break;

					// ---- Nesting (bdo/dir/subDoc allowed under paragraph content) ----
					case BidirectionalOverride nestedBdo:
						WalkBidirectionalOverride(nestedBdo, s, v);
						break;
					case BidirectionalEmbedding nestedDir:
						WalkBidirectionalEmbedding(nestedDir, s, v);
						break;
					case SubDocumentReference subDoc:
						v.VisitSubDocumentReference(subDoc, s);
						break;

					// ---- MC wrapper ----
					case AlternateContent ac:
						WalkAlternateContent(ac, s, v);
						break;

					default:
						ForwardUnknown("BidirectionalOverride", child, s, v);
						break;
				}
			}
		}
	}

	private static bool IsWSmartTag(OpenXmlUnknownElement unk) =>
		unk != null
		&& unk.LocalName == "smartTag"
		&& (unk.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
			|| unk.NamespaceUri == "http://purl.oclc.org/ooxml/wordprocessingml/main"); // some tools map to purl

	private void WalkSmartTagRun(OpenXmlUnknownElement smart, IDxpStyleResolver s, IDxpVisitor v)
	{
		if (!IsWSmartTag(smart))
			throw Unsupported("SmartTag (expected w:smartTag)", smart);

		// Extract smartTag attributes per spec: w:element, w:uri
		var wNs = smart.NamespaceUri; // current w namespace
		string elementName = smart.GetAttribute("element", wNs).Value ?? string.Empty;
		string elementUri = smart.GetAttribute("uri", wNs).Value ?? string.Empty;

			using (v.VisitSmartTagRunBegin(smart, elementName, elementUri, s))
			{
				// Optional properties child: <w:smartTagPr> with <w:attr> entries
				// SDK class for <w:attr> is Wordprocessing.CustomXmlAttribute (yes, despite the name).
			var smartTagPr = smart.ChildElements
				.OfType<OpenXmlUnknownElement>()
				.FirstOrDefault(e => e.LocalName == "smartTagPr" && e.NamespaceUri == wNs);
			if (smartTagPr != null)
			{
				// Surface attributes/properties to the visitor
				var attrs = smartTagPr.Elements<CustomXmlAttribute>().ToList();
				v.VisitSmartTagProperties(smartTagPr, attrs, s);
			}

			// Walk run-level content (everything except <w:smartTagPr>)
			foreach (var child in smart.ChildElements)
			{
				if (child == smartTagPr)
					continue;

				switch (child)
				{
					// ---- Core inline content ----
					case Run r:
						WalkRun(r, s, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, s, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, s, v);
						break;
					case SimpleField fld:
						WalkSimpleField(fld, s, v);
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, s, v);
						break;

					// Drawings & legacy pict can appear inline here
					case Drawing d:
						TryWalkDrawingTextBox(d, s, v);
						break;
					case Picture pict:
						TryWalkVmlTextBox(pict, s, v);
						break;

					// Anchors / permissions / proofing / range markup
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, s);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, s);
						break;
					case CommentRangeStart crs:
						v.VisitCommentRangeStart(crs, s);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, s);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, s);
						break;
					case ProofError perr:
						v.VisitProofError(perr, s);
						break;

					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, s);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, s);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, s);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, s);
						break;

					// Tracked-change run containers
					case InsertedRun insRun:
						WalkInsertedRun(insRun, s, v);
						break;
					case DeletedRun delRun:
						WalkDeletedRun(delRun, s, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, s);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, s);
						break;

					// Office Math (inline & display)
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, s);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, s);
						break;

					// MC wrapper (smartTag can contain AlternateContent)
					case AlternateContent ac:
						WalkAlternateContent(ac, s, v);
						break;

					// Nested smartTag (rare but legal) – recurse on unknown
					case OpenXmlUnknownElement unk when IsWSmartTag(unk):
						WalkSmartTagRun(unk, s, v);
						break;

					default:
						ForwardUnknown("SmartTag", child, s, v);
						break;
				}
			}
		}
	}

	private void WalkSdtRun(SdtRun sdtRun, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (v.VisitSdtRunBegin(sdtRun, s))
		{
			// (1) Properties (optional)
			var pr = sdtRun.SdtProperties;
			if (pr != null)
				v.VisitSdtProperties(pr, s);

			// (2) Content (optional per schema — do NOT throw if missing)
			var content = sdtRun.SdtContentRun;
			if (content != null)
			{
				using (v.VisitSdtContentRunBegin(content, s))
				{
					foreach (var child in content.ChildElements)
					{
						switch (child)
						{
							// ---- Core inline content ----
							case Run r:
								WalkRun(r, s, v);
								break;
							case Hyperlink link:
								WalkHyperlink(link, s, v);
								break;
							case SdtRun nestedSdt:
								WalkSdtRun(nestedSdt, s, v);
								break;
							case SimpleField fld:
								WalkSimpleField(fld, s, v);
								break;
							case CustomXmlRun cxr:
								WalkCustomXmlRun(cxr, s, v);
								break;
							// Not found in SDK: w:smartTag serialized as unknown
							case OpenXmlUnknownElement smart
								when smart.LocalName == "smartTag"
									&& (smart.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
										|| smart.NamespaceUri == "http://purl.oclc.org/ooxml/wordprocessingml/main"):
								WalkSmartTagRun(smart, s, v);
								break;

							// ---- Anchors / range markup ----
							case BookmarkStart bs:
								v.VisitBookmarkStart(bs, s);
								break;
							case BookmarkEnd be:
								v.VisitBookmarkEnd(be, s);
								break;
							case CommentRangeStart crs:
								v.VisitCommentRangeStart(crs, s);
								break;
							case CommentRangeEnd cre:
								break;
							case PermStart ps:
								v.VisitPermStart(ps, s);
								break;
							case PermEnd pe:
								v.VisitPermEnd(pe, s);
								break;
							case ProofError perr:
								v.VisitProofError(perr, s);
								break;

							// ---- Move ranges (location containers) ----
							case MoveFromRangeStart mfrs:
								v.VisitMoveFromRangeStart(mfrs, s);
								break;
							case MoveFromRangeEnd mfre:
								v.VisitMoveFromRangeEnd(mfre, s);
								break;
							case MoveToRangeStart mtrs:
								v.VisitMoveToRangeStart(mtrs, s);
								break;
							case MoveToRangeEnd mtre:
								v.VisitMoveToRangeEnd(mtre, s);
								break;

							// ---- customXml ranges (start/end) ----
							case CustomXmlInsRangeStart cxInsS:
								v.VisitCustomXmlInsRangeStart(cxInsS, s);
								break;
							case CustomXmlInsRangeEnd cxInsE:
								v.VisitCustomXmlInsRangeEnd(cxInsE, s);
								break;
							case CustomXmlDelRangeStart cxDelS:
								v.VisitCustomXmlDelRangeStart(cxDelS, s);
								break;
							case CustomXmlDelRangeEnd cxDelE:
								v.VisitCustomXmlDelRangeEnd(cxDelE, s);
								break;
							case CustomXmlMoveFromRangeStart cxMfS:
								v.VisitCustomXmlMoveFromRangeStart(cxMfS, s);
								break;
							case CustomXmlMoveFromRangeEnd cxMfE:
								v.VisitCustomXmlMoveFromRangeEnd(cxMfE, s);
								break;
							case CustomXmlMoveToRangeStart cxMtS:
								v.VisitCustomXmlMoveToRangeStart(cxMtS, s);
								break;
							case CustomXmlMoveToRangeEnd cxMtE:
								v.VisitCustomXmlMoveToRangeEnd(cxMtE, s);
								break;

							// ---- Tracked-change run containers ----
							case InsertedRun insRun:
								WalkInsertedRun(insRun, s, v);
								break;
							case DeletedRun delRun:
								WalkDeletedRun(delRun, s, v);
								break;
							case MoveFromRun moveFromRun:
								v.VisitMoveFromRun(moveFromRun, s);
								break;
							case MoveToRun moveToRun:
								v.VisitMoveToRun(moveToRun, s);
								break;

							// ---- Office Math (inline & display) ----
							case DocumentFormat.OpenXml.Math.OfficeMath oMath:
								v.VisitOMath(oMath, s);
								break;
							case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
								v.VisitOMathParagraph(oMathPara, s);
								break;

							// ---- Bidi & subdocument ----
							case BidirectionalOverride bdo:
								WalkBidirectionalOverride(bdo, s, v);
								break;
							case BidirectionalEmbedding bdi:
								WalkBidirectionalEmbedding(bdi, s, v);
								break;
							case SubDocumentReference subDoc:
								v.VisitSubDocumentReference(subDoc, s);
								break;

							// ---- MC wrapper ----
							case AlternateContent ac:
								WalkAlternateContent(ac, s, v);
								break;

							default:
								ForwardUnknown("SdtContentRun", child, s, v);
								break;
						}
					}
				}
			}

			// (3) End-char run properties (optional) — still visit even if content was absent
			var endPr = sdtRun.SdtEndCharProperties;
			if (endPr != null)
				v.VisitSdtEndCharProperties(endPr, s);
		}
	}



		private void WalkSimpleField(SimpleField fld, IDxpStyleResolver s, IDxpVisitor v)
		{
			// w:fldSimple – simple field whose result is represented by its child content
			// Attributes: w:instr (field code), w:dirty, w:fldLock. Behavior: children are the current field result.
		using (v.VisitSimpleFieldBegin(fld, s))
		{
			// Optional field data payload (<w:fldData>), rarely used.
			if (fld.FieldData is { } data)
				v.VisitFieldData(data, s);

			foreach (var child in fld.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content (run-universe) ----
					case Run r:
						WalkRun(r, s, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, s, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, s, v);
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, s, v);
						break;
					case OpenXmlUnknownElement smart
						when smart.LocalName == "smartTag" && smart.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main":
					{
						WalkSmartTagRun(smart, s, v);
						break;
					}

					// ---- Anchors / range markup ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, s);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, s);
						break;
					case CommentRangeStart crs:
						v.VisitCommentRangeStart(crs, s);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, s);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, s);
						break;
					case ProofError perr:
						v.VisitProofError(perr, s);
						break;

					// ---- Move ranges (location containers) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, s);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, s);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, s);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, s);
						break;

					// ---- Tracked-change run containers ----
					case InsertedRun insRun:
						WalkInsertedRun(insRun, s, v);
						break;
					case DeletedRun delRun:
						WalkDeletedRun(delRun, s, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, s);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, s);
						break;

					// ---- Office Math (inline & display) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, s);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, s);
						break;

					// ---- Bidi containers & subdocument ----
					case BidirectionalOverride bdo:
						WalkBidirectionalOverride(bdo, s, v);
						break;
					case BidirectionalEmbedding bdi:
						WalkBidirectionalEmbedding(bdi, s, v);
						break;
					case SubDocumentReference subDoc:
						v.VisitSubDocumentReference(subDoc, s);
						break;

					// ---- Markup Compatibility wrapper ----
					case AlternateContent ac:
						WalkAlternateContent(ac, s, v);
						break;

					// ---- Elements listed as parents of fldSimple, not children; fallthrough is correct ----
					// (e.g., another fldSimple wrapping this one is allowed *as parent*, not common as child.)

					default:
						ForwardUnknown("SimpleField", child, s, v);
						break;
				}
			}
		}
	}

	private void WalkCustomXmlRun(CustomXmlRun cxr, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (v.VisitCustomXmlRunBegin(cxr, s))
		{
			// Properties <w:customXmlPr> (element name/namespace, optional data binding metadata)
			var pr = cxr.CustomXmlProperties;
			if (pr != null)
				v.VisitCustomXmlProperties(pr, s);

			foreach (var child in cxr.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content (run-universe) ----
					case Run r:
						WalkRun(r, s, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, s, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, s, v);
						break;
					case SimpleField fld:
						WalkSimpleField(fld, s, v);
						break;

					// Nested inline customXml/smartTag (both valid in CT_CustomXmlRun)
					case CustomXmlRun nested:
						WalkCustomXmlRun(nested, s, v);
						break;

					// Not found in SDK
					case OpenXmlUnknownElement smart
						when smart.LocalName == "smartTag" && smart.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main":
					{
						WalkSmartTagRun(smart, s, v);
						break;
					}

					// ---- Bookmarks & comment anchors ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, s);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, s);
						break;
					case CommentRangeStart crs:
						v.VisitCommentRangeStart(crs, s);
						break;
					case CommentRangeEnd cre:
						break;

					// ---- Permissions ----
					case PermStart ps:
						v.VisitPermStart(ps, s);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, s);
						break;

					// ---- customXml range markup (insert/delete/move; start/end) ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, s);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, s);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, s);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, s);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, s);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, s);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, s);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, s);
						break;

					// ---- Tracked change/move run containers (inline) ----
					case InsertedRun insRun:
						WalkInsertedRun(insRun, s, v);
						break;
					case DeletedRun delRun:
						WalkDeletedRun(delRun, s, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, s);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, s);
						break;

					// ---- Move location containers (range start/end) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, s);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, s);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, s);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, s);
						break;

					// ---- Office Math (inline & paragraph forms) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, s);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, s);
						break;

					// ---- Bidi containers ----
					case BidirectionalOverride bdo:
						WalkBidirectionalOverride(bdo, s, v);
						break;
					case BidirectionalEmbedding bdi:
						WalkBidirectionalEmbedding(bdi, s, v);
						break;

					// ---- Subdocument anchor ----
					case SubDocumentReference subDoc:
						v.VisitSubDocumentReference(subDoc, s);
						break;

					// ---- Markup Compatibility wrapper (practically appears anywhere) ----
					case AlternateContent ac:
						WalkAlternateContent(ac, s, v);
						break;

					// ---- Proofing errors ----
					case ProofError perr:
						v.VisitProofError(perr, s);
						break;

					default:
						ForwardUnknown("CustomXmlRun", child, s, v);
						break;
				}
			}
		}
	}

	private void TryEmitInlineComment(string id, IDxpStyleResolver s, IDxpVisitor v)
	{
		var thread = _comments.GetThreadForAnchor(id);
		if (thread == null || thread.Comments.Count == 0)
			return;

		if (v is visitors.DxpMarkdownVisitor mdv)
		{
			mdv.VisitCommentThread(id, thread, s, info => WalkCommentContent(info, s, v));
			return;
		}

		v.VisitCommentThread(id, thread, s);
	}

	private void WalkCommentContent(DxpCommentInfo info, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (PushCurrentPart(info.Part ?? _currentPart))
		{
			if (info.Blocks != null && info.Blocks.Count > 0)
			{
				foreach (var block in info.Blocks)
					WalkBlock(block, s, v);

				_style.ResetStyle(v);
				return;
			}

			if (string.IsNullOrEmpty(info.Text))
				return;

			var paragraph = new Paragraph(new Run(new Text(info.Text)));
			WalkBlock(paragraph, s, v);
			_style.ResetStyle(v);
		}
	}



	private void WalkDeletedRun(DeletedRun dr, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (v.VisitDeletedRunBegin(dr, s))
		{
			foreach (var child in dr.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content allowed inside CT_RunTrackChange ----
					case Run r:
						WalkRun(r, s, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, s, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, s, v);
						break;
					case SimpleField fld:
						// Let WalkSimpleField open the visitor scope and extract instr/flags
						WalkSimpleField(fld, s, v);
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, s, v);
						break;
					// Not found in SDK
					case OpenXmlUnknownElement smart
						when smart.LocalName == "smartTag" && smart.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main":
					{
						WalkSmartTagRun(smart, s, v);
						break;
					}

					// ---- Anchors / proofing / permissions (range markup) ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, s);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, s);
						break;
					case CommentRangeStart crs:
						v.VisitCommentRangeStart(crs, s);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, s);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, s);
						break;
					case ProofError perr:
						v.VisitProofError(perr, s);
						break;

					// ---- Move ranges (location containers) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, s);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, s);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, s);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, s);
						break;

					// ---- customXml range markup (start/end) ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, s);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, s);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, s);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, s);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, s);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, s);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, s);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, s);
						break;

					// ---- Tracked-change run containers (nesting is allowed) ----
					case InsertedRun insRun:
						WalkInsertedRun(insRun, s, v);
						break;
					case DeletedRun innerDel:
						// Nested del inside del is rare but legal per CT_RunTrackChange; recurse.
						WalkDeletedRun(innerDel, s, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, s);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, s);
						break;

					// ---- Office Math (both forms are permitted) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, s);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, s);
						break;

					// ---- Markup Compatibility wrapper ----
					case AlternateContent ac:
						WalkAlternateContent(ac, s, v);
						break;

					default:
						ForwardUnknown("DeletedRun", child, s, v);
						break;
				}
			}
		}
	}


	private void WalkInsertedRun(InsertedRun ir, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (v.VisitInsertedRunBegin(ir, s))
		{
			foreach (var child in ir.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content ----
					case Run r:
						WalkRun(r, s, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, s, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, s, v);
						break;
					case SimpleField fld:
						WalkSimpleField(fld, s, v); // lets the callee open/close the visitor scope
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, s, v);
						break;
					// Not found in SDK
					case OpenXmlUnknownElement smart
						when smart.LocalName == "smartTag" && smart.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main":
					{
						WalkSmartTagRun(smart, s, v);
						break;
					}

					// ---- Anchors / proofing / permissions ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, s);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, s);
						break;
					case CommentRangeStart crs:
						v.VisitCommentRangeStart(crs, s);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, s);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, s);
						break;
					case ProofError perr:
						v.VisitProofError(perr, s);
						break;

					// ---- Move ranges (location containers) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, s);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, s);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, s);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, s);
						break;

					// ---- customXml range markup (start/end) ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, s);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, s);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, s);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, s);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, s);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, s);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, s);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, s);
						break;

					// ---- Tracked-change run containers (nesting allowed) ----
					case InsertedRun innerIns:
						WalkInsertedRun(innerIns, s, v);
						break;
					case DeletedRun dr:
						WalkDeletedRun(dr, s, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, s);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, s);
						break;

					// ---- Office Math (inline & display) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, s);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, s);
						break;

					// ---- Markup Compatibility wrapper ----
					case AlternateContent ac:
						WalkAlternateContent(ac, s, v);
						break;

					default:
						ForwardUnknown("InsertedRun", child, s, v);
						break;
				}
			}
		}
	}



	private void WalkHyperlink(Hyperlink link, IDxpStyleResolver s, IDxpVisitor v)
	{
		string? target = ResolveHyperlinkTarget(link);
		using (v.VisitHyperlinkBegin(link, target, s))
		{

			// <w:hyperlink> can host the full run-level “inline universe”
			foreach (var child in link.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content ----
					case Run r:
						WalkRun(r, s, v);
						break;
					case Hyperlink nested:
						WalkHyperlink(nested, s, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, s, v);
						break;
					case SimpleField fld:
						WalkSimpleField(fld, s, v);
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, s, v);
						break;

					// ---- Anchors / permissions / proofing ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, s);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, s);
						break;
					case CommentRangeStart crs:
						v.VisitCommentRangeStart(crs, s);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, s);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, s);
						break;
					case ProofError perr:
						v.VisitProofError(perr, s);
						break;

					// ---- Move ranges (location containers) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, s);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, s);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, s);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, s);
						break;

					// ---- customXml ranges (start/end) + Office 2010 conflict ranges ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, s);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, s);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, s);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, s);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, s);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, s);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, s);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, s);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeStart cxCis:
						v.VisitCustomXmlConflictInsertionRangeStart(cxCis, s);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeEnd cxCie:
						v.VisitCustomXmlConflictInsertionRangeEnd(cxCie, s);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeStart cxCds:
						v.VisitCustomXmlConflictDeletionRangeStart(cxCds, s);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeEnd cxCde:
						v.VisitCustomXmlConflictDeletionRangeEnd(cxCde, s);
						break;

					// ---- Tracked-change run containers ----
					case InsertedRun insRun:
						WalkInsertedRun(insRun, s, v);
						break;
					case DeletedRun delRun:
						WalkDeletedRun(delRun, s, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, s);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, s);
						break;

					// ---- Office Math (inline & display) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, s);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, s);
						break;
						// When fine-grained m:* nodes are exposed elsewhere, forward them to v.VisitOMathElement here.

					// ---- Bidi & subdocument ----
					case BidirectionalOverride bdo:
						WalkBidirectionalOverride(bdo, s, v);
						break;
					case BidirectionalEmbedding bdi:
						WalkBidirectionalEmbedding(bdi, s, v);
						break;
					case SubDocumentReference subDoc:
						v.VisitSubDocumentReference(subDoc, s);
						break;

					// ---- ContentPart (Office 2010) ----
					case ContentPart cp:
						v.VisitContentPart(cp, s);
						break;

					// ---- Markup Compatibility wrapper (can appear anywhere) ----
					case AlternateContent ac:
						WalkAlternateContent(ac, s, v);
						break;

					default:
						ForwardUnknown("Hyperlink", child, s, v);
						break;
				}
			}

			// Ensure inline styles are closed before closing the anchor so tag order stays valid.
			_style.ResetStyle(v);
		}
	}

	private string? ResolveHyperlinkTarget(Hyperlink link)
	{
		// Anchor links are direct
		if (!string.IsNullOrEmpty(link.Anchor?.Value))
		{
			_referencedAnchors.Add(link.Anchor!.Value!);
			return "#" + link.Anchor!.Value;
		}

		var relId = link.Id?.Value;
		var part = _currentPart ?? _main;
		if (string.IsNullOrEmpty(relId) || part == null)
			return null;

		var rel = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
		if (rel == null)
			return null;

		return rel.Uri?.ToString();
	}

	// Helper: split the whitespace-delimited Requires value into prefixes.
	private static IReadOnlyList<string> GetRequiredPrefixes(AlternateContentChoice ch)
	{
		string? val = ch.Requires?.Value; // <-- correct source for mc:Requires
		if (string.IsNullOrWhiteSpace(val))
			return Array.Empty<string>();

		char[]? NullSeparator = null!;
		return val!.Split(NullSeparator, StringSplitOptions.RemoveEmptyEntries);
	}

	// Visitor asks: “Do I support this Choice?” If yes, we process it and stop.
	// If none are accepted, we process Fallback (if present). Else drop content per MCE.
	private void WalkAlternateContent(AlternateContent ac, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (v.VisitAlternateContentBegin(ac, s))
		{
			foreach (var choice in ac.Elements<AlternateContentChoice>())
			{
				var required = GetRequiredPrefixes(choice); // eg: ["w14","wps"]
															// The visitor should return true ONLY if it supports every required namespace.
				if (v.AcceptAlternateContentChoice(choice, required, s))
				{
					WalkAlternateContentSelectedContainer(choice, s, v);
					return;
				}
			}

			var fallback = ac.Elements<AlternateContentFallback>().FirstOrDefault();
			if (fallback != null)
			{
				WalkAlternateContentSelectedContainer(fallback, s, v);
				return;
			}

			// No accepted Choice and no Fallback => content is ignored (as if it didn't exist).
			// This matches the MCE preprocessing model.
		}
	}

	private void WalkAlternateContentSelectedContainer(OpenXmlElement container, IDxpStyleResolver s, IDxpVisitor v)
	{
		foreach (var child in container.ChildElements)
		{
			switch (child)
			{
				// Block-level
				case Paragraph:
				case Table:
				case BookmarkStart:
				case BookmarkEnd:
				case SdtBlock:
				case CustomXmlBlock:
				case AltChunk:
					WalkBlock(child, s, v);
					break;

				// Inline (common)
				case Run r:
					WalkRun(r, s, v);
					break;
				case Hyperlink link:
					WalkHyperlink(link, s, v);
					break;
				case SdtRun sdtRun:
					WalkSdtRun(sdtRun, s, v);
					break;
				case SimpleField fld:
					WalkSimpleField(fld, s, v);
					break;

				// Drawings / VML textboxes
				case Drawing d:
					TryWalkDrawingTextBox(d, s, v);
					break;
				case Picture pict:
					TryWalkVmlTextBox(pict, s, v);
					break;

				// Other allowed inline containers
				case ContentPart cp:
					v.VisitContentPart(cp, s);
					break;
				case DocumentFormat.OpenXml.Math.OfficeMath oMath:
					v.VisitOMath(oMath, s);
					break;
				case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
					v.VisitOMathParagraph(oMathPara, s);
					break;

				// Nested MC
				case AlternateContent nested:
					WalkAlternateContent(nested, s, v);
					break;

				default:
					ForwardUnknown("AlternateContent selected branch", child, s, v);
					break;
			}
		}
	}

	private void ForwardUnknown(string context, OpenXmlElement el, IDxpStyleResolver s, IDxpVisitor v)
	{
		v.VisitUnknown(context, el, s);
	}

	private void TryWalkDrawingTextBox(Drawing d, IDxpStyleResolver s, IDxpVisitor v)
	{
		var info = _drawings.TryResolveDrawingInfo(_currentPart ?? _main, d);
		using (v.VisitDrawingBegin(d, info, s))
		{
			// Look for Office 2010 Wordprocessing shape textbox: <wps:txbx>
			// SDK types live under DocumentFormat.OpenXml.Office2010.Word.DrawingShape
			var txbx = d
				.Descendants<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.TextBoxInfo2>() // wps:txbx
				.FirstOrDefault();
			if (txbx == null)
				return;

			var content = txbx.GetFirstChild<TextBoxContent>(); // w:txbxContent
			if (content == null)
				return;

			WalkTextBoxContent(content, s, v);
		}
	}

	private void TryWalkVmlTextBox(Picture pict, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (v.VisitLegacyPictureBegin(pict, s))
		{
			// VML textbox under w:pict/v:textbox
			var vmlTxbx = pict
				.Descendants<DocumentFormat.OpenXml.Vml.TextBox>() // v:textbox
				.FirstOrDefault();
			if (vmlTxbx == null)
				return;

			var content = vmlTxbx.GetFirstChild<TextBoxContent>(); // w:txbxContent
			if (content == null)
				return;

			WalkTextBoxContent(content, s, v);
		}
	}

	private void WalkTextBoxContent(TextBoxContent txbx, IDxpStyleResolver s, IDxpVisitor v)
	{
		// Optional: let the visitor know we’re entering a text box body
		using (v.VisitTextBoxContentBegin(txbx, s))
		{
			foreach (var child in txbx.ChildElements)
			{
				switch (child)
				{
					case Paragraph p:
					case Table t:
						case SdtBlock sdt:
						case CustomXmlBlock cx:
						case AltChunk ac:
							WalkBlock(child, s, v); // reuse the normal block dispatcher
							break;

					// Math is also allowed here per SDK child list for TextBoxContent
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, s);
						break;
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, s);
						break;

					default:
						ForwardUnknown("TextBoxContent", child, s, v);
						break;
				}
			}
		}
	}

	private void WalkRun(Run r, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (v.VisitRunBegin(r, s))
		{
			// Resolve style only if we can find a paragraph context.
			var para = r.Ancestors<Paragraph>().FirstOrDefault();
			bool hasRenderable = r.ChildElements.Any(child =>
				child is Text or DeletedText or NoBreakHyphen or TabChar or Break or CarriageReturn or Drawing);
			if (para != null)
			{
				if (hasRenderable)
				{
					DxpStyleEffectiveRunStyle style = s.ResolveRunStyle(para, r);
					_style.ApplyStyle(style, v);
				}
			}
			else
			{
				// No paragraph ancestor — surface to the visitor but keep walking content.
				ForwardUnknown("Run (no Paragraph ancestor)", r, s, v);
			}

			foreach (var child in r.ChildElements)
			{
				switch (child)
				{
					case NoBreakHyphen h:
						v.VisitNoBreakHyphen(h, s);
						break;

					case LastRenderedPageBreak pb:
						v.VisitLastRenderedPageBreak(pb, s);
						break;

					case RunProperties rp:
						v.VisitRunProperties(rp, s);
						break;

					case DeletedText dt:
						v.VisitDeletedText(dt, s);
						break;

					case Text t:
						v.VisitText(t, s);
						break;

					case TabChar tab:
						v.VisitTab(tab, s);
						break;

					case Break br:
						v.VisitBreak(br, s);
						break;

					case CarriageReturn cr:
						v.VisitCarriageReturn(cr, s);
						break;

					case Drawing d:
						TryWalkDrawingTextBox(d, s, v);
						break;

					case FieldChar fc:
					{
						var t = fc.FieldCharType?.Value;

						if (t == FieldCharValues.Begin)
						{
							v.VisitComplexFieldBegin(fc, s);
							_fieldStack.Push(new FieldFrame { SeenSeparate = false, ResultScope = null });
						}
						else if (t == FieldCharValues.Separate)
						{
							if (_fieldStack.Count > 0)
							{
								var top = _fieldStack.Pop();
								if (!top.SeenSeparate)
								{
									v.VisitComplexFieldSeparate(fc, s);
									top.SeenSeparate = true;
									if (top.ResultScope == null)
										top.ResultScope = v.VisitComplexFieldResultBegin(s);
								}
								_fieldStack.Push(top);
							}
							else
							{
								// stray separate; surface but don’t crash
								v.VisitComplexFieldSeparate(fc, s);
							}
						}
						else if (t == FieldCharValues.End)
						{
							if (_fieldStack.Count > 0)
							{
								var top = _fieldStack.Pop();
								top.ResultScope?.Dispose();
								v.VisitComplexFieldEnd(fc, s);
							}
							else
							{
								// stray end; surface but don’t crash
								v.VisitComplexFieldEnd(fc, s);
							}
						}
						// Other FieldChar types (rare) — ignore.
						break;
					}

					case FieldCode code:
					{
						// FieldCode.Text can be null; InnerText is a safe fallback
						var instr = code.Text ?? code.InnerText ?? string.Empty;
						v.VisitComplexFieldInstruction(code, instr, s);
						// Do not emit as visible text; instruction is not the result.
						break;
					}

					case FootnoteReference fr:
					{
						long fnId = fr.Id?.Value ?? 0;
						if (_footnotes.Resolve(fnId, out int fnIndex))
							v.VisitFootnoteReference(fr, fnId, fnIndex, s);
						break;
					}

					case CommentReference cref:
					{
						string id = cref.Id?.Value ?? string.Empty;
						TryEmitInlineComment(id, s, v);
						break;
					}

					case AlternateContent ac:
						WalkAlternateContent(ac, s, v);
						break;

					// Legacy DATE/PAGENUM-style blocks (non-editable placeholders)
					case DayShort ds:
						v.VisitDayShort(ds, s);
						break;
					case MonthShort ms:
						v.VisitMonthShort(ms, s);
						break;
					case YearShort ys:
						v.VisitYearShort(ys, s);
						break;
					case DayLong dl:
						v.VisitDayLong(dl, s);
						break;
					case MonthLong ml:
						v.VisitMonthLong(ml, s);
						break;
					case YearLong yl:
						v.VisitYearLong(yl, s);
						break;
					case PageNumber pn:
						v.VisitPageNumber(pn, s);
						break;

					// Marks and references (footnotes/endnotes/annotations/separators)
					case AnnotationReferenceMark arm:
						v.VisitAnnotationReference(arm, s);
						break;
					case FootnoteReferenceMark frm:
						if (_footnotes.Resolve(_CurrentFootnoteId ?? 0, out int index))
							v.VisitFootnoteReferenceMark(frm, _CurrentFootnoteId, index, s);
						break;
					case EndnoteReferenceMark erm:
						v.VisitEndnoteReferenceMark(erm, s);
						break;
					case EndnoteReference enr:
						v.VisitEndnoteReference(enr, s);
						break;
					case SeparatorMark sep:
						v.VisitSeparatorMark(sep, s);
						break;
					case ContinuationSeparatorMark csep:
						v.VisitContinuationSeparatorMark(csep, s);
						break;

					// Characters / inline controls
					case SoftHyphen sh:
						v.VisitSoftHyphen(sh, s);
						break;
					case SymbolChar sym:
						v.VisitSymbol(sym, s);
						break;
					case PositionalTab ptab:
						v.VisitPositionalTab(ptab, s);
						break;
					case Ruby ruby:
						WalkRuby(ruby, s, v);
						break;

					// Fields (deleted instruction text)
					case DeletedFieldCode dfc:
						v.VisitDeletedFieldCode(dfc, s);
						break;

					/* Legacy/object content within run */
					case EmbeddedObject obj:
						v.VisitEmbeddedObject(obj, s);
						break;
					case Picture pict:
						TryWalkVmlTextBox(pict, s, v);
						break;

					case ContentPart cp:
						v.VisitContentPart(cp, s);
						break;

					default:
						ForwardUnknown("Run child", child, s, v);
						break;
				}
			}
		}
	}


	// Walk a ruby (phonetic guide) inline container: <w:ruby> -> <w:rubyPr>, <w:rt>, <w:rubyBase>
	private void WalkRuby(Ruby ruby, IDxpStyleResolver s, IDxpVisitor v)
	{
		// Begin ruby scope (visitor can choose how to render a ruby run)
		using (v.VisitRubyBegin(ruby, s))
		{
			// --- Properties: <w:rubyPr> controls alignment/size/raise/lang of the ruby text ---
				var pr = ruby.GetFirstChild<RubyProperties>(); // SDK class for <w:rubyPr>
				if (pr != null)
					v.VisitRubyProperties(pr, s); // exposes rubyAlign, hps, hpsRaise, hpsBaseText, lid, dirty
												  // Spec: rubyPr is required. We log if missing, but don’t throw to be resilient.
												  // (Phonetic Guide Properties per CT_RubyPr.)  // ISO/IEC 29500 Part 1: w:rubyPr.

				// --- Ruby text: <w:rt> holds phonetic text in a required single <w:r> ---
				RubyContent? rt = ruby.GetFirstChild<RubyContent>(); // SDK: RubyContent for <w:rt>
				if (rt != null)
					WalkRubyContent(rt, isBase: false, s, v); // CT_RubyContent (phonetic guide text).

				// --- Base text: <w:rubyBase> holds the base characters in a required single <w:r> ---
				RubyBase? rb = ruby.GetFirstChild<RubyBase>();    // SDK: RubyBase for <w:rubyBase>
				if (rb != null)
					WalkRubyContent(rb, isBase: true, s, v); // CT_RubyContent (base text).
			}
		}

		// Walk CT_RubyContent (<w:rt> or <w:rubyBase>): spec says it MUST contain exactly one <w:r>,
		// but can also include select inline/range markup (proofErr, bookmarks, math, etc.).
	private void WalkRubyContent(RubyContentType rc, bool isBase, IDxpStyleResolver s, IDxpVisitor v)
	{
		// Allow the visitor to differentiate ruby text vs base text if useful
		using (v.VisitRubyContentBegin(rc, isBase, s))
		{
			foreach (var child in rc.ChildElements)
			{
				switch (child)
				{
						case Run r:
							WalkRun(r, s, v); // reuse the existing run walker
						break;

					// Inline/range markup permitted by CT_RubyContent — forward to existing visitor hooks:
					case ProofError pe:
						v.VisitProofError(pe, s);
					break;

				case BookmarkStart bs:
					v.VisitBookmarkStart(bs, s);
					break;
				case BookmarkEnd be:
					v.VisitBookmarkEnd(be, s);
					break;

					case PermStart ps:
						v.VisitPermStart(ps, s);
						break;
					case PermEnd pe2:
						v.VisitPermEnd(pe2, s);
						break;

					// Tracked move/ins/del regions allowed here (container-level markers):
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, s);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, s);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, s);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, s);
						break;

					case Inserted ins:
						WalkInserted(ins, s, v);
						break;
					case Deleted del:
						WalkDeleted(del, s, v);
						break;

					// Office Math is explicitly allowed here:
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, s);
						break;
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, s);
						break;

					// Compatibility: if spec groups (EG_RunLevelElts / EG_RangeMarkup / EG_MathContent)
					// expose additional children via SDK updates, route them or throw here.

					default:
						ForwardUnknown(isBase ? "RubyBase" : "RubyText", child, s, v);
						break;
				}
			}
		}
	}

	// CT_TrackChange leaf: w:del under trPr = deleted table row,
	// w:del under pPr/rPr = deleted paragraph mark. No children to walk.
	private void WalkDeleted(Deleted del, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (v.VisitDeletedBegin(del, s))
		{

			// Case 1: table row deletion (w:trPr/w:del)
			if (del.Parent is TableRowProperties trPr)
			{
				var tr = trPr.Parent as TableRow; // usually non-null when in a live tree
				v.VisitDeletedTableRowMark(del, trPr, tr, s); // tell visitor: the row is marked deleted
				return;
			}

			// Case 2: paragraph mark deletion (w:pPr/w:rPr/w:del)
			// This marks the *paragraph mark* deleted (contents merged with next para per spec).
			if (del.Parent is RunProperties rPr && rPr.Parent is ParagraphProperties pPr)
			{
				var p = pPr.Parent as Paragraph;
				v.VisitDeletedParagraphMark(del, pPr, p, s); // tell visitor: paragraph mark is deleted
				return;
			}

			// Anything else is unexpected for w:del (schema lists trPr and rPr parents).
			ForwardUnknown("Deleted", del, s, v);
			return;
		}
	}


	// CT_TrackChange leaf: w:ins marks an insertion on (1) a table row, (2) paragraph numbering props, or (3) the paragraph mark.
	// It has no children; we only need to determine the parent scope and notify the visitor.  Spec: w:ins under trPr / numPr / (pPr/rPr).
	private void WalkInserted(Inserted ins, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (v.VisitInsertedBegin(ins, s))
		{
			// (1) Inserted table row: <w:trPr><w:ins .../></w:trPr>  ⇒ the row itself is marked as inserted.
			// Ref: "Inserted Table Row" section + example under Parent Elements trPr. 
			if (ins.Parent is TableRowProperties trPr)
			{
				var tr = trPr.Parent as TableRow;
				v.VisitInsertedTableRowMark(ins, trPr, tr, s);
				return;
			}

			// (2) Inserted numbering properties: <w:pPr><w:numPr>...<w:ins .../></w:numPr></w:pPr>
			// Ref: "Inserted Numbering Properties" section; parent element listed as numPr.
			if (ins.Parent is NumberingProperties numPr)
			{
				var pPr = numPr.Parent as ParagraphProperties;
				var p = pPr?.Parent as Paragraph;
				v.VisitInsertedNumberingProperties(ins, numPr, pPr, p, s);
				return;
			}

			// (3) Inserted paragraph mark: <w:pPr><w:rPr><w:ins .../></w:rPr></w:pPr>
			// Ref: "Inserted Paragraph" section; parent elements listed as rPr (within pPr).
			if (ins.Parent is RunProperties rPr && rPr.Parent is ParagraphProperties pPr2)
			{
				var p = pPr2.Parent as Paragraph;
				v.VisitInsertedParagraphMark(ins, pPr2, p, s);
				return;
			}

			// Any other scope would be unexpected for w:ins per the schema; keep strict to surface it early.
			ForwardUnknown("Inserted", ins, s, v);
			return;
		}
	}

	private void WalkEndnote(Endnote fn, int index, long id, IDxpStyleResolver s, IDxpVisitor v)
	{
		_CurrentFootnoteId = id;
		try
		{
			using (PushCurrentPart(_main?.EndnotesPart))
			using (v.VisitEndnoteBegin(fn, id, index, s))
			{
				foreach (var child in fn.ChildElements)
				{
					switch (child)
					{
						// Block content
						case Paragraph p:
							WalkBlock(p, s, v);
							break;
						case Table t:
							WalkBlock(t, s, v);
							break;
						case SdtBlock sdt:
							WalkSdtBlock(sdt, s, v);
							break;
						case CustomXmlBlock cx:
							WalkCustomXmlBlock(cx, s, v);
							break;
						case AltChunk ac:
							v.VisitAltChunk(ac, s);
							break;

						// Math
						case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
							v.VisitOMathParagraph(oMathPara, s);
							break;
						case DocumentFormat.OpenXml.Math.OfficeMath oMath:
							v.VisitOMath(oMath, s);
							break;

						// Anchors
						case BookmarkStart bs:
							v.VisitBookmarkStart(bs, s);
							break;
						case BookmarkEnd be:
							v.VisitBookmarkEnd(be, s);
							break;
						case CommentRangeStart crs:
							v.VisitCommentRangeStart(crs, s);
							break;
						case CommentRangeEnd cre:
							break;

						// Permissions & proofing
						case PermStart ps:
							v.VisitPermStart(ps, s);
							break;
						case PermEnd pe:
							v.VisitPermEnd(pe, s);
							break;
						case ProofError pr:
							v.VisitProofError(pr, s);
							break;

						// Tracked ranges
						case Inserted ins:
							WalkInserted(ins, s, v);
							break;
						case Deleted del:
							WalkDeleted(del, s, v);
							break;
						case MoveFromRangeStart mfrs:
							v.VisitMoveFromRangeStart(mfrs, s);
							break;
						case MoveFromRangeEnd mfre:
							v.VisitMoveFromRangeEnd(mfre, s);
							break;
						case MoveToRangeStart mtrs:
							v.VisitMoveToRangeStart(mtrs, s);
							break;
						case MoveToRangeEnd mtre:
							v.VisitMoveToRangeEnd(mtre, s);
							break;
						case CustomXmlInsRangeStart cins:
							v.VisitCustomXmlInsRangeStart(cins, s);
							break;
						case CustomXmlInsRangeEnd cine:
							v.VisitCustomXmlInsRangeEnd(cine, s);
							break;
						case CustomXmlDelRangeStart cdls:
							v.VisitCustomXmlDelRangeStart(cdls, s);
							break;
						case CustomXmlDelRangeEnd cdle:
							v.VisitCustomXmlDelRangeEnd(cdle, s);
							break;
						case CustomXmlMoveFromRangeStart cmfs:
							v.VisitCustomXmlMoveFromRangeStart(cmfs, s);
							break;
						case CustomXmlMoveFromRangeEnd cmfe:
							v.VisitCustomXmlMoveFromRangeEnd(cmfe, s);
							break;
						case CustomXmlMoveToRangeStart cmts:
							v.VisitCustomXmlMoveToRangeStart(cmts, s);
							break;
						case CustomXmlMoveToRangeEnd cmte:
							v.VisitCustomXmlMoveToRangeEnd(cmte, s);
							break;

						default:
							ForwardUnknown("Endnote", child, s, v);
							break;
					}
				}
				_style.ResetStyle(v);
			}
		}
		finally
		{
			_CurrentFootnoteId = null; // ensure no carryover
		}
	}



	private void WalkFootnote(Footnote fn, int index, long id, IDxpStyleResolver s, IDxpVisitor v)
	{
		_CurrentFootnoteId = id;
		try
		{
			using (PushCurrentPart(_main?.FootnotesPart))
			using (v.VisitFootnoteBegin(fn, id, index, s))
			{
				foreach (var child in fn.ChildElements)
				{
					switch (child)
					{
						// Block content
						case Paragraph p:
							WalkBlock(p, s, v);
							break;
						case Table t:
							WalkBlock(t, s, v);
							break;
						case SdtBlock sdt:
							WalkSdtBlock(sdt, s, v);
							break;
						case CustomXmlBlock cx:
							WalkCustomXmlBlock(cx, s, v);
							break;
						case AltChunk ac:
							v.VisitAltChunk(ac, s);
							break;

						// Math
						case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
							v.VisitOMathParagraph(oMathPara, s);
							break;
						case DocumentFormat.OpenXml.Math.OfficeMath oMath:
							v.VisitOMath(oMath, s);
							break;

						// Anchors
						case BookmarkStart bs:
							v.VisitBookmarkStart(bs, s);
							break;
						case BookmarkEnd be:
							v.VisitBookmarkEnd(be, s);
							break;
						case CommentRangeStart crs:
							v.VisitCommentRangeStart(crs, s);
							break;
						case CommentRangeEnd cre:
							break;

						// Permissions & proofing
						case PermStart ps:
							v.VisitPermStart(ps, s);
							break;
						case PermEnd pe:
							v.VisitPermEnd(pe, s);
							break;
						case ProofError pr:
							v.VisitProofError(pr, s);
							break;

						// Tracked ranges
						case Inserted ins:
							WalkInserted(ins, s, v);
							break;
						case Deleted del:
							WalkDeleted(del, s, v);
							break;
						case MoveFromRangeStart mfrs:
							v.VisitMoveFromRangeStart(mfrs, s);
							break;
						case MoveFromRangeEnd mfre:
							v.VisitMoveFromRangeEnd(mfre, s);
							break;
						case MoveToRangeStart mtrs:
							v.VisitMoveToRangeStart(mtrs, s);
							break;
						case MoveToRangeEnd mtre:
							v.VisitMoveToRangeEnd(mtre, s);
							break;
						case CustomXmlInsRangeStart cins:
							v.VisitCustomXmlInsRangeStart(cins, s);
							break;
						case CustomXmlInsRangeEnd cine:
							v.VisitCustomXmlInsRangeEnd(cine, s);
							break;
						case CustomXmlDelRangeStart cdls:
							v.VisitCustomXmlDelRangeStart(cdls, s);
							break;
						case CustomXmlDelRangeEnd cdle:
							v.VisitCustomXmlDelRangeEnd(cdle, s);
							break;
						case CustomXmlMoveFromRangeStart cmfs:
							v.VisitCustomXmlMoveFromRangeStart(cmfs, s);
							break;
						case CustomXmlMoveFromRangeEnd cmfe:
							v.VisitCustomXmlMoveFromRangeEnd(cmfe, s);
							break;
						case CustomXmlMoveToRangeStart cmts:
							v.VisitCustomXmlMoveToRangeStart(cmts, s);
							break;
						case CustomXmlMoveToRangeEnd cmte:
							v.VisitCustomXmlMoveToRangeEnd(cmte, s);
							break;

						default:
							ForwardUnknown("Footnote", child, s, v);
							break;
					}
				}
				_style.ResetStyle(v);
			}
		}
		finally
		{
			_CurrentFootnoteId = null; // ensure no carryover
		}
	}


	private void WalkCustomXmlBlock(CustomXmlBlock cx, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (v.VisitCustomXmlBlockBegin(cx, s))
		{
			// Properties (<w:customXmlPr>) carry element name/namespace and optional data binding metadata.
			var pr = cx.CustomXmlProperties;
			if (pr != null)
				v.VisitCustomXmlProperties(pr, s);

			foreach (var child in cx.ChildElements)
			{
				switch (child)
				{
					// ---- Block-level content allowed by CT_CustomXmlBlock ----
					case Paragraph p:
						WalkBlock(p, s, v);
						break;
					case Table t:
						WalkBlock(t, s, v);
						break;
					case SdtBlock sdt:
						WalkSdtBlock(sdt, s, v);
						break;
					case CustomXmlBlock nested:
						WalkCustomXmlBlock(nested, s, v);
						break;

					// ---- Office Math (explicitly permitted here) ----
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, s);
						break;
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, s);
						break;

					// ---- Bookmark/comment anchors (range markup) ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, s);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, s);
						break;
					case CommentRangeStart crs:
						v.VisitCommentRangeStart(crs, s);
						break;
					case CommentRangeEnd cre:
						break;

					// ---- Permissions & proofing anchors ----
					case PermStart ps:
						v.VisitPermStart(ps, s);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, s);
						break;
					case ProofError perr:
						v.VisitProofError(perr, s);
						break;

					// ---- Custom XML range markup (insert/delete/move; start/end) ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, s);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, s);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, s);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, s);
						break;
					case CustomXmlMoveFromRangeStart cxMfs:
						v.VisitCustomXmlMoveFromRangeStart(cxMfs, s);
						break;
					case CustomXmlMoveFromRangeEnd cxMfe:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfe, s);
						break;
					case CustomXmlMoveToRangeStart cxMts:
						v.VisitCustomXmlMoveToRangeStart(cxMts, s);
						break;
					case CustomXmlMoveToRangeEnd cxMte:
						v.VisitCustomXmlMoveToRangeEnd(cxMte, s);
						break;

					// ---- Office 2010 conflict range markup (optional) ----
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeStart cxCis:
						v.VisitCustomXmlConflictInsertionRangeStart(cxCis, s);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeEnd cxCie:
						v.VisitCustomXmlConflictInsertionRangeEnd(cxCie, s);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeStart cxCds:
						v.VisitCustomXmlConflictDeletionRangeStart(cxCds, s);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeEnd cxCde:
						v.VisitCustomXmlConflictDeletionRangeEnd(cxCde, s);
						break;

						// ---- Run-level change containers that are valid children here ----
						case InsertedRun insRun:
							WalkInsertedRun(insRun, s, v);
							break;
						case DeletedRun delRun:
							WalkDeletedRun(delRun, s, v);
							break;
						case MoveFromRun mfr:
							v.VisitMoveFromRun(mfr, s); // WalkMoveFromRun can be used if expanded
							break;
						case MoveToRun mtr:
							v.VisitMoveToRun(mtr, s);   // WalkMoveToRun can be used if expanded
							break;

					// ---- ContentPart (Office 2010) – embedded external content reference ----
					case ContentPart cp:
						v.VisitContentPart(cp, s);
						break;

					default:
						ForwardUnknown("CustomXmlBlock", child, s, v);
						break;
				}
			}
		}
	}


	private void WalkSdtBlock(SdtBlock sdt, IDxpStyleResolver s, IDxpVisitor v)
	{
		using (v.VisitSdtBlockBegin(sdt, s))
		{
			// (1) Properties (optional)
			var pr = sdt.GetFirstChild<SdtProperties>();
			if (pr != null)
				v.VisitSdtProperties(pr, s);

			// (2) Content (optional per schema — do NOT throw if missing)
			var content = sdt.SdtContentBlock;
			if (content == null)
			{
				// Empty SDT is valid: just end the SDT scope
				return;
			}

			using (v.VisitSdtContentBlockBegin(content, s))
			{
				foreach (var child in content.ChildElements)
				{
					// SDT block content is normal block content
					WalkBlock(child, s, v);
				}
			}
		}
	}


	private IDisposable PushCurrentPart(OpenXmlPart? part)
	{
		var previous = _currentPart;
		if (part != null)
			_currentPart = part;
		return Disposable.Create(() => _currentPart = previous);
	}


	// ---- Strict error reporting ----

	private static NotSupportedException Unsupported(string context, OpenXmlElement el)
	{
		// LocalName gives a stable XML-ish name; GetType gives the SDK type.
		var name = el.LocalName;
		var type = el.GetType().FullName ?? el.GetType().Name;

		// OuterXml can be huge; include a small snippet.
		var snippet = el.OuterXml;
		if (snippet.Length > 300)
			snippet = snippet.Substring(0, 300) + "…";

		return new NotSupportedException(
			$"Unsupported element in {context}: <{name}> ({type}). Snippet: {snippet}"
		);
	}
}
