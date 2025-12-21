using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using Microsoft.Extensions.Logging;
using System.Xml.Linq;

namespace DocxportNet.Walker;

public class DxpWalker
{
	private readonly ILogger? _logger;

	public DxpWalker(ILogger? logger = null)
	{
		_logger = logger;
	}

	public void Accept(string docxPath, DxpIVisitor v)
	{
		using var doc = WordprocessingDocument.Open(docxPath, false);
		Accept(doc, v);
	}

	public void Accept(WordprocessingDocument doc, DxpIVisitor v)
	{
		if (doc.MainDocumentPart == null)
			return;

		DxpDocumentContext documentContext = new DxpDocumentContext(doc);

		documentContext.MainDocumentPart = doc.MainDocumentPart;
		using (documentContext.PushCurrentPart(doc.MainDocumentPart))
		{

			var settings = doc.MainDocumentPart.DocumentSettingsPart?.Settings;
			if (settings != null)
				v.VisitDocumentSettings(settings, documentContext);

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
			WalkDocumentBody(body, documentContext, v);
			// Remove the global “walk all headers/footers” to avoid duplicates (see #2)

			// Footnotes/Endnotes
			foreach (var fn in documentContext.Footnotes.GetFootnotes())
				WalkFootnote(fn.Item1, fn.Item2, fn.Item3, documentContext, v);
			foreach (var en in documentContext.Endnotes.GetEndnotes())
				WalkEndnote(en.Item1, en.Item2, en.Item3, documentContext, v);

			WalkBibliography(doc, documentContext, v);

			// Global cleanup: dispose any unterminated complex field result scopes
			if (documentContext.CurrentFields.FieldStack.Count > 0)
			{
				int leaked = documentContext.CurrentFields.FieldStack.Count;
				_logger?.LogWarning("Detected {LeakedFields} unterminated complex field(s) at end of document; disposing result scopes defensively.", leaked);
				while (documentContext.CurrentFields.FieldStack.Count > 0)
				{
					var frame = documentContext.CurrentFields.FieldStack.Pop();
					frame.ResultScope?.Dispose();
				}
			}

			documentContext.CurrentPart = null;
		}
	}

	private void WalkBibliography(WordprocessingDocument doc, DxpDocumentContext d, DxpIVisitor v)
	{
		// There is no MainDocumentPart.BibliographyPart property in the SDK.
		// Fetch the part via the container API.
		CustomXmlPart? bibPart = doc.MainDocumentPart?
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

	private void WalkDocumentBody(Body body, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitDocumentBodyBegin(body, d))
		{
			// TODO - we expect to always have at least 1 section, if the document
			// has none, expect BuildSections to create a default one like Word would.
			List<SectionSlice> sections = DxpSections.SplitDocumentBodyIntoSections(body);

			foreach (SectionSlice section in sections)
			{
				WalkSection(section, d, v);
			}
		}
	}

	private void WalkSection(SectionSlice section, DxpDocumentContext d, DxpIVisitor v)
	{
		SectionLayout layout = DxpSections.CreateSectionLayout(section.Properties);
		d.EnterSection(section.Properties, layout);

		using (v.VisitSectionBegin(section.Properties, layout, d))
		{
			// header
			HeaderReference? headerRef = DxpSections.FindFirstSectionHeaderReference(section.Properties);
			if (headerRef != null)
				WalkHeaderReference(headerRef, d, v);

			// body
			WalkSectionBody(section, d, v);

			// footer
			FooterReference? footerRef = DxpSections.FindLastSectionFooterReference(section.Properties);
			if (footerRef != null)
				WalkFooterReference(footerRef, d, v);
		}
	}

	private void WalkSectionBody(SectionSlice section, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitSectionBodyBegin(section.Properties, d))
		{
			foreach (var child in section.Blocks)
			{
				WalkBlock(child, d, v);
			}
		}
	}

	private void WalkBlock(OpenXmlElement block, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitBlockBegin(block, d))
		{
			switch (block)
			{
				case Paragraph p:
					WalkParagraph(p, d, v);
					break;

				case Table t:
					WalkTable(t, d, v);
					break;

				case BookmarkStart bs:
				{
					var name = bs.Name?.Value;
					v.VisitBookmarkStart(bs, d);
					d.StyleTracker.ResetStyle(d, v);
					return;
				}
				case BookmarkEnd be:
					v.VisitBookmarkEnd(be, d);
					d.StyleTracker.ResetStyle(d, v);
					return;

				case SectionProperties sp:
					v.VisitSectionProperties(sp, d);
					break;

				case SdtBlock sdt:
					WalkSdtBlock(sdt, d, v);
					break;

				case CustomXmlBlock cx:
					WalkCustomXmlBlock(cx, d, v);
					break;

				case AltChunk ac:
					v.VisitAltChunk(ac, d);
					break;

				case ContentPart cp:
					v.VisitContentPart(cp, d);
					break;

				// Anchors / permissions / proofing
				case CommentRangeStart crs:
					WalkCommentRangeStart(crs, d, v);
					break;
				case CommentReference cref:
					WalkCommentReference(cref, d, v);
					break;
				case CommentRangeEnd:
					// nothing to do for inline-at-start policy
					break;
				case PermStart ps:
					v.VisitPermStart(ps, d);
					break;
				case PermEnd pe:
					v.VisitPermEnd(pe, d);
					break;
				case ProofError per:
					v.VisitProofError(per, d);
					break;

				// Tracked-change markers at block scope
				case Inserted ins:
					WalkInserted(ins, d, v);
					break;
				case Deleted del:
					WalkDeleted(del, d, v);
					break;

				// Move ranges (location containers)
				case MoveFromRangeStart mfrs:
					v.VisitMoveFromRangeStart(mfrs, d);
					break;
				case MoveFromRangeEnd mfre:
					v.VisitMoveFromRangeEnd(mfre, d);
					break;
				case MoveToRangeStart mtrs:
					v.VisitMoveToRangeStart(mtrs, d);
					break;
				case MoveToRangeEnd mtre:
					v.VisitMoveToRangeEnd(mtre, d);
					break;

				// customXml range markup (start/end) + Office 2010 conflict ranges
				case CustomXmlInsRangeStart cxInsS:
					v.VisitCustomXmlInsRangeStart(cxInsS, d);
					break;
				case CustomXmlInsRangeEnd cxInsE:
					v.VisitCustomXmlInsRangeEnd(cxInsE, d);
					break;
				case CustomXmlDelRangeStart cxDelS:
					v.VisitCustomXmlDelRangeStart(cxDelS, d);
					break;
				case CustomXmlDelRangeEnd cxDelE:
					v.VisitCustomXmlDelRangeEnd(cxDelE, d);
					break;
				case CustomXmlMoveFromRangeStart cxMfS:
					v.VisitCustomXmlMoveFromRangeStart(cxMfS, d);
					break;
				case CustomXmlMoveFromRangeEnd cxMfE:
					v.VisitCustomXmlMoveFromRangeEnd(cxMfE, d);
					break;
				case CustomXmlMoveToRangeStart cxMtS:
					v.VisitCustomXmlMoveToRangeStart(cxMtS, d);
					break;
				case CustomXmlMoveToRangeEnd cxMtE:
					v.VisitCustomXmlMoveToRangeEnd(cxMtE, d);
					break;

				case DocumentFormat.OpenXml.Office2010.Word.ConflictInsertion cIns:
					v.VisitConflictInsertion(cIns, d);
					break; // w14:conflictIns
				case DocumentFormat.OpenXml.Office2010.Word.ConflictDeletion cDel:
					v.VisitConflictDeletion(cDel, d);
					break; // w14:conflictDel
				case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeStart cxCis:
					v.VisitCustomXmlConflictInsertionRangeStart(cxCis, d);
					break;
				case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeEnd cxCie:
					v.VisitCustomXmlConflictInsertionRangeEnd(cxCie, d);
					break;
				case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeStart cxCds:
					v.VisitCustomXmlConflictDeletionRangeStart(cxCds, d);
					break;
				case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeEnd cxCde:
					v.VisitCustomXmlConflictDeletionRangeEnd(cxCde, d);
					break;

				default:
					WalkUnknown("Block", block, d, v);
					break;
			}
		}
	}

	private void WalkHeaderReference(HeaderReference hr, DxpDocumentContext d, DxpIVisitor v)
	{
		var relId = hr.Id?.Value;
		if (string.IsNullOrEmpty(relId))
			return;

		if (d.MainDocumentPart?.GetPartById(relId!) is HeaderPart part && part.Header is Header hdr)
		{
			var kind = hr.Type?.Value ?? HeaderFooterValues.Default;
			using (d.PushCurrentPart(part))
			using (v.VisitSectionHeaderBegin(hdr, kind, d))
			{
				foreach (var child in hdr.ChildElements)
					WalkBlock(child, d, v);
			}
			d.StyleTracker.ResetStyle(d, v);
		}
	}

	private void WalkFooterReference(FooterReference fr, DxpDocumentContext d, DxpIVisitor v)
	{
		var relId = fr.Id?.Value;
		if (string.IsNullOrEmpty(relId))
			return;

		if (d.MainDocumentPart?.GetPartById(relId!) is FooterPart part && part.Footer is Footer ftr)
		{
			var kind = fr.Type?.Value ?? HeaderFooterValues.Default;
			using (d.PushCurrentPart(part))
			using (v.VisitSectionFooterBegin(ftr, kind, d))
			{
				foreach (var child in ftr.ChildElements)
					WalkBlock(child, d, v);
			}
			d.StyleTracker.ResetStyle(d, v);
		}
	}

	private void WalkTable(Table tbl, DxpDocumentContext d, DxpIVisitor v)
	{

		DxpTableModel model = d.Tables.BuildTableModel(tbl);
		var tableContext = new DxpTableContext(tbl, tbl.GetFirstChild<TableProperties>(), tbl.GetFirstChild<TableGrid>());
		int rowIndex = 0;

		using (v.VisitTableBegin(tbl, model, d, tableContext))
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
						tableContext.SetGrid(grid);
						v.VisitTableGrid(grid, d); // columns & default widths
						break;

					case TableRow tr:
					{
						int currentRow = rowIndex++;
						WalkTableRow(tr, d, v, tableContext, currentRow);
						break;
					}

					case SdtRow sdtRow:
						using (v.VisitSdtRowBegin(sdtRow, d))
						{
							var content = sdtRow.SdtContentRow;
							if (content != null)
							{
								foreach (var inner in content.Elements<TableRow>())
								{
									int currentRow = rowIndex++;
									WalkTableRow(inner, d, v, tableContext, currentRow);
								}
							}
						}
						break;

					case CustomXmlRow cxRow:
						using (v.VisitCustomXmlRowBegin(cxRow, d))
						{
							foreach (var inner in cxRow.Elements<TableRow>())
							{
								int currentRow = rowIndex++;
								WalkTableRow(inner, d, v, tableContext, currentRow);
							}
						}
						break;

					// Anchors / tracked ranges allowed under w:tbl
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;
					case ProofError perr:
						v.VisitProofError(perr, d);
						break;
					case InsertedRun ins:
						WalkInsertedRun(ins, d, v);
						break;
					case DeletedRun del:
						WalkDeletedRun(del, d, v);
						break;
					case MoveFromRun mfr:
						v.VisitMoveFromRun(mfr, d);
						break;
					case MoveToRun mtr:
						v.VisitMoveToRun(mtr, d);
						break;

						// customXml range start/end + Office 2010 conflict ranges also valid here
						// (handled the same way as other dispatchers)

					// Rare but permitted under w:tbl (SDK lists these)
					case DocumentFormat.OpenXml.Math.OfficeMath m:
						v.VisitOMath(m, d);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph mp:
						v.VisitOMathParagraph(mp, d);
						break;

					default:
						WalkUnknown("Table", child, d, v);
						break;
				}
			}
			d.StyleTracker.ResetStyle(d, v);
		}
	}

	private void WalkTableRow(TableRow tr, DxpDocumentContext d, DxpIVisitor v, DxpTableContext tableContext, int rowIndex)
	{
		bool isHeader = tr.TableRowProperties?.GetFirstChild<TableHeader>() != null;
		var rowContext = new DxpTableRowContext(tableContext, rowIndex, isHeader);

		using (v.VisitTableRowBegin(tr, rowContext, d))
		{
			int columnIndex = 0;
			foreach (var child in tr.ChildElements)
			{
				switch (child)
				{
					case TableCell tc:
						WalkTableCell(tc, d, v, rowContext, columnIndex);
						columnIndex++;
						break;

					case TableRowProperties trp:
						v.VisitTableRowProperties(trp, d);
						break;

					case SdtCell sdtCell:
					{
						using (v.VisitSdtCellBegin(sdtCell, d))
						{
							// <w:sdtCell><w:sdtContent><w:tc>…</w:tc></w:sdtContent></w:sdtCell>
							var content = sdtCell.SdtContentCell;
							if (content != null)
							{
								foreach (var inner in content.Elements<TableCell>())
								{
									WalkTableCell(inner, d, v, rowContext, columnIndex);
									columnIndex++;
								}
							}
						}
						break;
					}

					case CustomXmlCell cxCell:
					{
						using (v.VisitCustomXmlCellBegin(cxCell, d))
						{
							// <w:customXmlCell> may contain one or more <w:tc>
							foreach (var inner in cxCell.Elements<TableCell>())
							{
								WalkTableCell(inner, d, v, rowContext, columnIndex);
								columnIndex++;
							}
						}
						break;
					}

						// Keep other row-level cases (e.g., trPr) or forward unknowns:
					default:
						WalkUnknown("TableRow child", child, d, v);
						break;
				}
			}
		}
	}




	private void WalkTableCell(TableCell tc, DxpDocumentContext d, DxpIVisitor v, DxpTableRowContext rowContext, int columnIndex)
	{
		TableCellProperties? tcp = null;
		if (tc.HasChildren)
			tcp = tc.GetFirstChild<TableCellProperties>();

		var cellContext = new DxpTableCellContext(rowContext, rowContext.Index, columnIndex, 1, 1, tcp);
		using (v.VisitTableCellBegin(tc, cellContext, d))
		{

			bool sawBlock = false; // enforce CT_Tc rule: at least one block-level element

			foreach (var child in tc.ChildElements)
			{
				switch (child)
				{
					// ---- Block-level content inside a cell (EG_BlockLevelElts) ----
					case Paragraph p:
						sawBlock = true;
						WalkBlock(p, d, v);
						break;

					case Table t:
						sawBlock = true;
						WalkBlock(t, d, v);
						break;

					case SdtBlock sdt:
						sawBlock = true;
						WalkSdtBlock(sdt, d, v);
						break;

					case CustomXmlBlock cx:
						sawBlock = true;
						WalkCustomXmlBlock(cx, d, v);
						break;

					case AltChunk ac:
						sawBlock = true;
						v.VisitAltChunk(ac, d);
						break;

					// ---- Math (allowed directly under tc) ----
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						sawBlock = true;
						v.VisitOMathParagraph(oMathPara, d);
						break;
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						sawBlock = true;
						v.VisitOMath(oMath, d);
						break;

					// ---- Anchors / permissions / proofing (directly allowed in tc) ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;
					case ProofError perr:
						v.VisitProofError(perr, d);
						break;

					// ---- Tracked-change containers and range start/end (directly allowed) ----
					case Inserted ins:
						WalkInserted(ins, d, v);
						break;
					case Deleted del:
						WalkDeleted(del, d, v);
						break;
					case MoveFromRun mfr:
						v.VisitMoveFromRun(mfr, d);
						break;
					case MoveToRun mtr:
						v.VisitMoveToRun(mtr, d);
						break;
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break;

					// ---- customXml range markup (start/end) ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, d);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, d);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, d);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, d);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, d);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, d);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, d);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, d);
						break;

					default:
						WalkUnknown("TableCell", child, d, v);
						break;
				}
			}

			if (!sawBlock)
				throw new InvalidOperationException("w:tc must contain at least one block-level element (p/tbl/sdt/customXml/altChunk/math).");
		}
	}


	private void WalkParagraph(Paragraph p, DxpDocumentContext d, DxpIVisitor v)
	{
		if (!DxpParagraphs.HasRenderableParagraphContent(p))
			return;

		using (d.PushParagraph(p, out DxpParagraphContext paragraphContext))
		using (v.VisitParagraphBegin(p, d, paragraphContext.Marker, paragraphContext.Indent))
		{
			foreach (var child in p.ChildElements)
			{
				switch (child)
				{
					case ProofError pe:
						v.VisitProofError(pe, d);
						break;

					case DeletedRun dr:
						WalkDeletedRun(dr, d, v);
						break;

					case InsertedRun ir:
						WalkInsertedRun(ir, d, v);
						break;

					case ParagraphProperties pp:
						v.VisitParagraphProperties(pp, d);
						break;

					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;

					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;

					case Run r:
						WalkRun(r, d, v);
						break;

					case Hyperlink link:
						WalkHyperlink(link, d, v);
						break;

					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd:
						// nothing to do for inline-at-start policy
						break;

					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, d, v);
						break; // w:customXml (inline custom XML container).

					case SimpleField fld:
						WalkSimpleField(fld, d, v);
						break; // w:fldSimple (simple field; contains runs/hyperlinks).

					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, d, v);
						break; // w:sdt (run-level SDT).

					case PermStart ps:
						v.VisitPermStart(ps, d);
						break; // w:permStart (editing permission range start).
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break; // w:permEnd (editing permission range end).

					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break; // w:moveFromRangeStart (tracked move-out start).
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break; // w:moveFromRangeEnd (tracked move-out end).
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break; // w:moveToRangeStart (tracked move-in start).
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break; // w:moveToRangeEnd (tracked move-in end).

					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, d);
						break; // w:customXmlInsRangeStart (customXml insert range start).
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, d);
						break; // w:customXmlInsRangeEnd (customXml insert range end).
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, d);
						break; // w:customXmlDelRangeStart (customXml delete range start).
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, d);
						break; // w:customXmlDelRangeEnd (customXml delete range end).
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, d);
						break; // w:customXmlMoveFromRangeStart (customXml move-from start).
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, d);
						break; // w:customXmlMoveFromRangeEnd (customXml move-from end).
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, d);
						break; // w:customXmlMoveToRangeStart (customXml move-to start).
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, d);
						break; // w:customXmlMoveToRangeEnd (customXml move-to end).

					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeStart cxCis:
						v.VisitCustomXmlConflictInsertionRangeStart(cxCis, d);
						break; // w14:customXmlConflictInsRangeStart (Office 2010 conflict insert start).
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeEnd cxCie:
						v.VisitCustomXmlConflictInsertionRangeEnd(cxCie, d);
						break; // w14:customXmlConflictInsRangeEnd (Office 2010 conflict insert end).
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeStart cxCds:
						v.VisitCustomXmlConflictDeletionRangeStart(cxCds, d);
						break; // w14:customXmlConflictDelRangeStart (Office 2010 conflict delete start).
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeEnd cxCde:
						v.VisitCustomXmlConflictDeletionRangeEnd(cxCde, d);
						break; // w14:customXmlConflictDelRangeEnd (Office 2010 conflict delete end).

					case MoveFromRun mfr:
						v.VisitMoveFromRun(mfr, d);
						break; // w:moveFrom (run container for moved-out text).
					case MoveToRun mtr:
						v.VisitMoveToRun(mtr, d);
						break; // w:moveTo (run container for moved-in text).

					case ContentPart cp:
						v.VisitContentPart(cp, d);
						break; // w:contentPart (external content reference; Office 2010+).

					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break; // m:oMathPara (display math paragraph).
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break; // m:oMath (inline math object).

					case DocumentFormat.OpenXml.Math.Accent mAccent:
						v.VisitOMathElement(mAccent, d);
						break; // m:acc (math accent – direct child allowed).
					case DocumentFormat.OpenXml.Math.Bar mBar:
						v.VisitOMathElement(mBar, d);
						break; // m:bar (math bar).
					case DocumentFormat.OpenXml.Math.Box mBox:
						v.VisitOMathElement(mBox, d);
						break; // m:box (math box).
					case DocumentFormat.OpenXml.Math.BorderBox mBorderBox:
						v.VisitOMathElement(mBorderBox, d);
						break; // m:borderBox (math border box).
					case DocumentFormat.OpenXml.Math.Delimiter mDelim:
						v.VisitOMathElement(mDelim, d);
						break; // m:d (delimiter).
					case DocumentFormat.OpenXml.Math.EquationArray mEqArr:
						v.VisitOMathElement(mEqArr, d);
						break; // m:eqArr (equation array).
					case DocumentFormat.OpenXml.Math.Fraction mFrac:
						v.VisitOMathElement(mFrac, d);
						break; // m:f (fraction).
					case DocumentFormat.OpenXml.Math.MathFunction mFunc:
						v.VisitOMathElement(mFunc, d);
						break; // m:func (function).
					case DocumentFormat.OpenXml.Math.GroupChar mGroupChr:
						v.VisitOMathElement(mGroupChr, d);
						break; // m:groupChr (group character).
					case DocumentFormat.OpenXml.Math.LimitLower mLimLow:
						v.VisitOMathElement(mLimLow, d);
						break; // m:limLow (lower limit).
					case DocumentFormat.OpenXml.Math.LimitUpper mLimUpp:
						v.VisitOMathElement(mLimUpp, d);
						break; // m:limUpp (upper limit).
					case DocumentFormat.OpenXml.Math.Matrix mMat:
						v.VisitOMathElement(mMat, d);
						break; // m:m (matrix).
					case DocumentFormat.OpenXml.Math.Nary mNary:
						v.VisitOMathElement(mNary, d);
						break; // m:nary (n-ary operator).
					case DocumentFormat.OpenXml.Math.Phantom mPhant:
						v.VisitOMathElement(mPhant, d);
						break; // m:phantom (phantom).
					case DocumentFormat.OpenXml.Math.Radical mRad:
						v.VisitOMathElement(mRad, d);
						break; // m:rad (radical).
					case DocumentFormat.OpenXml.Math.PreSubSuper mPreSubSup:
						v.VisitOMathElement(mPreSubSup, d);
						break; // m:preSubSup (presub/superscript).
					case DocumentFormat.OpenXml.Math.Subscript mSub:
						v.VisitOMathElement(mSub, d);
						break; // m:s (subscript).
					case DocumentFormat.OpenXml.Math.SubSuperscript mSubSup:
						v.VisitOMathElement(mSubSup, d);
						break; // m:sSub (sub-superscript).
					case DocumentFormat.OpenXml.Math.Superscript mSup:
						v.VisitOMathElement(mSup, d);
						break; // m:sup (superscript).
					case DocumentFormat.OpenXml.Math.Run mMathRun:
						v.VisitOMathRun(mMathRun, d);
						break; // m:r (math run).

					case BidirectionalOverride bdo:
						WalkBidirectionalOverride(bdo, d, v);
						break; // w:bdo (Bidi override; Office 2010+).
					case BidirectionalEmbedding bdi:
						WalkBidirectionalEmbedding(bdi, d, v);
						break; // w:dir (Bidi embedding; Office 2010+).

					case SubDocumentReference subDoc:
						v.VisitSubDocumentReference(subDoc, d);
						break; // w:subDoc (subdocument anchor).

					case AlternateContent ac:
						WalkAlternateContent(ac, d, v);
						break; // mc:AlternateContent (compat wrapper around inline kids). (MC allows it here)

					// Comment anchors / tracked change containers will show up as elements too.
					// These currently throw to surface unhandled cases.
					default:
						WalkUnknown("Paragraph", child, d, v);
						break;
				}
			}

			d.StyleTracker.ResetStyle(d, v);
		}
	}

	private void WalkBidirectionalEmbedding(BidirectionalEmbedding bdi, DxpDocumentContext d, DxpIVisitor v)
	{
		// w:dir (Bidirectional Embedding). Attribute w:val is the embedding direction (e.g., "rtl"/"ltr").
		// Children per CT_DirContentRun include inline content, math, tracked ranges, customXml ranges, and nesting. 
		// Ref: SDK "BidirectionalEmbedding" child list & examples. 
		using (v.VisitBidirectionalEmbeddingBegin(bdi, d))
		{
			foreach (var child in bdi.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content ----
					case Run r:
						WalkRun(r, d, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, d, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, d, v);
						break;
					case SimpleField fld:
						WalkSimpleField(fld, d, v);
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, d, v);
						break;

					// ---- Range anchors / bookmarks / comments / permissions ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;
					case ProofError perr:
						v.VisitProofError(perr, d);
						break;

					// ---- Move ranges (location containers) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break;

					// ---- customXml ranges (start/end) + Office 2010 conflict ranges ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, d);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, d);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, d);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, d);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, d);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, d);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, d);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, d);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeStart cxCis:
						v.VisitCustomXmlConflictInsertionRangeStart(cxCis, d);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeEnd cxCie:
						v.VisitCustomXmlConflictInsertionRangeEnd(cxCie, d);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeStart cxCds:
						v.VisitCustomXmlConflictDeletionRangeStart(cxCds, d);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeEnd cxCde:
						v.VisitCustomXmlConflictDeletionRangeEnd(cxCde, d);
						break;

					// ---- Tracked-change run containers ----
					case InsertedRun insRun:
						WalkInsertedRun(insRun, d, v);
						break;
					case DeletedRun delRun:
						WalkDeletedRun(delRun, d, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, d);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, d);
						break;

					// ---- Office Math (inline & display) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;

					// ---- Nesting: bdo/dir/subDoc are allowed inside w:dir ----
					case BidirectionalOverride bdo:
						WalkBidirectionalOverride(bdo, d, v);
						break;
					case BidirectionalEmbedding nestedDir:
						WalkBidirectionalEmbedding(nestedDir, d, v);
						break;
					case SubDocumentReference subDoc:
						v.VisitSubDocumentReference(subDoc, d);
						break;

					// ---- MC wrapper (commonly appears anywhere inline) ----
					case AlternateContent ac:
						WalkAlternateContent(ac, d, v);
						break;

					default:
						WalkUnknown("BidirectionalEmbedding", child, d, v);
						break;
				}
			}
		}
	}

	private void WalkBidirectionalOverride(BidirectionalOverride bdo, DxpDocumentContext d, DxpIVisitor v)
	{
		// w:bdo (BiDi override). Direction via w:val (e.g., "rtl"/"ltr").
		using (v.VisitBidirectionalOverrideBegin(bdo, d))
		{
			foreach (var child in bdo.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content ----
					case Run r:
						WalkRun(r, d, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, d, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, d, v);
						break;
					case SimpleField fld:
						WalkSimpleField(fld, d, v);
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, d, v);
						break;
					// Not found in SDK
					case OpenXmlUnknownElement smart
						when smart.LocalName == "smartTag" && smart.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main":
					{
						WalkSmartTagRun(smart, d, v);
						break;
					}

					// ---- Range anchors / bookmarks / comments / permissions ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;
					case ProofError perr:
						v.VisitProofError(perr, d);
						break;

					// ---- Move locations (range start/end) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break;

					// ---- customXml ranges (start/end) ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, d);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, d);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, d);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, d);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, d);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, d);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, d);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, d);
						break;

					// ---- Tracked-change run containers ----
					case InsertedRun insRun:
						WalkInsertedRun(insRun, d, v);
						break;
					case DeletedRun delRun:
						WalkDeletedRun(delRun, d, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, d);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, d);
						break;

					// ---- Office Math (inline & display) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;

					// ---- Nesting (bdo/dir/subDoc allowed under paragraph content) ----
					case BidirectionalOverride nestedBdo:
						WalkBidirectionalOverride(nestedBdo, d, v);
						break;
					case BidirectionalEmbedding nestedDir:
						WalkBidirectionalEmbedding(nestedDir, d, v);
						break;
					case SubDocumentReference subDoc:
						v.VisitSubDocumentReference(subDoc, d);
						break;

					// ---- MC wrapper ----
					case AlternateContent ac:
						WalkAlternateContent(ac, d, v);
						break;

					default:
						WalkUnknown("BidirectionalOverride", child, d, v);
						break;
				}
			}
		}
	}


	private void WalkSmartTagRun(OpenXmlUnknownElement smart, DxpDocumentContext d, DxpIVisitor v)
	{
		if (!DxpSmartTags.IsWSmartTag(smart))
			throw BuildUnsupportedException("SmartTag (expected w:smartTag)", smart);

		// Extract smartTag attributes per spec: w:element, w:uri
		var wNs = smart.NamespaceUri; // current w namespace
		string elementName = smart.GetAttribute("element", wNs).Value ?? string.Empty;
		string elementUri = smart.GetAttribute("uri", wNs).Value ?? string.Empty;

			using (v.VisitSmartTagRunBegin(smart, elementName, elementUri, d))
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
				v.VisitSmartTagProperties(smartTagPr, attrs, d);
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
						WalkRun(r, d, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, d, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, d, v);
						break;
					case SimpleField fld:
						WalkSimpleField(fld, d, v);
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, d, v);
						break;

					// Drawings & legacy pict can appear inline here
					case Drawing drw:
						WalkDrawingTextBox(drw, d, v);
						break;
					case Picture pict:
						WalkVmlTextBox(pict, d, v);
						break;

					// Anchors / permissions / proofing / range markup
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;
					case ProofError perr:
						v.VisitProofError(perr, d);
						break;

					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break;

					// Tracked-change run containers
					case InsertedRun insRun:
						WalkInsertedRun(insRun, d, v);
						break;
					case DeletedRun delRun:
						WalkDeletedRun(delRun, d, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, d);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, d);
						break;

					// Office Math (inline & display)
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;

					// MC wrapper (smartTag can contain AlternateContent)
					case AlternateContent ac:
						WalkAlternateContent(ac, d, v);
						break;

					// Nested smartTag (rare but legal) – recurse on unknown
					case OpenXmlUnknownElement unk when DxpSmartTags.IsWSmartTag(unk):
						WalkSmartTagRun(unk, d, v);
						break;

					default:
						WalkUnknown("SmartTag", child, d, v);
						break;
				}
			}
		}
	}

	private void WalkSdtRun(SdtRun sdtRun, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitSdtRunBegin(sdtRun, d))
		{
			// (1) Properties (optional)
			var pr = sdtRun.SdtProperties;
			if (pr != null)
				v.VisitSdtProperties(pr, d);

			// (2) Content (optional per schema — do NOT throw if missing)
			var content = sdtRun.SdtContentRun;
			if (content != null)
			{
				using (v.VisitSdtContentRunBegin(content, d))
				{
					foreach (var child in content.ChildElements)
					{
						switch (child)
						{
							// ---- Core inline content ----
							case Run r:
								WalkRun(r, d, v);
								break;
							case Hyperlink link:
								WalkHyperlink(link, d, v);
								break;
							case SdtRun nestedSdt:
								WalkSdtRun(nestedSdt, d, v);
								break;
							case SimpleField fld:
								WalkSimpleField(fld, d, v);
								break;
							case CustomXmlRun cxr:
								WalkCustomXmlRun(cxr, d, v);
								break;
							// Not found in SDK: w:smartTag serialized as unknown
							case OpenXmlUnknownElement smart
								when smart.LocalName == "smartTag"
									&& (smart.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
										|| smart.NamespaceUri == "http://purl.oclc.org/ooxml/wordprocessingml/main"):
								WalkSmartTagRun(smart, d, v);
								break;

							// ---- Anchors / range markup ----
							case BookmarkStart bs:
								v.VisitBookmarkStart(bs, d);
								break;
							case BookmarkEnd be:
								v.VisitBookmarkEnd(be, d);
								break;
							case CommentRangeStart crs:
								WalkCommentRangeStart(crs, d, v);
								break;
							case CommentReference cref:
								WalkCommentReference(cref, d, v);
								break;
							case CommentRangeEnd cre:
								break;
							case PermStart ps:
								v.VisitPermStart(ps, d);
								break;
							case PermEnd pe:
								v.VisitPermEnd(pe, d);
								break;
							case ProofError perr:
								v.VisitProofError(perr, d);
								break;

							// ---- Move ranges (location containers) ----
							case MoveFromRangeStart mfrs:
								v.VisitMoveFromRangeStart(mfrs, d);
								break;
							case MoveFromRangeEnd mfre:
								v.VisitMoveFromRangeEnd(mfre, d);
								break;
							case MoveToRangeStart mtrs:
								v.VisitMoveToRangeStart(mtrs, d);
								break;
							case MoveToRangeEnd mtre:
								v.VisitMoveToRangeEnd(mtre, d);
								break;

							// ---- customXml ranges (start/end) ----
							case CustomXmlInsRangeStart cxInsS:
								v.VisitCustomXmlInsRangeStart(cxInsS, d);
								break;
							case CustomXmlInsRangeEnd cxInsE:
								v.VisitCustomXmlInsRangeEnd(cxInsE, d);
								break;
							case CustomXmlDelRangeStart cxDelS:
								v.VisitCustomXmlDelRangeStart(cxDelS, d);
								break;
							case CustomXmlDelRangeEnd cxDelE:
								v.VisitCustomXmlDelRangeEnd(cxDelE, d);
								break;
							case CustomXmlMoveFromRangeStart cxMfS:
								v.VisitCustomXmlMoveFromRangeStart(cxMfS, d);
								break;
							case CustomXmlMoveFromRangeEnd cxMfE:
								v.VisitCustomXmlMoveFromRangeEnd(cxMfE, d);
								break;
							case CustomXmlMoveToRangeStart cxMtS:
								v.VisitCustomXmlMoveToRangeStart(cxMtS, d);
								break;
							case CustomXmlMoveToRangeEnd cxMtE:
								v.VisitCustomXmlMoveToRangeEnd(cxMtE, d);
								break;

							// ---- Tracked-change run containers ----
							case InsertedRun insRun:
								WalkInsertedRun(insRun, d, v);
								break;
							case DeletedRun delRun:
								WalkDeletedRun(delRun, d, v);
								break;
							case MoveFromRun moveFromRun:
								v.VisitMoveFromRun(moveFromRun, d);
								break;
							case MoveToRun moveToRun:
								v.VisitMoveToRun(moveToRun, d);
								break;

							// ---- Office Math (inline & display) ----
							case DocumentFormat.OpenXml.Math.OfficeMath oMath:
								v.VisitOMath(oMath, d);
								break;
							case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
								v.VisitOMathParagraph(oMathPara, d);
								break;

							// ---- Bidi & subdocument ----
							case BidirectionalOverride bdo:
								WalkBidirectionalOverride(bdo, d, v);
								break;
							case BidirectionalEmbedding bdi:
								WalkBidirectionalEmbedding(bdi, d, v);
								break;
							case SubDocumentReference subDoc:
								v.VisitSubDocumentReference(subDoc, d);
								break;

							// ---- MC wrapper ----
							case AlternateContent ac:
								WalkAlternateContent(ac, d, v);
								break;

							default:
								WalkUnknown("SdtContentRun", child, d, v);
								break;
						}
					}
				}
			}

			// (3) End-char run properties (optional) — still visit even if content was absent
			var endPr = sdtRun.SdtEndCharProperties;
			if (endPr != null)
				v.VisitSdtEndCharProperties(endPr, d);
		}
	}

	private void WalkSimpleField(SimpleField fld, DxpDocumentContext d, DxpIVisitor v)
	{
		// w:fldSimple – simple field whose result is represented by its child content
		// Attributes: w:instr (field code), w:dirty, w:fldLock. Behavior: children are the current field result.
		var frame = new FieldFrame { SeenSeparate = true, InResult = true, SuppressResult = false };
		d.CurrentFields.FieldStack.Push(frame);
		using (v.VisitSimpleFieldBegin(fld, d))
		{
			// Optional field data payload (<w:fldData>), rarely used.
			if (fld.FieldData is { } data)
				v.VisitFieldData(data, d);

			foreach (var child in fld.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content (run-universe) ----
					case Run r:
						WalkRun(r, d, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, d, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, d, v);
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, d, v);
						break;
					case OpenXmlUnknownElement smart
						when smart.LocalName == "smartTag" && smart.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main":
					{
						WalkSmartTagRun(smart, d, v);
						break;
					}

					// ---- Anchors / range markup ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;
					case ProofError perr:
						v.VisitProofError(perr, d);
						break;

					// ---- Move ranges (location containers) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break;

					// ---- Tracked-change run containers ----
					case InsertedRun insRun:
						WalkInsertedRun(insRun, d, v);
						break;
					case DeletedRun delRun:
						WalkDeletedRun(delRun, d, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, d);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, d);
						break;

					// ---- Office Math (inline & display) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;

					// ---- Bidi containers & subdocument ----
					case BidirectionalOverride bdo:
						WalkBidirectionalOverride(bdo, d, v);
						break;
					case BidirectionalEmbedding bdi:
						WalkBidirectionalEmbedding(bdi, d, v);
						break;
					case SubDocumentReference subDoc:
						v.VisitSubDocumentReference(subDoc, d);
						break;

					// ---- Markup Compatibility wrapper ----
					case AlternateContent ac:
						WalkAlternateContent(ac, d, v);
						break;

					// ---- Elements listed as parents of fldSimple, not children; fallthrough is correct ----
					// (e.g., another fldSimple wrapping this one is allowed *as parent*, not common as child.)

					default:
						WalkUnknown("SimpleField", child, d, v);
						break;
				}
			}
		}
		d.CurrentFields.FieldStack.Pop();
	}

	private void WalkCustomXmlRun(CustomXmlRun cxr, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitCustomXmlRunBegin(cxr, d))
		{
			// Properties <w:customXmlPr> (element name/namespace, optional data binding metadata)
			var pr = cxr.CustomXmlProperties;
			if (pr != null)
				v.VisitCustomXmlProperties(pr, d);

			foreach (var child in cxr.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content (run-universe) ----
					case Run r:
						WalkRun(r, d, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, d, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, d, v);
						break;
					case SimpleField fld:
						WalkSimpleField(fld, d, v);
						break;

					// Nested inline customXml/smartTag (both valid in CT_CustomXmlRun)
					case CustomXmlRun nested:
						WalkCustomXmlRun(nested, d, v);
						break;

					// Not found in SDK
					case OpenXmlUnknownElement smart
						when smart.LocalName == "smartTag" && smart.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main":
					{
						WalkSmartTagRun(smart, d, v);
						break;
					}

					// ---- Bookmarks & comment anchors ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;

					// ---- Permissions ----
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;

					// ---- customXml range markup (insert/delete/move; start/end) ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, d);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, d);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, d);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, d);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, d);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, d);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, d);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, d);
						break;

					// ---- Tracked change/move run containers (inline) ----
					case InsertedRun insRun:
						WalkInsertedRun(insRun, d, v);
						break;
					case DeletedRun delRun:
						WalkDeletedRun(delRun, d, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, d);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, d);
						break;

					// ---- Move location containers (range start/end) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break;

					// ---- Office Math (inline & paragraph forms) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;

					// ---- Bidi containers ----
					case BidirectionalOverride bdo:
						WalkBidirectionalOverride(bdo, d, v);
						break;
					case BidirectionalEmbedding bdi:
						WalkBidirectionalEmbedding(bdi, d, v);
						break;

					// ---- Subdocument anchor ----
					case SubDocumentReference subDoc:
						v.VisitSubDocumentReference(subDoc, d);
						break;

					// ---- Markup Compatibility wrapper (practically appears anywhere) ----
					case AlternateContent ac:
						WalkAlternateContent(ac, d, v);
						break;

					// ---- Proofing errors ----
					case ProofError perr:
						v.VisitProofError(perr, d);
						break;

					default:
						WalkUnknown("CustomXmlRun", child, d, v);
						break;
				}
			}
		}
	}

	private void WalkDeletedRun(DeletedRun dr, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitDeletedRunBegin(dr, d))
		{
			foreach (var child in dr.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content allowed inside CT_RunTrackChange ----
					case Run r:
						WalkRun(r, d, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, d, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, d, v);
						break;
					case SimpleField fld:
						// Let WalkSimpleField open the visitor scope and extract instr/flags
						WalkSimpleField(fld, d, v);
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, d, v);
						break;
					// Not found in SDK
					case OpenXmlUnknownElement smart
						when smart.LocalName == "smartTag" && smart.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main":
					{
						WalkSmartTagRun(smart, d, v);
						break;
					}

					// ---- Anchors / proofing / permissions (range markup) ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;
					case ProofError perr:
						v.VisitProofError(perr, d);
						break;

					// ---- Move ranges (location containers) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break;

					// ---- customXml range markup (start/end) ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, d);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, d);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, d);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, d);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, d);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, d);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, d);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, d);
						break;

					// ---- Tracked-change run containers (nesting is allowed) ----
					case InsertedRun insRun:
						WalkInsertedRun(insRun, d, v);
						break;
					case DeletedRun innerDel:
						// Nested del inside del is rare but legal per CT_RunTrackChange; recurse.
						WalkDeletedRun(innerDel, d, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, d);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, d);
						break;

					// ---- Office Math (both forms are permitted) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;

					// ---- Markup Compatibility wrapper ----
					case AlternateContent ac:
						WalkAlternateContent(ac, d, v);
						break;

					default:
						WalkUnknown("DeletedRun", child, d, v);
						break;
				}
			}
		}
	}


	private void WalkInsertedRun(InsertedRun ir, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitInsertedRunBegin(ir, d))
		{
			foreach (var child in ir.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content ----
					case Run r:
						WalkRun(r, d, v);
						break;
					case Hyperlink link:
						WalkHyperlink(link, d, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, d, v);
						break;
					case SimpleField fld:
						WalkSimpleField(fld, d, v); // lets the callee open/close the visitor scope
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, d, v);
						break;
					// Not found in SDK
					case OpenXmlUnknownElement smart
						when smart.LocalName == "smartTag" && smart.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main":
					{
						WalkSmartTagRun(smart, d, v);
						break;
					}

					// ---- Anchors / proofing / permissions ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;
					case ProofError perr:
						v.VisitProofError(perr, d);
						break;

					// ---- Move ranges (location containers) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break;

					// ---- customXml range markup (start/end) ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, d);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, d);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, d);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, d);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, d);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, d);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, d);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, d);
						break;

					// ---- Tracked-change run containers (nesting allowed) ----
					case InsertedRun innerIns:
						WalkInsertedRun(innerIns, d, v);
						break;
					case DeletedRun dr:
						WalkDeletedRun(dr, d, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, d);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, d);
						break;

					// ---- Office Math (inline & display) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;

					// ---- Markup Compatibility wrapper ----
					case AlternateContent ac:
						WalkAlternateContent(ac, d, v);
						break;

					default:
						WalkUnknown("InsertedRun", child, d, v);
						break;
				}
			}
		}
	}



	private void WalkHyperlink(Hyperlink link, DxpDocumentContext d, DxpIVisitor v)
	{
		DxpLinkAnchor? target = DxpHyperlinks.ResolveHyperlinkTarget(link, d.CurrentPart ?? d.MainDocumentPart);

		using (v.VisitHyperlinkBegin(link, target, d))
		{

			// <w:hyperlink> can host the full run-level “inline universe”
			foreach (var child in link.ChildElements)
			{
				switch (child)
				{
					// ---- Core inline content ----
					case Run r:
						WalkRun(r, d, v);
						break;
					case Hyperlink nested:
						WalkHyperlink(nested, d, v);
						break;
					case SdtRun sdtRun:
						WalkSdtRun(sdtRun, d, v);
						break;
					case SimpleField fld:
						WalkSimpleField(fld, d, v);
						break;
					case CustomXmlRun cxr:
						WalkCustomXmlRun(cxr, d, v);
						break;

					// ---- Anchors / permissions / proofing ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;
					case ProofError perr:
						v.VisitProofError(perr, d);
						break;

					// ---- Move ranges (location containers) ----
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break;

					// ---- customXml ranges (start/end) + Office 2010 conflict ranges ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, d);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, d);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, d);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, d);
						break;
					case CustomXmlMoveFromRangeStart cxMfS:
						v.VisitCustomXmlMoveFromRangeStart(cxMfS, d);
						break;
					case CustomXmlMoveFromRangeEnd cxMfE:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfE, d);
						break;
					case CustomXmlMoveToRangeStart cxMtS:
						v.VisitCustomXmlMoveToRangeStart(cxMtS, d);
						break;
					case CustomXmlMoveToRangeEnd cxMtE:
						v.VisitCustomXmlMoveToRangeEnd(cxMtE, d);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeStart cxCis:
						v.VisitCustomXmlConflictInsertionRangeStart(cxCis, d);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeEnd cxCie:
						v.VisitCustomXmlConflictInsertionRangeEnd(cxCie, d);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeStart cxCds:
						v.VisitCustomXmlConflictDeletionRangeStart(cxCds, d);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeEnd cxCde:
						v.VisitCustomXmlConflictDeletionRangeEnd(cxCde, d);
						break;

					// ---- Tracked-change run containers ----
					case InsertedRun insRun:
						WalkInsertedRun(insRun, d, v);
						break;
					case DeletedRun delRun:
						WalkDeletedRun(delRun, d, v);
						break;
					case MoveFromRun moveFromRun:
						v.VisitMoveFromRun(moveFromRun, d);
						break;
					case MoveToRun moveToRun:
						v.VisitMoveToRun(moveToRun, d);
						break;

					// ---- Office Math (inline & display) ----
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;
						// When fine-grained m:* nodes are exposed elsewhere, forward them to v.VisitOMathElement here.

					// ---- Bidi & subdocument ----
					case BidirectionalOverride bdo:
						WalkBidirectionalOverride(bdo, d, v);
						break;
					case BidirectionalEmbedding bdi:
						WalkBidirectionalEmbedding(bdi, d, v);
						break;
					case SubDocumentReference subDoc:
						v.VisitSubDocumentReference(subDoc, d);
						break;

					// ---- ContentPart (Office 2010) ----
					case ContentPart cp:
						v.VisitContentPart(cp, d);
						break;

					// ---- Markup Compatibility wrapper (can appear anywhere) ----
					case AlternateContent ac:
						WalkAlternateContent(ac, d, v);
						break;

					default:
						WalkUnknown("Hyperlink", child, d, v);
						break;
				}
			}

			// Ensure inline styles are closed before closing the anchor so tag order stays valid.
			d.StyleTracker.ResetStyle(d, v);
		}
	}

	// Visitor asks: “Do I support this Choice?” If yes, we process it and stop.
	// If none are accepted, we process Fallback (if present). Else drop content per MCE.
	private void WalkAlternateContent(AlternateContent ac, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitAlternateContentBegin(ac, d))
		{
			foreach (var choice in ac.Elements<AlternateContentChoice>())
			{
				var required = DxpAlternateContents.GetRequiredPrefixes(choice); // eg: ["w14","wps"]
															// The visitor should return true ONLY if it supports every required namespace.
				if (v.AcceptAlternateContentChoice(choice, required, d))
				{
					WalkAlternateContentSelectedContainer(choice, d, v);
					return;
				}
			}

			var fallback = ac.Elements<AlternateContentFallback>().FirstOrDefault();
			if (fallback != null)
			{
				WalkAlternateContentSelectedContainer(fallback, d, v);
				return;
			}

			// No accepted Choice and no Fallback => content is ignored (as if it didn't exist).
			// This matches the MCE preprocessing model.
		}
	}

	private void WalkAlternateContentSelectedContainer(OpenXmlElement container, DxpDocumentContext d, DxpIVisitor v)
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
					WalkBlock(child, d, v);
					break;

				// Inline (common)
				case Run r:
					WalkRun(r, d, v);
					break;
				case Hyperlink link:
					WalkHyperlink(link, d, v);
					break;
				case SdtRun sdtRun:
					WalkSdtRun(sdtRun, d, v);
					break;
				case SimpleField fld:
					WalkSimpleField(fld, d, v);
					break;

				// Drawings / VML textboxes
				case Drawing drw:
					WalkDrawingTextBox(drw, d, v);
					break;
				case Picture pict:
					WalkVmlTextBox(pict, d, v);
					break;

				// Other allowed inline containers
				case ContentPart cp:
					v.VisitContentPart(cp, d);
					break;
				case DocumentFormat.OpenXml.Math.OfficeMath oMath:
					v.VisitOMath(oMath, d);
					break;
				case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
					v.VisitOMathParagraph(oMathPara, d);
					break;

				// Nested MC
				case AlternateContent nested:
					WalkAlternateContent(nested, d, v);
					break;

				default:
					WalkUnknown("AlternateContent selected branch", child, d, v);
					break;
			}
		}
	}

	private void WalkUnknown(string context, OpenXmlElement el, DxpDocumentContext d, DxpIVisitor v)
	{
		v.VisitUnknown(context, el, d);
	}

	private void WalkDrawingTextBox(Drawing drw, DxpDocumentContext d, DxpIVisitor v)
	{
		var info = d.Drawings.TryResolveDrawingInfo(d.CurrentPart ?? d.MainDocumentPart, drw);
		using (v.VisitDrawingBegin(drw, info, d))
		{
			// Look for Office 2010 Wordprocessing shape textbox: <wps:txbx>
			// SDK types live under DocumentFormat.OpenXml.Office2010.Word.DrawingShape
			var txbx = drw
				.Descendants<DocumentFormat.OpenXml.Office2010.Word.DrawingShape.TextBoxInfo2>() // wps:txbx
				.FirstOrDefault();
			if (txbx == null)
				return;

			var content = txbx.GetFirstChild<TextBoxContent>(); // w:txbxContent
			if (content == null)
				return;

			WalkTextBoxContent(content, d, v);
		}
	}

	private void WalkVmlTextBox(Picture pict, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitLegacyPictureBegin(pict, d))
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

			WalkTextBoxContent(content, d, v);
		}
	}

	private void WalkTextBoxContent(TextBoxContent txbx, DxpDocumentContext d, DxpIVisitor v)
	{
		// Optional: let the visitor know we’re entering a text box body
		using (v.VisitTextBoxContentBegin(txbx, d))
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
							WalkBlock(child, d, v); // reuse the normal block dispatcher
							break;

					// Math is also allowed here per SDK child list for TextBoxContent
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;

					default:
						WalkUnknown("TextBoxContent", child, d, v);
						break;
				}
			}
		}
	}

	private void WalkRun(Run r, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitRunBegin(r, d))
		{
			// Resolve style only if we can find a paragraph context.
			var para = r.Ancestors<Paragraph>().FirstOrDefault();
			bool hasRenderable = r.ChildElements.Any(child =>
				child is Text or DeletedText or NoBreakHyphen or TabChar or Break or CarriageReturn or Drawing);
			if (para != null)
			{
				if (hasRenderable)
				{
					DxpStyleEffectiveRunStyle style = d.Styles.ResolveRunStyle(para, r);
					d.StyleTracker.ApplyStyle(style, d, v);
				}
			}
			else
			{
				// No paragraph ancestor — surface to the visitor but keep walking content.
				WalkUnknown("Run (no Paragraph ancestor)", r, d, v);
			}

			foreach (var child in r.ChildElements)
			{
				switch (child)
				{
					case NoBreakHyphen h:
						v.VisitNoBreakHyphen(h, d);
						break;

					case LastRenderedPageBreak pb:
						v.VisitLastRenderedPageBreak(pb, d);
						break;

					case RunProperties rp:
						v.VisitRunProperties(rp, d);
						break;

					case DeletedText dt:
						v.VisitDeletedText(dt, d);
						break;

					case Text t:
						v.VisitText(t, d);
						break;

					case TabChar tab:
						v.VisitTab(tab, d);
						break;

					case Break br:
						v.VisitBreak(br, d);
						break;

					case CarriageReturn cr:
						v.VisitCarriageReturn(cr, d);
						break;

					case Drawing drw:
						WalkDrawingTextBox(drw, d, v);
						break;

					case FieldChar fc:
					{
						var t = fc.FieldCharType?.Value;

						if (t == FieldCharValues.Begin)
						{
							var frame = new FieldFrame { SeenSeparate = false, ResultScope = null, InResult = false, SuppressResult = false };
							d.CurrentFields.FieldStack.Push(frame);
							v.VisitComplexFieldBegin(fc, d);
						}
						else if (t == FieldCharValues.Separate)
						{
							if (d.CurrentFields.FieldStack.Count > 0)
							{
								var top = d.CurrentFields.FieldStack.Pop();
								if (!top.SeenSeparate)
								{
									v.VisitComplexFieldSeparate(fc, d);
									top.SeenSeparate = true;
									top.InResult = true;
									if (top.ResultScope == null)
										top.ResultScope = v.VisitComplexFieldResultBegin(d);
								}
								d.CurrentFields.FieldStack.Push(top);
							}
							else
							{
								// stray separate; surface but don’t crash
								v.VisitComplexFieldSeparate(fc, d);
							}
						}
						else if (t == FieldCharValues.End)
						{
							if (d.CurrentFields.FieldStack.Count > 0)
							{
								var top = d.CurrentFields.FieldStack.Pop();
								top.InResult = false;
								top.ResultScope?.Dispose();
								v.VisitComplexFieldEnd(fc, d);
							}
							else
							{
								// stray end; surface but don’t crash
								v.VisitComplexFieldEnd(fc, d);
							}
						}
						// Other FieldChar types (rare) — ignore.
						break;
					}

					case FieldCode code:
					{
						// FieldCode.Text can be null; InnerText is a safe fallback
						var instr = code.Text ?? code.InnerText ?? string.Empty;
						v.VisitComplexFieldInstruction(code, instr, d);
						// Do not emit as visible text; instruction is not the result.
						break;
					}

					case FootnoteReference fr:
					{
						long fnId = fr.Id?.Value ?? 0;
						if (d.Footnotes.Resolve(fnId, out int fnIndex))
							v.VisitFootnoteReference(fr, new DxpFootnoteContext(fnId, fnIndex), d);
						break;
					}

					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;

					case AlternateContent ac:
						WalkAlternateContent(ac, d, v);
						break;

					// Legacy DATE/PAGENUM-style blocks (non-editable placeholders)
					case DayShort ds:
						v.VisitDayShort(ds, d);
						break;
					case MonthShort ms:
						v.VisitMonthShort(ms, d);
						break;
					case YearShort ys:
						v.VisitYearShort(ys, d);
						break;
					case DayLong dl:
						v.VisitDayLong(dl, d);
						break;
					case MonthLong ml:
						v.VisitMonthLong(ml, d);
						break;
					case YearLong yl:
						v.VisitYearLong(yl, d);
						break;
					case PageNumber pn:
						v.VisitPageNumber(pn, d);
						break;

					// Marks and references (footnotes/endnotes/annotations/separators)
					case AnnotationReferenceMark arm:
						v.VisitAnnotationReference(arm, d);
						break;
					case FootnoteReferenceMark frm:
						if (d.Footnotes.Resolve(d.CurrentFootnote?.Id ?? 0, out int index))
							v.VisitFootnoteReferenceMark(frm, new DxpFootnoteContext(d.CurrentFootnote.Id, index), d);
						break;
					case EndnoteReferenceMark erm:
						v.VisitEndnoteReferenceMark(erm, d);
						break;
					case EndnoteReference enr:
						v.VisitEndnoteReference(enr, d);
						break;
					case SeparatorMark sep:
						v.VisitSeparatorMark(sep, d);
						break;
					case ContinuationSeparatorMark csep:
						v.VisitContinuationSeparatorMark(csep, d);
						break;

					// Characters / inline controls
					case SoftHyphen sh:
						v.VisitSoftHyphen(sh, d);
						break;
					case SymbolChar sym:
						v.VisitSymbol(sym, d);
						break;
					case PositionalTab ptab:
						v.VisitPositionalTab(ptab, d);
						break;
					case Ruby ruby:
						WalkRuby(ruby, d, v);
						break;

					// Fields (deleted instruction text)
					case DeletedFieldCode dfc:
						v.VisitDeletedFieldCode(dfc, d);
						break;

					/* Legacy/object content within run */
					case EmbeddedObject obj:
						v.VisitEmbeddedObject(obj, d);
						break;
					case Picture pict:
						WalkVmlTextBox(pict, d, v);
						break;

					case ContentPart cp:
						v.VisitContentPart(cp, d);
						break;

					default:
						WalkUnknown("Run child", child, d, v);
						break;
				}
			}
		}
	}

	// Walk a ruby (phonetic guide) inline container: <w:ruby> -> <w:rubyPr>, <w:rt>, <w:rubyBase>
	private void WalkRuby(Ruby ruby, DxpDocumentContext d, DxpIVisitor v)
	{
		// Begin ruby scope (visitor can choose how to render a ruby run)
		using (v.VisitRubyBegin(ruby, d))
		{
			// --- Properties: <w:rubyPr> controls alignment/size/raise/lang of the ruby text ---
				var pr = ruby.GetFirstChild<RubyProperties>(); // SDK class for <w:rubyPr>
				if (pr != null)
					v.VisitRubyProperties(pr, d); // exposes rubyAlign, hps, hpsRaise, hpsBaseText, lid, dirty
												  // Spec: rubyPr is required. We log if missing, but don’t throw to be resilient.
												  // (Phonetic Guide Properties per CT_RubyPr.)  // ISO/IEC 29500 Part 1: w:rubyPr.

				// --- Ruby text: <w:rt> holds phonetic text in a required single <w:r> ---
				RubyContent? rt = ruby.GetFirstChild<RubyContent>(); // SDK: RubyContent for <w:rt>
				if (rt != null)
					WalkRubyContent(rt, isBase: false, d, v); // CT_RubyContent (phonetic guide text).

				// --- Base text: <w:rubyBase> holds the base characters in a required single <w:r> ---
				RubyBase? rb = ruby.GetFirstChild<RubyBase>();    // SDK: RubyBase for <w:rubyBase>
				if (rb != null)
					WalkRubyContent(rb, isBase: true, d, v); // CT_RubyContent (base text).
			}
		}

		// Walk CT_RubyContent (<w:rt> or <w:rubyBase>): spec says it MUST contain exactly one <w:r>,
		// but can also include select inline/range markup (proofErr, bookmarks, math, etc.).
	private void WalkRubyContent(RubyContentType rc, bool isBase, DxpDocumentContext d, DxpIVisitor v)
	{
		// Allow the visitor to differentiate ruby text vs base text if useful
		using (v.VisitRubyContentBegin(rc, isBase, d))
		{
			foreach (var child in rc.ChildElements)
			{
				switch (child)
				{
						case Run r:
							WalkRun(r, d, v); // reuse the existing run walker
						break;

					// Inline/range markup permitted by CT_RubyContent — forward to existing visitor hooks:
					case ProofError pe:
						v.VisitProofError(pe, d);
					break;

				case BookmarkStart bs:
					v.VisitBookmarkStart(bs, d);
					break;
				case BookmarkEnd be:
					v.VisitBookmarkEnd(be, d);
					break;

					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe2:
						v.VisitPermEnd(pe2, d);
						break;

					// Tracked move/ins/del regions allowed here (container-level markers):
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break;

					case Inserted ins:
						WalkInserted(ins, d, v);
						break;
					case Deleted del:
						WalkDeleted(del, d, v);
						break;

					// Office Math is explicitly allowed here:
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;

					// Compatibility: if spec groups (EG_RunLevelElts / EG_RangeMarkup / EG_MathContent)
					// expose additional children via SDK updates, route them or throw here.

					default:
						WalkUnknown(isBase ? "RubyBase" : "RubyText", child, d, v);
						break;
				}
			}
		}
	}

	// CT_TrackChange leaf: w:del under trPr = deleted table row,
	// w:del under pPr/rPr = deleted paragraph mark. No children to walk.
	private void WalkDeleted(Deleted del, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitDeletedBegin(del, d))
		{

			// Case 1: table row deletion (w:trPr/w:del)
			if (del.Parent is TableRowProperties trPr)
			{
				var tr = trPr.Parent as TableRow; // usually non-null when in a live tree
				v.VisitDeletedTableRowMark(del, trPr, tr, d); // tell visitor: the row is marked deleted
				return;
			}

			// Case 2: paragraph mark deletion (w:pPr/w:rPr/w:del)
			// This marks the *paragraph mark* deleted (contents merged with next para per spec).
			if (del.Parent is RunProperties rPr && rPr.Parent is ParagraphProperties pPr)
			{
				var p = pPr.Parent as Paragraph;
				v.VisitDeletedParagraphMark(del, pPr, p, d); // tell visitor: paragraph mark is deleted
				return;
			}

			// Anything else is unexpected for w:del (schema lists trPr and rPr parents).
			WalkUnknown("Deleted", del, d, v);
			return;
		}
	}


	// CT_TrackChange leaf: w:ins marks an insertion on (1) a table row, (2) paragraph numbering props, or (3) the paragraph mark.
	// It has no children; we only need to determine the parent scope and notify the visitor.  Spec: w:ins under trPr / numPr / (pPr/rPr).
	private void WalkInserted(Inserted ins, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitInsertedBegin(ins, d))
		{
			// (1) Inserted table row: <w:trPr><w:ins .../></w:trPr>  ⇒ the row itself is marked as inserted.
			// Ref: "Inserted Table Row" section + example under Parent Elements trPr. 
			if (ins.Parent is TableRowProperties trPr)
			{
				var tr = trPr.Parent as TableRow;
				v.VisitInsertedTableRowMark(ins, trPr, tr, d);
				return;
			}

			// (2) Inserted numbering properties: <w:pPr><w:numPr>...<w:ins .../></w:numPr></w:pPr>
			// Ref: "Inserted Numbering Properties" section; parent element listed as numPr.
			if (ins.Parent is NumberingProperties numPr)
			{
				var pPr = numPr.Parent as ParagraphProperties;
				var p = pPr?.Parent as Paragraph;
				v.VisitInsertedNumberingProperties(ins, numPr, pPr, p, d);
				return;
			}

			// (3) Inserted paragraph mark: <w:pPr><w:rPr><w:ins .../></w:rPr></w:pPr>
			// Ref: "Inserted Paragraph" section; parent elements listed as rPr (within pPr).
			if (ins.Parent is RunProperties rPr && rPr.Parent is ParagraphProperties pPr2)
			{
				var p = pPr2.Parent as Paragraph;
				v.VisitInsertedParagraphMark(ins, pPr2, p, d);
				return;
			}

			// Any other scope would be unexpected for w:ins per the schema; keep strict to surface it early.
			WalkUnknown("Inserted", ins, d, v);
			return;
		}
	}

	private void WalkEndnote(Endnote fn, int index, long id, DxpDocumentContext d, DxpIVisitor v)
	{
		using (d.PushFootnote(id, index, out DxpFootnoteContext footnote))
		using (d.PushCurrentPart(d.MainDocumentPart?.EndnotesPart))
		using (v.VisitEndnoteBegin(fn, id, index, d))
		{
			foreach (var child in fn.ChildElements)
			{
				switch (child)
				{
					// Block content
					case Paragraph p:
						WalkBlock(p, d, v);
						break;
					case Table t:
						WalkBlock(t, d, v);
						break;
					case SdtBlock sdt:
						WalkSdtBlock(sdt, d, v);
						break;
					case CustomXmlBlock cx:
						WalkCustomXmlBlock(cx, d, v);
						break;
					case AltChunk ac:
						v.VisitAltChunk(ac, d);
						break;

					// Math
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;

					// Anchors
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;

					// Permissions & proofing
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;
					case ProofError pr:
						v.VisitProofError(pr, d);
						break;

					// Tracked ranges
					case Inserted ins:
						WalkInserted(ins, d, v);
						break;
					case Deleted del:
						WalkDeleted(del, d, v);
						break;
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break;
					case CustomXmlInsRangeStart cins:
						v.VisitCustomXmlInsRangeStart(cins, d);
						break;
					case CustomXmlInsRangeEnd cine:
						v.VisitCustomXmlInsRangeEnd(cine, d);
						break;
					case CustomXmlDelRangeStart cdls:
						v.VisitCustomXmlDelRangeStart(cdls, d);
						break;
					case CustomXmlDelRangeEnd cdle:
						v.VisitCustomXmlDelRangeEnd(cdle, d);
						break;
					case CustomXmlMoveFromRangeStart cmfs:
						v.VisitCustomXmlMoveFromRangeStart(cmfs, d);
						break;
					case CustomXmlMoveFromRangeEnd cmfe:
						v.VisitCustomXmlMoveFromRangeEnd(cmfe, d);
						break;
					case CustomXmlMoveToRangeStart cmts:
						v.VisitCustomXmlMoveToRangeStart(cmts, d);
						break;
					case CustomXmlMoveToRangeEnd cmte:
						v.VisitCustomXmlMoveToRangeEnd(cmte, d);
						break;

					default:
						WalkUnknown("Endnote", child, d, v);
						break;
				}
			}
			d.StyleTracker.ResetStyle(d, v);
		}
	}

	private void WalkFootnote(Footnote fn, int index, long id, DxpDocumentContext d, DxpIVisitor v)
	{
		using (d.PushFootnote(id, index, out DxpFootnoteContext footnote))
		using (d.PushCurrentPart(d.MainDocumentPart?.FootnotesPart))
		using (v.VisitFootnoteBegin(fn, footnote, d))
		{
			foreach (var child in fn.ChildElements)
			{
				switch (child)
				{
					// Block content
					case Paragraph p:
						WalkBlock(p, d, v);
						break;
					case Table t:
						WalkBlock(t, d, v);
						break;
					case SdtBlock sdt:
						WalkSdtBlock(sdt, d, v);
						break;
					case CustomXmlBlock cx:
						WalkCustomXmlBlock(cx, d, v);
						break;
					case AltChunk ac:
						v.VisitAltChunk(ac, d);
						break;

					// Math
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;

					// Anchors
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;

					// Permissions & proofing
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;
					case ProofError pr:
						v.VisitProofError(pr, d);
						break;

					// Tracked ranges
					case Inserted ins:
						WalkInserted(ins, d, v);
						break;
					case Deleted del:
						WalkDeleted(del, d, v);
						break;
					case MoveFromRangeStart mfrs:
						v.VisitMoveFromRangeStart(mfrs, d);
						break;
					case MoveFromRangeEnd mfre:
						v.VisitMoveFromRangeEnd(mfre, d);
						break;
					case MoveToRangeStart mtrs:
						v.VisitMoveToRangeStart(mtrs, d);
						break;
					case MoveToRangeEnd mtre:
						v.VisitMoveToRangeEnd(mtre, d);
						break;
					case CustomXmlInsRangeStart cins:
						v.VisitCustomXmlInsRangeStart(cins, d);
						break;
					case CustomXmlInsRangeEnd cine:
						v.VisitCustomXmlInsRangeEnd(cine, d);
						break;
					case CustomXmlDelRangeStart cdls:
						v.VisitCustomXmlDelRangeStart(cdls, d);
						break;
					case CustomXmlDelRangeEnd cdle:
						v.VisitCustomXmlDelRangeEnd(cdle, d);
						break;
					case CustomXmlMoveFromRangeStart cmfs:
						v.VisitCustomXmlMoveFromRangeStart(cmfs, d);
						break;
					case CustomXmlMoveFromRangeEnd cmfe:
						v.VisitCustomXmlMoveFromRangeEnd(cmfe, d);
						break;
					case CustomXmlMoveToRangeStart cmts:
						v.VisitCustomXmlMoveToRangeStart(cmts, d);
						break;
					case CustomXmlMoveToRangeEnd cmte:
						v.VisitCustomXmlMoveToRangeEnd(cmte, d);
						break;

					default:
						WalkUnknown("Footnote", child, d, v);
						break;
				}
			}
			d.StyleTracker.ResetStyle(d, v);
		}
	}

	private void WalkCustomXmlBlock(CustomXmlBlock cx, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitCustomXmlBlockBegin(cx, d))
		{
			// Properties (<w:customXmlPr>) carry element name/namespace and optional data binding metadata.
			var pr = cx.CustomXmlProperties;
			if (pr != null)
				v.VisitCustomXmlProperties(pr, d);

			foreach (var child in cx.ChildElements)
			{
				switch (child)
				{
					// ---- Block-level content allowed by CT_CustomXmlBlock ----
					case Paragraph p:
						WalkBlock(p, d, v);
						break;
					case Table t:
						WalkBlock(t, d, v);
						break;
					case SdtBlock sdt:
						WalkSdtBlock(sdt, d, v);
						break;
					case CustomXmlBlock nested:
						WalkCustomXmlBlock(nested, d, v);
						break;

					// ---- Office Math (explicitly permitted here) ----
					case DocumentFormat.OpenXml.Math.Paragraph oMathPara:
						v.VisitOMathParagraph(oMathPara, d);
						break;
					case DocumentFormat.OpenXml.Math.OfficeMath oMath:
						v.VisitOMath(oMath, d);
						break;

					// ---- Bookmark/comment anchors (range markup) ----
					case BookmarkStart bs:
						v.VisitBookmarkStart(bs, d);
						break;
					case BookmarkEnd be:
						v.VisitBookmarkEnd(be, d);
						break;
					case CommentRangeStart crs:
						WalkCommentRangeStart(crs, d, v);
						break;
					case CommentReference cref:
						WalkCommentReference(cref, d, v);
						break;
					case CommentRangeEnd cre:
						break;

					// ---- Permissions & proofing anchors ----
					case PermStart ps:
						v.VisitPermStart(ps, d);
						break;
					case PermEnd pe:
						v.VisitPermEnd(pe, d);
						break;
					case ProofError perr:
						v.VisitProofError(perr, d);
						break;

					// ---- Custom XML range markup (insert/delete/move; start/end) ----
					case CustomXmlInsRangeStart cxInsS:
						v.VisitCustomXmlInsRangeStart(cxInsS, d);
						break;
					case CustomXmlInsRangeEnd cxInsE:
						v.VisitCustomXmlInsRangeEnd(cxInsE, d);
						break;
					case CustomXmlDelRangeStart cxDelS:
						v.VisitCustomXmlDelRangeStart(cxDelS, d);
						break;
					case CustomXmlDelRangeEnd cxDelE:
						v.VisitCustomXmlDelRangeEnd(cxDelE, d);
						break;
					case CustomXmlMoveFromRangeStart cxMfs:
						v.VisitCustomXmlMoveFromRangeStart(cxMfs, d);
						break;
					case CustomXmlMoveFromRangeEnd cxMfe:
						v.VisitCustomXmlMoveFromRangeEnd(cxMfe, d);
						break;
					case CustomXmlMoveToRangeStart cxMts:
						v.VisitCustomXmlMoveToRangeStart(cxMts, d);
						break;
					case CustomXmlMoveToRangeEnd cxMte:
						v.VisitCustomXmlMoveToRangeEnd(cxMte, d);
						break;

					// ---- Office 2010 conflict range markup (optional) ----
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeStart cxCis:
						v.VisitCustomXmlConflictInsertionRangeStart(cxCis, d);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictInsertionRangeEnd cxCie:
						v.VisitCustomXmlConflictInsertionRangeEnd(cxCie, d);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeStart cxCds:
						v.VisitCustomXmlConflictDeletionRangeStart(cxCds, d);
						break;
					case DocumentFormat.OpenXml.Office2010.Word.CustomXmlConflictDeletionRangeEnd cxCde:
						v.VisitCustomXmlConflictDeletionRangeEnd(cxCde, d);
						break;

						// ---- Run-level change containers that are valid children here ----
						case InsertedRun insRun:
							WalkInsertedRun(insRun, d, v);
							break;
						case DeletedRun delRun:
							WalkDeletedRun(delRun, d, v);
							break;
						case MoveFromRun mfr:
							v.VisitMoveFromRun(mfr, d); // WalkMoveFromRun can be used if expanded
							break;
						case MoveToRun mtr:
							v.VisitMoveToRun(mtr, d);   // WalkMoveToRun can be used if expanded
							break;

					// ---- ContentPart (Office 2010) – embedded external content reference ----
					case ContentPart cp:
						v.VisitContentPart(cp, d);
						break;

					default:
						WalkUnknown("CustomXmlBlock", child, d, v);
						break;
				}
			}
		}
	}

	private void WalkSdtBlock(SdtBlock sdt, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitSdtBlockBegin(sdt, d))
		{
			// (1) Properties (optional)
			var pr = sdt.GetFirstChild<SdtProperties>();
			if (pr != null)
				v.VisitSdtProperties(pr, d);

			// (2) Content (optional per schema — do NOT throw if missing)
			var content = sdt.SdtContentBlock;
			if (content == null)
			{
				// Empty SDT is valid: just end the SDT scope
				return;
			}

			using (v.VisitSdtContentBlockBegin(content, d))
			{
				foreach (var child in content.ChildElements)
				{
					// SDT block content is normal block content
					WalkBlock(child, d, v);
				}
			}
		}
	}

	private void WalkCommentContent(DxpCommentInfo info, DxpCommentThread thread, DxpDocumentContext d, DxpIVisitor v)
	{
		using (v.VisitCommentBegin(info, thread, d))
		{
			using (d.PushCurrentPart(info.Part ?? d.CurrentPart))
			{
				if (info.Blocks != null && info.Blocks.Count > 0)
				{
					foreach (var block in info.Blocks)
						WalkBlock(block, d, v);

					d.StyleTracker.ResetStyle(d, v);
					return;
				}

				if (string.IsNullOrEmpty(info.Text))
					return;

				var paragraph = new Paragraph(new Run(new Text(info.Text)));
				WalkBlock(paragraph, d, v);
				d.StyleTracker.ResetStyle(d, v);
			}
		}
	}

	// ---- Strict error reporting ----

	private static NotSupportedException BuildUnsupportedException(string context, OpenXmlElement el)
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


	private void WalkCommentReference(CommentReference cref, DxpDocumentContext d, DxpIVisitor v)
	{
		string id = cref.Id?.Value ?? string.Empty;
		WalkCommentThread(id, d, v);
	}

	private void WalkCommentRangeStart(CommentRangeStart crs, DxpDocumentContext d, DxpIVisitor v)
	{
		string id = crs.Id?.Value ?? string.Empty;
		WalkCommentThread(id, d, v);
	}

	private void WalkCommentThread(string id, DxpDocumentContext d, DxpIVisitor v)
	{
		DxpCommentThread? thread = d.Comments.GetThreadForAnchor(id);
		if (thread == null || thread.Comments.Count == 0)
			return;

		using (v.VisitCommentThreadBegin(id, thread, d))
		{
			foreach (DxpCommentInfo info in thread.Comments)
			{
				WalkCommentContent(info, thread, d, v);
			}
		}
	}

	private sealed class DxpTableContext : DxpITableContext
	{
		public Table Table { get; }
		public TableProperties? Properties { get; private set; }
		public TableGrid? Grid { get; private set; }

		public DxpTableContext(Table table, TableProperties? properties, TableGrid? grid)
		{
			Table = table;
			Properties = properties;
			Grid = grid;
		}

		public void SetGrid(TableGrid grid)
		{
			Grid = grid;
		}
	}

	private sealed class DxpTableRowContext : DxpITableRowContext
	{
		public DxpITableContext Table { get; }
		public bool IsHeader { get; }
		public int Index { get; }

		public DxpTableRowContext(DxpITableContext table, int index, bool isHeader)
		{
			Table = table;
			Index = index;
			IsHeader = isHeader;
		}
	}

	private sealed class DxpTableCellContext : DxpITableCellContext
	{
		public DxpITableRowContext Row { get; }
		public int RowIndex { get; }
		public int ColumnIndex { get; }
		public int RowSpan { get; }
		public int ColSpan { get; }
		public TableCellProperties? Properties { get; }

		public DxpTableCellContext(DxpITableRowContext row, int rowIndex, int columnIndex, int rowSpan, int colSpan, TableCellProperties? properties)
		{
			Row = row;
			RowIndex = rowIndex;
			ColumnIndex = columnIndex;
			RowSpan = rowSpan;
			ColSpan = colSpan;
			Properties = properties;
		}
	}
}
