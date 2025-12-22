using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using Microsoft.Extensions.Logging;
using System.Xml.Linq;

namespace DocxportNet.Visitors;

public class DxpStyleVisitor : DxpIStyleVisitor
{
	public virtual void StyleBoldBegin(DxpIDocumentContext d) { }
	public virtual void StyleBoldEnd(DxpIDocumentContext d) { }

	public virtual void StyleItalicBegin(DxpIDocumentContext d) { }
	public virtual void StyleItalicEnd(DxpIDocumentContext d) { }

	public virtual void StyleUnderlineBegin(DxpIDocumentContext d) { }
	public virtual void StyleUnderlineEnd(DxpIDocumentContext d) { }

	public virtual void StyleStrikeBegin(DxpIDocumentContext d) { }
	public virtual void StyleStrikeEnd(DxpIDocumentContext d) { }

	public virtual void StyleDoubleStrikeBegin(DxpIDocumentContext d) { }
	public virtual void StyleDoubleStrikeEnd(DxpIDocumentContext d) { }

	public virtual void StyleSuperscriptBegin(DxpIDocumentContext d) { }
	public virtual void StyleSuperscriptEnd(DxpIDocumentContext d) { }

	public virtual void StyleSubscriptBegin(DxpIDocumentContext d) { }
	public virtual void StyleSubscriptEnd(DxpIDocumentContext d) { }

	public virtual void StyleSmallCapsBegin(DxpIDocumentContext d) { }
	public virtual void StyleSmallCapsEnd(DxpIDocumentContext d) { }

	public virtual void StyleAllCapsBegin(DxpIDocumentContext d) { }
	public virtual void StyleAllCapsEnd(DxpIDocumentContext d) { }

	public virtual void StyleFontBegin(DxpFont font, DxpIDocumentContext d) { }
	public virtual void StyleFontEnd(DxpIDocumentContext d) { }
}



public class DxpVisitor : DxpStyleVisitor, DxpIVisitor
{
	private readonly ILogger? _logger;
	public DxpVisitor(ILogger? logger)
	{
		_logger = logger;
	}


	private void Ignored(string method, object? element = null)
	{
		// Logs the method name; if an element is provided, include its runtime type (and OpenXml local name when available).
		string elem =
			element is null ? "" :
			element is OpenXmlElement oxe ? $" ({oxe.GetType().Name}, LocalName='{oxe.LocalName}')" :
			$" ({element.GetType().Name})";

		_logger?.LogInformation($"{nameof(DxpVisitor)}: ignored {method}{elem}");
	}

	public virtual void VisitTableLayout(Table t, DxpTableModel model) => Ignored(nameof(VisitTableLayout), t);

	public virtual IDisposable VisitBlockBegin(OpenXmlElement child, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitBlockBegin), child);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitDocumentBodyBegin(Body body, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitDocumentBodyBegin), body);
		return DxpDisposable.Empty;
	}

	public virtual void VisitBookmarkEnd(BookmarkEnd be, DxpIDocumentContext d) => Ignored(nameof(VisitBookmarkEnd), be);
	public virtual void VisitBookmarkStart(BookmarkStart bs, DxpIDocumentContext d) => Ignored(nameof(VisitBookmarkStart), bs);
	public virtual void VisitBreak(Break br, DxpIDocumentContext d) => Ignored(nameof(VisitBreak), br);
	public virtual void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d) => Ignored(nameof(VisitCarriageReturn), cr);

	public virtual IDisposable VisitDeletedRunBegin(DeletedRun dr, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitDeletedRunBegin), dr);
		return DxpDisposable.Empty;
	}

	public virtual void VisitDeletedText(DeletedText dt, DxpIDocumentContext d) => Ignored(nameof(VisitDeletedText), dt);

	public virtual IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitHyperlinkBegin), link);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitInsertedRunBegin(InsertedRun ir, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitInsertedRunBegin), ir);
		return DxpDisposable.Empty;
	}

	public virtual void VisitLastRenderedPageBreak(LastRenderedPageBreak pb, DxpIDocumentContext d) => Ignored(nameof(VisitLastRenderedPageBreak), pb);
	public virtual void VisitNoBreakHyphen(NoBreakHyphen h, DxpIDocumentContext d) => Ignored(nameof(VisitNoBreakHyphen), h);

	public virtual IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
	{
		Ignored(nameof(VisitParagraphBegin), p);
		return DxpDisposable.Empty;
	}

	public virtual void VisitProofError(ProofError pe, DxpIDocumentContext d) => Ignored(nameof(VisitProofError), pe);

	public virtual IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitRunBegin), r);
		return DxpDisposable.Empty;
	}
	public virtual void VisitTab(TabChar tab, DxpIDocumentContext d) => Ignored(nameof(VisitTab), tab);

	public virtual IDisposable VisitTableBegin(Table t, DxpTableModel model, DxpIDocumentContext d, DxpITableContext table)
	{
		Ignored(nameof(VisitTableBegin), t);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitTableCellBegin(TableCell tc, DxpITableCellContext cell, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitTableCellBegin), tc);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitTableRowBegin(TableRow tr, DxpITableRowContext row, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitTableRowBegin), tr);
		return DxpDisposable.Empty;
	}

	public virtual void VisitText(Text t, DxpIDocumentContext d) => Ignored(nameof(VisitText), t);

	public virtual void VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d) => Ignored(nameof(VisitDrawingBegin), d);

	public virtual void VisitFootnoteReference(FootnoteReference fr, DxpIFootnoteContext footnote, DxpIDocumentContext d) => Ignored(nameof(VisitFootnoteReference), fr);

	public virtual IDisposable VisitFootnoteBegin(Footnote fn, DxpIFootnoteContext footnote, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitFootnoteBegin), fn);
		return DxpDisposable.Empty;
	}

	public virtual void VisitFootnoteReferenceMark(FootnoteReferenceMark m, DxpIFootnoteContext footnote, DxpIDocumentContext d) => Ignored(nameof(VisitFootnoteReferenceMark), m);

	public virtual IDisposable VisitCommentThreadBegin(string anchorId, DxpCommentThread thread, DxpIDocumentContext d)
	{
		Ignored($"{nameof(DxpVisitor)}: ignored {nameof(VisitCommentThreadBegin)} (anchorId='{anchorId}', count={thread.Comments.Count})");
		return DxpDisposable.Empty;
	}

	public virtual void VisitDayShort(DayShort ds, DxpIDocumentContext d) => Ignored(nameof(VisitDayShort), ds);
	public virtual void VisitMonthShort(MonthShort ms, DxpIDocumentContext d) => Ignored(nameof(VisitMonthShort), ms);
	public virtual void VisitYearShort(YearShort ys, DxpIDocumentContext d) => Ignored(nameof(VisitYearShort), ys);

	public virtual void VisitDayLong(DayLong dl, DxpIDocumentContext d) => Ignored(nameof(VisitDayLong), dl);
	public virtual void VisitMonthLong(MonthLong ml, DxpIDocumentContext d) => Ignored(nameof(VisitMonthLong), ml);
	public virtual void VisitYearLong(YearLong yl, DxpIDocumentContext d) => Ignored(nameof(VisitYearLong), yl);

	public virtual void VisitPageNumber(PageNumber pn, DxpIDocumentContext d) => Ignored(nameof(VisitPageNumber), pn);
	public virtual void VisitAnnotationReference(AnnotationReferenceMark arm, DxpIDocumentContext d) => Ignored(nameof(VisitAnnotationReference), arm);

	public virtual void VisitEndnoteReferenceMark(EndnoteReferenceMark erm, DxpIDocumentContext d) => Ignored(nameof(VisitEndnoteReferenceMark), erm);
	public virtual void VisitEndnoteReference(EndnoteReference enr, DxpIDocumentContext d) => Ignored(nameof(VisitEndnoteReference), enr);

	public virtual void VisitSeparatorMark(SeparatorMark sep, DxpIDocumentContext d) => Ignored(nameof(VisitSeparatorMark), sep);
	public virtual void VisitContinuationSeparatorMark(ContinuationSeparatorMark csep, DxpIDocumentContext d) => Ignored(nameof(VisitContinuationSeparatorMark), csep);

	public virtual void VisitSoftHyphen(SoftHyphen sh, DxpIDocumentContext d) => Ignored(nameof(VisitSoftHyphen), sh);
	public virtual void VisitSymbol(SymbolChar sym, DxpIDocumentContext d) => Ignored(nameof(VisitSymbol), sym);
	public virtual void VisitPositionalTab(PositionalTab ptab, DxpIDocumentContext d) => Ignored(nameof(VisitPositionalTab), ptab);

	public virtual IDisposable VisitRubyBegin(Ruby ruby, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitRubyBegin), ruby);
		return DxpDisposable.Empty;
	}

	public virtual void VisitDeletedFieldCode(DeletedFieldCode dfc, DxpIDocumentContext d) => Ignored(nameof(VisitDeletedFieldCode), dfc);
	public virtual void VisitEmbeddedObject(EmbeddedObject obj, DxpIDocumentContext d) => Ignored(nameof(VisitEmbeddedObject), obj);

	public virtual void VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d) => Ignored(nameof(VisitLegacyPictureBegin), pict);
	public virtual IDisposable VisitRubyContentBegin(RubyContentType rc, bool isBase, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitRubyContentBegin), rc);
		return DxpDisposable.Empty;
	}

	public virtual void VisitPermStart(PermStart ps, DxpIDocumentContext d) => Ignored(nameof(VisitPermStart), ps);
	public virtual void VisitPermEnd(PermEnd pe2, DxpIDocumentContext d) => Ignored(nameof(VisitPermEnd), pe2);

	public virtual void VisitMoveFromRangeStart(MoveFromRangeStart mfrs, DxpIDocumentContext d) => Ignored(nameof(VisitMoveFromRangeStart), mfrs);
	public virtual void VisitMoveFromRangeEnd(MoveFromRangeEnd mfre, DxpIDocumentContext d) => Ignored(nameof(VisitMoveFromRangeEnd), mfre);
	public virtual void VisitMoveToRangeStart(MoveToRangeStart mtrs, DxpIDocumentContext d) => Ignored(nameof(VisitMoveToRangeStart), mtrs);
	public virtual void VisitMoveToRangeEnd(MoveToRangeEnd mtre, DxpIDocumentContext d) => Ignored(nameof(VisitMoveToRangeEnd), mtre);

	public virtual IDisposable VisitInsertedBegin(Inserted ins, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitInsertedBegin), ins);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitDeletedBegin(Deleted del, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitDeletedBegin), del);
		return DxpDisposable.Empty;
	}

	public virtual void VisitOMathParagraph(DocumentFormat.OpenXml.Math.Paragraph oMathPara, DxpIDocumentContext d) => Ignored(nameof(VisitOMathParagraph), oMathPara);
	public virtual void VisitOMath(DocumentFormat.OpenXml.Math.OfficeMath oMath, DxpIDocumentContext d) => Ignored(nameof(VisitOMath), oMath);

	public virtual void VisitDeletedTableRowMark(Deleted del, TableRowProperties trPr, TableRow? tr, DxpIDocumentContext d) => Ignored(nameof(VisitDeletedTableRowMark), del);
	public virtual void VisitDeletedParagraphMark(Deleted del, ParagraphProperties pPr, Paragraph? p, DxpIDocumentContext d) => Ignored(nameof(VisitDeletedParagraphMark), del);

	public virtual void VisitInsertedParagraphMark(Inserted ins, ParagraphProperties pPr2, Paragraph? p, DxpIDocumentContext d) => Ignored(nameof(VisitInsertedParagraphMark), ins);

	public virtual void VisitInsertedNumbering(Inserted ins, DxpMarker? marker, DxpStyleEffectiveIndentTwips indent, Paragraph? p, DxpIDocumentContext d) => Ignored(nameof(VisitInsertedNumbering), ins);

	public virtual void VisitInsertedTableRowMark(Inserted ins, TableRowProperties trPr, TableRow? tr, DxpIDocumentContext d) => Ignored(nameof(VisitInsertedTableRowMark), ins);

	public virtual void VisitCustomXmlInsRangeStart(CustomXmlInsRangeStart cins, DxpIDocumentContext d) => Ignored(nameof(VisitCustomXmlInsRangeStart), cins);
	public virtual void VisitCustomXmlInsRangeEnd(CustomXmlInsRangeEnd cine, DxpIDocumentContext d) => Ignored(nameof(VisitCustomXmlInsRangeEnd), cine);

	public virtual void VisitCustomXmlDelRangeStart(CustomXmlDelRangeStart cdls, DxpIDocumentContext d) => Ignored(nameof(VisitCustomXmlDelRangeStart), cdls);
	public virtual void VisitCustomXmlDelRangeEnd(CustomXmlDelRangeEnd cdle, DxpIDocumentContext d) => Ignored(nameof(VisitCustomXmlDelRangeEnd), cdle);

	public virtual void VisitCustomXmlMoveFromRangeStart(CustomXmlMoveFromRangeStart cmfs, DxpIDocumentContext d) => Ignored(nameof(VisitCustomXmlMoveFromRangeStart), cmfs);
	public virtual void VisitCustomXmlMoveFromRangeEnd(CustomXmlMoveFromRangeEnd cmfe, DxpIDocumentContext d) => Ignored(nameof(VisitCustomXmlMoveFromRangeEnd), cmfe);

	public virtual void VisitCustomXmlMoveToRangeStart(CustomXmlMoveToRangeStart cmts, DxpIDocumentContext d) => Ignored(nameof(VisitCustomXmlMoveToRangeStart), cmts);
	public virtual void VisitCustomXmlMoveToRangeEnd(CustomXmlMoveToRangeEnd cmte, DxpIDocumentContext d) => Ignored(nameof(VisitCustomXmlMoveToRangeEnd), cmte);

	public virtual IDisposable VisitSdtBlockBegin(SdtBlock sdt, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitSdtBlockBegin), sdt);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitCustomXmlBlockBegin(CustomXmlBlock cx, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitCustomXmlBlockBegin), cx);
		return DxpDisposable.Empty;
	}

	public virtual void VisitAltChunk(AltChunk ac, DxpIDocumentContext d) => Ignored(nameof(VisitAltChunk), ac);

	public virtual IDisposable VisitSdtContentBlockBegin(SdtContentBlock content, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitSdtContentBlockBegin), content);
		return DxpDisposable.Empty;
	}

	public virtual void VisitCustomXmlConflictInsertionRangeStart(CustomXmlConflictInsertionRangeStart cxCis, DxpIDocumentContext d) => Ignored(nameof(VisitCustomXmlConflictInsertionRangeStart), cxCis);
	public virtual void VisitCustomXmlConflictInsertionRangeEnd(CustomXmlConflictInsertionRangeEnd cxCie, DxpIDocumentContext d) => Ignored(nameof(VisitCustomXmlConflictInsertionRangeEnd), cxCie);

	public virtual void VisitCustomXmlConflictDeletionRangeStart(CustomXmlConflictDeletionRangeStart cxCds, DxpIDocumentContext d) => Ignored(nameof(VisitCustomXmlConflictDeletionRangeStart), cxCds);
	public virtual void VisitCustomXmlConflictDeletionRangeEnd(CustomXmlConflictDeletionRangeEnd cxCde, DxpIDocumentContext d) => Ignored(nameof(VisitCustomXmlConflictDeletionRangeEnd), cxCde);

	public virtual void VisitMoveFromRun(MoveFromRun mfr, DxpIDocumentContext d) => Ignored(nameof(VisitMoveFromRun), mfr);
	public virtual void VisitMoveToRun(MoveToRun mtr, DxpIDocumentContext d) => Ignored(nameof(VisitMoveToRun), mtr);

	public virtual void VisitContentPart(DocumentFormat.OpenXml.Wordprocessing.ContentPart cp, DxpIDocumentContext d) => Ignored(nameof(VisitContentPart), cp);

	public virtual IDisposable VisitCustomXmlRunBegin(CustomXmlRun cxr, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitCustomXmlRunBegin), cxr);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitSimpleFieldBegin), fld);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitSdtRunBegin(SdtRun sdtRun, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitSdtRunBegin), sdtRun);
		return DxpDisposable.Empty;
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Accent mAccent, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mAccent);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Bar mBar, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mBar);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Box mBox, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mBox);
	public virtual void VisitOMathRun(DocumentFormat.OpenXml.Math.Run mMathRun, DxpIDocumentContext d) => Ignored(nameof(VisitOMathRun), mMathRun);

	public virtual IDisposable VisitBidirectionalOverrideBegin(BidirectionalOverride bdo, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitBidirectionalOverrideBegin), bdo);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitBidirectionalEmbeddingBegin(BidirectionalEmbedding bdi, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitBidirectionalEmbeddingBegin), bdi);
		return DxpDisposable.Empty;
	}

	public virtual void VisitSubDocumentReference(SubDocumentReference subDoc, DxpIDocumentContext d) => Ignored(nameof(VisitSubDocumentReference), subDoc);

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.BorderBox mBorderBox, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mBorderBox);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Delimiter mDelim, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mDelim);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.EquationArray mEqArr, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mEqArr);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Fraction mFrac, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mFrac);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.MathFunction mFunc, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mFunc);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.GroupChar mGroupChr, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mGroupChr);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.LimitLower mLimLow, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mLimLow);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.LimitUpper mLimUpp, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mLimUpp);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Matrix mMat, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mMat);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Nary mNary, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mNary);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Phantom mPhant, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mPhant);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Radical mRad, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mRad);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.PreSubSuper mPreSubSup, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mPreSubSup);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Subscript mSub, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mSub);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.SubSuperscript mSubSup, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mSubSup);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Superscript mSup, DxpIDocumentContext d) => Ignored(nameof(VisitOMathElement), mSup);

	public virtual IDisposable VisitSdtContentRunBegin(SdtContentRun content, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitSdtContentRunBegin), content);
		return DxpDisposable.Empty;
	}

	public virtual void VisitFieldData(FieldData data, DxpIDocumentContext d) => Ignored(nameof(VisitFieldData), data);
	public virtual void VisitConflictInsertion(ConflictInsertion cIns, DxpIDocumentContext d) => Ignored(nameof(VisitConflictInsertion), cIns);
	public virtual void VisitConflictDeletion(ConflictDeletion cDel, DxpIDocumentContext d) => Ignored(nameof(VisitConflictDeletion), cDel);

	public virtual IDisposable VisitSdtRowBegin(SdtRow sdtRow, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitSdtRowBegin), sdtRow);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitCustomXmlRowBegin(CustomXmlRow cxRow, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitCustomXmlRowBegin), cxRow);
		return DxpDisposable.Empty;
	}

	public bool AcceptAlternateContentChoice(AlternateContentChoice choice, IReadOnlyList<string> required, DxpIDocumentContext d)
	{
		Ignored(nameof(AcceptAlternateContentChoice), choice);
		return false; // base behavior: don’t accept any choices
	}

	public virtual IDisposable VisitEndnoteBegin(Endnote item1, long item3, int item2, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitEndnoteBegin), item1);
		return DxpDisposable.Empty;
	}

	IDisposable DxpIVisitor.VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitDrawingBegin", drw);
		return DxpDisposable.Empty;
	}

	IDisposable DxpIVisitor.VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitLegacyPictureBegin", pict);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitSmartTagRunBegin(OpenXmlUnknownElement smart, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitSmartTagRunBegin), smart);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitTextBoxContentBegin(TextBoxContent txbx, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitTextBoxContentBegin), txbx);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitSmartTagRunBegin(OpenXmlUnknownElement smart, string elementName, string elementUri, DxpIDocumentContext d)
	{
		_logger?.LogInformation($"{nameof(DxpVisitor)}: ignored {nameof(VisitSmartTagRunBegin)} (elementName='{elementName}', elementUri='{elementUri}')");
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitAlternateContentBegin(AlternateContent ac, DxpIDocumentContext d)
	{
		Ignored(nameof(VisitAlternateContentBegin), ac);
		return DxpDisposable.Empty;
	}

	public virtual void VisitUnknown(string context, OpenXmlElement el, DxpIDocumentContext d)
	{
		_logger?.LogInformation($"{nameof(VisitUnknown)}: ignored {el.GetType().FullName} (context='{context}')");
	}

	public virtual void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitComplexFieldBegin", begin);
	}

	public virtual void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitComplexFieldInstruction", instr);
	}

	public virtual void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitComplexFieldSeparate", separate);
	}

	public virtual IDisposable VisitComplexFieldResultBegin(DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitComplexFieldResultBegin");
		return DxpDisposable.Empty;
	}

	public virtual void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitComplexFieldCachedResultText");
	}

	public virtual void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitComplexFieldEnd", end);
	}

	public virtual IDisposable VisitSdtCellBegin(SdtCell sdtCell, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitSdtCellBegin", sdtCell);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitCustomXmlCellBegin(CustomXmlCell cxCell, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitCustomXmlCellBegin", cxCell);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitSectionHeaderBegin(Header hdr, object value, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitSectionHeaderBegin", hdr);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitSectionFooterBegin(Footer ftr, object value, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitSectionFooterBegin", ftr);
		return DxpDisposable.Empty;
	}

	public virtual void VisitDocumentProperties(IPackageProperties core, IReadOnlyList<CustomFileProperty> custom, IReadOnlyList<DxpTimelineEvent> timeline, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitDocumentProperties", core);
	}

	public virtual void VisitBibliographySources(CustomXmlPart bibliographyPart, XDocument bib, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitBibliographySources", bibliographyPart);
	}

	public virtual IDisposable VisitSectionBegin(SectionProperties properties, SectionLayout layout, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitSectionBegin", properties);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitSectionBodyBegin(SectionProperties properties, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitSectionBodyBegin", properties);
		return DxpDisposable.Empty;
	}

	public virtual IDisposable VisitCommentBegin(DxpCommentInfo c, DxpCommentThread thread, DxpIDocumentContext d)
	{
		Ignored("IDocxVisitor.VisitCommentBegin", c);
		return DxpDisposable.Empty;
	}
}
