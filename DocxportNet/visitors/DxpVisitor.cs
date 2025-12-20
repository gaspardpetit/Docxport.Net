using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.api;
using DocxportNet.walker;
using Microsoft.Extensions.Logging;
using static DocxportNet.walker.DxpWalker;

namespace DocxportNet.visitors;

public class DxpStyleVisitor : IDxpStyleVisitor
{
	public virtual void StyleBoldBegin() { }
	public virtual void StyleBoldEnd() { }

	public virtual void StyleItalicBegin() { }
	public virtual void StyleItalicEnd() { }

	public virtual void StyleUnderlineBegin() { }
	public virtual void StyleUnderlineEnd() { }

	public virtual void StyleStrikeBegin() { }
	public virtual void StyleStrikeEnd() { }

	public virtual void StyleDoubleStrikeBegin() { }
	public virtual void StyleDoubleStrikeEnd() { }

	public virtual void StyleSuperscriptBegin() { }
	public virtual void StyleSuperscriptEnd() { }

	public virtual void StyleSubscriptBegin() { }
	public virtual void StyleSubscriptEnd() { }

	public virtual void StyleSmallCapsBegin() { }
	public virtual void StyleSmallCapsEnd() { }

	public virtual void StyleAllCapsBegin() { }
	public virtual void StyleAllCapsEnd() { }

	public virtual void StyleFontBegin(string? fontName, int? fontSizeHalfPoints) { }
	public virtual void StyleFontEnd() { }
}



public class DxpVisitor : DxpStyleVisitor, IDxpVisitor
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

	public virtual void VisitTableCellLayout(TableCell tc, int row, int col, int rowSpan, int colSpan) => Ignored(nameof(VisitTableCellLayout), tc);
	public virtual void VisitTableLayout(Table t, DxpTableModel model) => Ignored(nameof(VisitTableLayout), t);

	public virtual IDisposable VisitBlockBegin(OpenXmlElement child, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitBlockBegin), child);
		return Disposable.Empty;
	}

	public virtual void VisitTableRowProperties(TableRowProperties trp, IDxpStyleResolver s) => Ignored(nameof(VisitTableRowProperties), trp);

	public virtual IDisposable VisitBodyBegin(Body body, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitBodyBegin), body);
		return Disposable.Empty;
	}

	public virtual void VisitBookmarkEnd(BookmarkEnd be, IDxpStyleResolver s) => Ignored(nameof(VisitBookmarkEnd), be);
	public virtual void VisitBookmarkStart(BookmarkStart bs, IDxpStyleResolver s) => Ignored(nameof(VisitBookmarkStart), bs);
	public virtual void VisitBreak(Break br, IDxpStyleResolver s) => Ignored(nameof(VisitBreak), br);
	public virtual void VisitCarriageReturn(CarriageReturn cr, IDxpStyleResolver s) => Ignored(nameof(VisitCarriageReturn), cr);

	public virtual IDisposable VisitDeletedRunBegin(DeletedRun dr, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitDeletedRunBegin), dr);
		return Disposable.Empty;
	}

	public virtual void VisitDeletedText(DeletedText dt, IDxpStyleResolver s) => Ignored(nameof(VisitDeletedText), dt);

	public virtual IDisposable VisitHyperlinkBegin(Hyperlink link, string? target, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitHyperlinkBegin), link);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitInsertedRunBegin(InsertedRun ir, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitInsertedRunBegin), ir);
		return Disposable.Empty;
	}

	public virtual void VisitLastRenderedPageBreak(LastRenderedPageBreak pb, IDxpStyleResolver s) => Ignored(nameof(VisitLastRenderedPageBreak), pb);
	public virtual void VisitNoBreakHyphen(NoBreakHyphen h, IDxpStyleResolver s) => Ignored(nameof(VisitNoBreakHyphen), h);

	public virtual IDisposable VisitParagraphBegin(Paragraph p, IDxpStyleResolver s, string? marker, int? numId, int? iLvl, DxpStyleEffectiveIndentTwips indent)
	{
		Ignored(nameof(VisitParagraphBegin), p);
		return Disposable.Empty;
	}

	public virtual void VisitParagraphProperties(ParagraphProperties pp, IDxpStyleResolver s) => Ignored(nameof(VisitParagraphProperties), pp);
	public virtual void VisitProofError(ProofError pe, IDxpStyleResolver s) => Ignored(nameof(VisitProofError), pe);

	public virtual IDisposable VisitRunBegin(Run r, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitRunBegin), r);
		return Disposable.Empty;
	}

	public virtual void VisitRunProperties(RunProperties rp, IDxpStyleResolver s) => Ignored(nameof(VisitRunProperties), rp);
	public virtual void VisitSectionProperties(SectionProperties sp, IDxpStyleResolver s) => Ignored(nameof(VisitSectionProperties), sp);
	public virtual void VisitTab(TabChar tab, IDxpStyleResolver s) => Ignored(nameof(VisitTab), tab);

	public virtual IDisposable VisitTableBegin(Table t, DxpTableModel model, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitTableBegin), t);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitTableCellBegin(TableCell tc, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitTableCellBegin), tc);
		return Disposable.Empty;
	}

	public virtual void VisitTableCellProperties(TableCellProperties tcp, IDxpStyleResolver s) => Ignored(nameof(VisitTableCellProperties), tcp);
	public virtual void VisitTableGrid(TableGrid tg, IDxpStyleResolver s) => Ignored(nameof(VisitTableGrid), tg);
	public virtual void VisitTableProperties(TableProperties tp, IDxpStyleResolver s) => Ignored(nameof(VisitTableProperties), tp);

	public virtual IDisposable VisitTableRowBegin(TableRow tr, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitTableRowBegin), tr);
		return Disposable.Empty;
	}

	public virtual void VisitText(Text t, IDxpStyleResolver s) => Ignored(nameof(VisitText), t);

	public virtual void VisitDrawingBegin(Drawing d, DxpDrawingInfo? info, IDxpStyleResolver s) => Ignored(nameof(VisitDrawingBegin), d);

	public virtual void VisitFootnoteReference(FootnoteReference fr, long id, int index, IDxpStyleResolver s) => Ignored(nameof(VisitFootnoteReference), fr);

	public virtual IDisposable VisitFootnoteBegin(Footnote fn, long id, int index, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitFootnoteBegin), fn);
		return Disposable.Empty;
	}

		public virtual void VisitFootnoteReferenceMark(FootnoteReferenceMark m, long? footnoteId, int index, IDxpStyleResolver s) => Ignored(nameof(VisitFootnoteReferenceMark), m);
	public virtual void VisitCommentThread(string anchorId, DxpCommentThread thread, IDxpStyleResolver s)
		=> _logger?.LogInformation($"{nameof(DxpVisitor)}: ignored {nameof(VisitCommentThread)} (anchorId='{anchorId}', count={thread.Comments.Count})");

	// Visitor defaults log ignored and return Disposable.Empty when needed


	public virtual void VisitDayShort(DayShort ds, IDxpStyleResolver s) => Ignored(nameof(VisitDayShort), ds);
	public virtual void VisitMonthShort(MonthShort ms, IDxpStyleResolver s) => Ignored(nameof(VisitMonthShort), ms);
	public virtual void VisitYearShort(YearShort ys, IDxpStyleResolver s) => Ignored(nameof(VisitYearShort), ys);

	public virtual void VisitDayLong(DayLong dl, IDxpStyleResolver s) => Ignored(nameof(VisitDayLong), dl);
	public virtual void VisitMonthLong(MonthLong ml, IDxpStyleResolver s) => Ignored(nameof(VisitMonthLong), ml);
	public virtual void VisitYearLong(YearLong yl, IDxpStyleResolver s) => Ignored(nameof(VisitYearLong), yl);

	public virtual void VisitPageNumber(PageNumber pn, IDxpStyleResolver s) => Ignored(nameof(VisitPageNumber), pn);
	public virtual void VisitAnnotationReference(AnnotationReferenceMark arm, IDxpStyleResolver s) => Ignored(nameof(VisitAnnotationReference), arm);

	public virtual void VisitEndnoteReferenceMark(EndnoteReferenceMark erm, IDxpStyleResolver s) => Ignored(nameof(VisitEndnoteReferenceMark), erm);
	public virtual void VisitEndnoteReference(EndnoteReference enr, IDxpStyleResolver s) => Ignored(nameof(VisitEndnoteReference), enr);

	public virtual void VisitSeparatorMark(SeparatorMark sep, IDxpStyleResolver s) => Ignored(nameof(VisitSeparatorMark), sep);
	public virtual void VisitContinuationSeparatorMark(ContinuationSeparatorMark csep, IDxpStyleResolver s) => Ignored(nameof(VisitContinuationSeparatorMark), csep);

	public virtual void VisitSoftHyphen(SoftHyphen sh, IDxpStyleResolver s) => Ignored(nameof(VisitSoftHyphen), sh);
	public virtual void VisitSymbol(SymbolChar sym, IDxpStyleResolver s) => Ignored(nameof(VisitSymbol), sym);
	public virtual void VisitPositionalTab(PositionalTab ptab, IDxpStyleResolver s) => Ignored(nameof(VisitPositionalTab), ptab);

	public virtual IDisposable VisitRubyBegin(Ruby ruby, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitRubyBegin), ruby);
		return Disposable.Empty;
	}

	public virtual void VisitDeletedFieldCode(DeletedFieldCode dfc, IDxpStyleResolver s) => Ignored(nameof(VisitDeletedFieldCode), dfc);
	public virtual void VisitEmbeddedObject(EmbeddedObject obj, IDxpStyleResolver s) => Ignored(nameof(VisitEmbeddedObject), obj);

	public virtual void VisitLegacyPictureBegin(Picture pict, IDxpStyleResolver s) => Ignored(nameof(VisitLegacyPictureBegin), pict);
	public virtual void VisitRubyProperties(RubyProperties pr, IDxpStyleResolver s) => Ignored(nameof(VisitRubyProperties), pr);

	public virtual IDisposable VisitRubyContentBegin(RubyContentType rc, bool isBase, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitRubyContentBegin), rc);
		return Disposable.Empty;
	}

	public virtual void VisitPermStart(PermStart ps, IDxpStyleResolver s) => Ignored(nameof(VisitPermStart), ps);
	public virtual void VisitPermEnd(PermEnd pe2, IDxpStyleResolver s) => Ignored(nameof(VisitPermEnd), pe2);

	public virtual void VisitMoveFromRangeStart(MoveFromRangeStart mfrs, IDxpStyleResolver s) => Ignored(nameof(VisitMoveFromRangeStart), mfrs);
	public virtual void VisitMoveFromRangeEnd(MoveFromRangeEnd mfre, IDxpStyleResolver s) => Ignored(nameof(VisitMoveFromRangeEnd), mfre);
	public virtual void VisitMoveToRangeStart(MoveToRangeStart mtrs, IDxpStyleResolver s) => Ignored(nameof(VisitMoveToRangeStart), mtrs);
	public virtual void VisitMoveToRangeEnd(MoveToRangeEnd mtre, IDxpStyleResolver s) => Ignored(nameof(VisitMoveToRangeEnd), mtre);

	public virtual IDisposable VisitInsertedBegin(Inserted ins, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitInsertedBegin), ins);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitDeletedBegin(Deleted del, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitDeletedBegin), del);
		return Disposable.Empty;
	}

	public virtual void VisitOMathParagraph(DocumentFormat.OpenXml.Math.Paragraph oMathPara, IDxpStyleResolver s) => Ignored(nameof(VisitOMathParagraph), oMathPara);
	public virtual void VisitOMath(DocumentFormat.OpenXml.Math.OfficeMath oMath, IDxpStyleResolver s) => Ignored(nameof(VisitOMath), oMath);

	public virtual void VisitDeletedTableRowMark(Deleted del, TableRowProperties trPr, TableRow? tr, IDxpStyleResolver s) => Ignored(nameof(VisitDeletedTableRowMark), del);
	public virtual void VisitDeletedParagraphMark(Deleted del, ParagraphProperties pPr, Paragraph? p, IDxpStyleResolver s) => Ignored(nameof(VisitDeletedParagraphMark), del);

	public virtual void VisitInsertedParagraphMark(Inserted ins, ParagraphProperties pPr2, Paragraph? p, IDxpStyleResolver s) => Ignored(nameof(VisitInsertedParagraphMark), ins);

	public virtual void VisitInsertedNumberingProperties(Inserted ins, NumberingProperties numPr, ParagraphProperties? pPr, Paragraph? p, IDxpStyleResolver s) => Ignored(nameof(VisitInsertedNumberingProperties), ins);

	public virtual void VisitInsertedTableRowMark(Inserted ins, TableRowProperties trPr, TableRow? tr, IDxpStyleResolver s) => Ignored(nameof(VisitInsertedTableRowMark), ins);

	public virtual void VisitCustomXmlInsRangeStart(CustomXmlInsRangeStart cins, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlInsRangeStart), cins);
	public virtual void VisitCustomXmlInsRangeEnd(CustomXmlInsRangeEnd cine, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlInsRangeEnd), cine);

	public virtual void VisitCustomXmlDelRangeStart(CustomXmlDelRangeStart cdls, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlDelRangeStart), cdls);
	public virtual void VisitCustomXmlDelRangeEnd(CustomXmlDelRangeEnd cdle, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlDelRangeEnd), cdle);

	public virtual void VisitCustomXmlMoveFromRangeStart(CustomXmlMoveFromRangeStart cmfs, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlMoveFromRangeStart), cmfs);
	public virtual void VisitCustomXmlMoveFromRangeEnd(CustomXmlMoveFromRangeEnd cmfe, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlMoveFromRangeEnd), cmfe);

	public virtual void VisitCustomXmlMoveToRangeStart(CustomXmlMoveToRangeStart cmts, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlMoveToRangeStart), cmts);
	public virtual void VisitCustomXmlMoveToRangeEnd(CustomXmlMoveToRangeEnd cmte, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlMoveToRangeEnd), cmte);

	public virtual IDisposable VisitSdtBlockBegin(SdtBlock sdt, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitSdtBlockBegin), sdt);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitCustomXmlBlockBegin(CustomXmlBlock cx, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitCustomXmlBlockBegin), cx);
		return Disposable.Empty;
	}

	public virtual void VisitAltChunk(AltChunk ac, IDxpStyleResolver s) => Ignored(nameof(VisitAltChunk), ac);

	public virtual void VisitCommentRangeStart(CommentRangeStart crs, IDxpStyleResolver s) => Ignored(nameof(VisitCommentRangeStart), crs);
	public virtual void VisitCommentRangeEnd(CommentRangeEnd cre, IDxpStyleResolver s) => Ignored(nameof(VisitCommentRangeEnd), cre);

	public virtual void VisitSdtProperties(SdtProperties pr, IDxpStyleResolver s) => Ignored(nameof(VisitSdtProperties), pr);

	public virtual IDisposable VisitSdtContentBlockBegin(SdtContentBlock content, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitSdtContentBlockBegin), content);
		return Disposable.Empty;
	}

	public virtual void VisitCustomXmlProperties(CustomXmlProperties pr, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlProperties), pr);

	public virtual void VisitCustomXmlConflictInsertionRangeStart(CustomXmlConflictInsertionRangeStart cxCis, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlConflictInsertionRangeStart), cxCis);
	public virtual void VisitCustomXmlConflictInsertionRangeEnd(CustomXmlConflictInsertionRangeEnd cxCie, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlConflictInsertionRangeEnd), cxCie);

	public virtual void VisitCustomXmlConflictDeletionRangeStart(CustomXmlConflictDeletionRangeStart cxCds, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlConflictDeletionRangeStart), cxCds);
	public virtual void VisitCustomXmlConflictDeletionRangeEnd(CustomXmlConflictDeletionRangeEnd cxCde, IDxpStyleResolver s) => Ignored(nameof(VisitCustomXmlConflictDeletionRangeEnd), cxCde);

	public virtual void VisitMoveFromRun(MoveFromRun mfr, IDxpStyleResolver s) => Ignored(nameof(VisitMoveFromRun), mfr);
	public virtual void VisitMoveToRun(MoveToRun mtr, IDxpStyleResolver s) => Ignored(nameof(VisitMoveToRun), mtr);

	public virtual void VisitContentPart(DocumentFormat.OpenXml.Wordprocessing.ContentPart cp, IDxpStyleResolver s) => Ignored(nameof(VisitContentPart), cp);

	public virtual IDisposable VisitCustomXmlRunBegin(CustomXmlRun cxr, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitCustomXmlRunBegin), cxr);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitSimpleFieldBegin(SimpleField fld, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitSimpleFieldBegin), fld);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitSdtRunBegin(SdtRun sdtRun, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitSdtRunBegin), sdtRun);
		return Disposable.Empty;
	}

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Accent mAccent, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mAccent);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Bar mBar, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mBar);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Box mBox, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mBox);
	public virtual void VisitOMathRun(DocumentFormat.OpenXml.Math.Run mMathRun, IDxpStyleResolver s) => Ignored(nameof(VisitOMathRun), mMathRun);

	public virtual IDisposable VisitBidirectionalOverrideBegin(BidirectionalOverride bdo, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitBidirectionalOverrideBegin), bdo);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitBidirectionalEmbeddingBegin(BidirectionalEmbedding bdi, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitBidirectionalEmbeddingBegin), bdi);
		return Disposable.Empty;
	}

	public virtual void VisitSubDocumentReference(SubDocumentReference subDoc, IDxpStyleResolver s) => Ignored(nameof(VisitSubDocumentReference), subDoc);

	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.BorderBox mBorderBox, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mBorderBox);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Delimiter mDelim, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mDelim);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.EquationArray mEqArr, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mEqArr);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Fraction mFrac, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mFrac);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.MathFunction mFunc, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mFunc);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.GroupChar mGroupChr, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mGroupChr);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.LimitLower mLimLow, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mLimLow);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.LimitUpper mLimUpp, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mLimUpp);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Matrix mMat, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mMat);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Nary mNary, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mNary);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Phantom mPhant, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mPhant);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Radical mRad, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mRad);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.PreSubSuper mPreSubSup, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mPreSubSup);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Subscript mSub, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mSub);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.SubSuperscript mSubSup, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mSubSup);
	public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Superscript mSup, IDxpStyleResolver s) => Ignored(nameof(VisitOMathElement), mSup);

	public virtual void VisitSdtEndCharProperties(SdtEndCharProperties endPr, IDxpStyleResolver s) => Ignored(nameof(VisitSdtEndCharProperties), endPr);

	public virtual IDisposable VisitSdtContentRunBegin(SdtContentRun content, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitSdtContentRunBegin), content);
		return Disposable.Empty;
	}

	public virtual void VisitFieldData(FieldData data, IDxpStyleResolver s) => Ignored(nameof(VisitFieldData), data);
	public virtual void VisitConflictInsertion(ConflictInsertion cIns, IDxpStyleResolver s) => Ignored(nameof(VisitConflictInsertion), cIns);
	public virtual void VisitConflictDeletion(ConflictDeletion cDel, IDxpStyleResolver s) => Ignored(nameof(VisitConflictDeletion), cDel);

	public virtual IDisposable VisitSdtRowBegin(SdtRow sdtRow, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitSdtRowBegin), sdtRow);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitCustomXmlRowBegin(CustomXmlRow cxRow, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitCustomXmlRowBegin), cxRow);
		return Disposable.Empty;
	}

	public bool AcceptAlternateContentChoice(AlternateContentChoice choice, IReadOnlyList<string> required, IDxpStyleResolver s)
	{
		Ignored(nameof(AcceptAlternateContentChoice), choice);
		return false; // base behavior: don’t accept any choices
	}

	public bool SupportsNamespaces(IReadOnlyList<string> mu)
	{
		Ignored(nameof(SupportsNamespaces));
		return false; // base behavior: no markup-compat namespaces supported
	}

	public virtual IDisposable VisitEndnoteBegin(Endnote item1, long item3, int item2, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitEndnoteBegin), item1);
		return Disposable.Empty;
	}

	IDisposable IDxpVisitor.VisitDrawingBegin(Drawing d, DxpDrawingInfo? info, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitDrawingBegin", d);
		return Disposable.Empty;
	}

	IDisposable IDxpVisitor.VisitLegacyPictureBegin(Picture pict, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitLegacyPictureBegin", pict);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitSmartTagRunBegin(OpenXmlUnknownElement smart, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitSmartTagRunBegin), smart);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitTextBoxContentBegin(TextBoxContent txbx, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitTextBoxContentBegin), txbx);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitSmartTagRunBegin(OpenXmlUnknownElement smart, string elementName, string elementUri, IDxpStyleResolver s)
	{
		_logger?.LogInformation($"{nameof(DxpVisitor)}: ignored {nameof(VisitSmartTagRunBegin)} (elementName='{elementName}', elementUri='{elementUri}')");
		return Disposable.Empty;
	}

	public virtual void VisitSmartTagProperties(OpenXmlUnknownElement smartTagPr, List<CustomXmlAttribute> attrs, IDxpStyleResolver s) => Ignored(nameof(VisitSmartTagProperties), smartTagPr);

	public virtual IDisposable VisitAlternateContentBegin(AlternateContent ac, IDxpStyleResolver s)
	{
		Ignored(nameof(VisitAlternateContentBegin), ac);
		return Disposable.Empty;
	}

	public virtual void VisitUnknown(string context, OpenXmlElement el, IDxpStyleResolver s)
	{
		_logger?.LogInformation($"{nameof(VisitUnknown)}: ignored {el.GetType().FullName} (context='{context}')");
	}

	public virtual void VisitComplexFieldBegin(FieldChar begin, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitComplexFieldBegin", begin);
	}

	public virtual void VisitComplexFieldInstruction(FieldCode instr, string text, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitComplexFieldInstruction", instr);
	}

	public virtual void VisitComplexFieldSeparate(FieldChar separate, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitComplexFieldSeparate", separate);
	}

	public virtual IDisposable VisitComplexFieldResultBegin(IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitComplexFieldResultBegin");
		return Disposable.Empty;
	}

	public virtual void VisitComplexFieldEnd(FieldChar end, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitComplexFieldEnd", end);
	}

	public virtual IDisposable VisitSdtCellBegin(SdtCell sdtCell, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitSdtCellBegin", sdtCell);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitCustomXmlCellBegin(CustomXmlCell cxCell, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitCustomXmlCellBegin", cxCell);
		return Disposable.Empty;
	}

	public virtual void VisitDocumentSettings(Settings settings, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitDocumentSettings", settings);
	}

	public virtual void VisitDocumentBackground(object background, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitDocumentBackground", background);
	}

	public virtual void VisitSectionLayout(SectionProperties sp, SectionLayout layout, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitSectionLayout", sp);
	}

	public virtual IDisposable VisitSectionHeaderBegin(Header hdr, object value, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitSectionHeaderBegin", hdr);
		return Disposable.Empty;
	}

	public virtual IDisposable VisitSectionFooterBegin(Footer ftr, object value, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitSectionFooterBegin", ftr);
		return Disposable.Empty;
	}

	public virtual void VisitCoreFileProperties(IPackageProperties core)
	{
		Ignored("IDocxVisitor.VisitCoreFileProperties", core);
	}

	public virtual void VisitCustomFileProperties(IEnumerable<CustomFileProperty> custom)
	{
		Ignored("IDocxVisitor.VisitCoreFileProperties", custom);
	}

	public virtual void VisitGlossaryDocumentBegin(GlossaryDocument gl, IDxpStyleResolver s)
	{
		Ignored("IDocxVisitor.VisitGlossaryDocumentBegin", gl);
	}

	public virtual void VisitBibliographySources(object bibliographyPart, object bib)
	{
		Ignored("IDocxVisitor.VisitBibliographySources", bibliographyPart);
	}
}
