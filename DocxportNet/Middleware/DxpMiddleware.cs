using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using System.Xml.Linq;

namespace DocxportNet.Middleware;

public abstract class DxpMiddleware : DxpIVisitor
{
    protected DxpMiddleware()
    {}

    public abstract DxpIVisitor? Next { get; }

    protected virtual bool ShouldForwardContent(DxpIDocumentContext d) => true;

    public virtual void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitComplexFieldBegin(begin, d);
    }

    public virtual void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitComplexFieldInstruction(instr, text, d);
    }

    public virtual void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitComplexFieldSeparate(separate, d);
    }

    public virtual IDisposable VisitComplexFieldResultBegin(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitComplexFieldResultBegin(d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitComplexFieldCachedResultText(text, d);
    }

    public virtual void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitComplexFieldEnd(end, d);
    }

    public virtual void StyleBoldBegin(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleBoldBegin(d);
    }

    public virtual void StyleBoldEnd(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleBoldEnd(d);
    }

    public virtual void StyleItalicBegin(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleItalicBegin(d);
    }

    public virtual void StyleItalicEnd(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleItalicEnd(d);
    }

    public virtual void StyleUnderlineBegin(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleUnderlineBegin(d);
    }

    public virtual void StyleUnderlineEnd(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleUnderlineEnd(d);
    }

    public virtual void StyleStrikeBegin(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleStrikeBegin(d);
    }

    public virtual void StyleStrikeEnd(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleStrikeEnd(d);
    }

    public virtual void StyleDoubleStrikeBegin(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleDoubleStrikeBegin(d);
    }

    public virtual void StyleDoubleStrikeEnd(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleDoubleStrikeEnd(d);
    }

    public virtual void StyleSuperscriptBegin(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleSuperscriptBegin(d);
    }

    public virtual void StyleSuperscriptEnd(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleSuperscriptEnd(d);
    }

    public virtual void StyleSubscriptBegin(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleSubscriptBegin(d);
    }

    public virtual void StyleSubscriptEnd(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleSubscriptEnd(d);
    }

    public virtual void StyleSmallCapsBegin(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleSmallCapsBegin(d);
    }

    public virtual void StyleSmallCapsEnd(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleSmallCapsEnd(d);
    }

    public virtual void StyleAllCapsBegin(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleAllCapsBegin(d);
    }

    public virtual void StyleAllCapsEnd(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleAllCapsEnd(d);
    }

    public virtual void StyleFontBegin(DxpFont font, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleFontBegin(font, d);
    }

    public virtual void StyleFontEnd(DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.StyleFontEnd(d);
    }

    public virtual void SetOutput(Stream stream)
    {
        Next?.SetOutput(stream);
    }

    public virtual bool AcceptAlternateContentChoice(AlternateContentChoice choice, IReadOnlyList<string> required, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return false;
        return Next?.AcceptAlternateContentChoice(choice, required, d) ?? false;
    }

    public virtual IDisposable VisitAlternateContentBegin(AlternateContent ac, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitAlternateContentBegin(ac, d) ?? DxpDisposable.Empty;
    }

    public virtual void VisitAltChunk(AltChunk ac, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitAltChunk(ac, d);
    }

    public virtual void VisitAnnotationReference(AnnotationReferenceMark arm, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitAnnotationReference(arm, d);
    }

    public virtual IDisposable VisitBidirectionalEmbeddingBegin(BidirectionalEmbedding bdi, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitBidirectionalEmbeddingBegin(bdi, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitBidirectionalOverrideBegin(BidirectionalOverride bdo, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitBidirectionalOverrideBegin(bdo, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitBibliographySources(CustomXmlPart bibliographyPart, XDocument bib, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitBibliographySources(bibliographyPart, bib, d);
    }

    public virtual IDisposable VisitBlockBegin(OpenXmlElement child, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitBlockBegin(child, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitBookmarkEnd(BookmarkEnd be, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitBookmarkEnd(be, d);
    }

    public virtual void VisitBookmarkStart(BookmarkStart bs, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitBookmarkStart(bs, d);
    }

    public virtual void VisitBreak(Break br, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitBreak(br, d);
    }

    public virtual void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCarriageReturn(cr, d);
    }

    public virtual IDisposable VisitCommentBegin(DxpCommentInfo c, DxpCommentThread thread, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitCommentBegin(c, thread, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitCommentThreadBegin(string anchorId, DxpCommentThread thread, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitCommentThreadBegin(anchorId, thread, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitConflictDeletion(ConflictDeletion cDel, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitConflictDeletion(cDel, d);
    }

    public virtual void VisitConflictInsertion(ConflictInsertion cIns, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitConflictInsertion(cIns, d);
    }

    public virtual void VisitContentPart(DocumentFormat.OpenXml.Wordprocessing.ContentPart cp, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitContentPart(cp, d);
    }

    public virtual void VisitContinuationSeparatorMark(ContinuationSeparatorMark csep, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitContinuationSeparatorMark(csep, d);
    }

    public virtual IDisposable VisitCustomXmlBlockBegin(CustomXmlBlock cx, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitCustomXmlBlockBegin(cx, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitCustomXmlCellBegin(CustomXmlCell cxCell, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitCustomXmlCellBegin(cxCell, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitCustomXmlConflictDeletionRangeEnd(CustomXmlConflictDeletionRangeEnd cxCde, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCustomXmlConflictDeletionRangeEnd(cxCde, d);
    }

    public virtual void VisitCustomXmlConflictDeletionRangeStart(CustomXmlConflictDeletionRangeStart cxCds, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCustomXmlConflictDeletionRangeStart(cxCds, d);
    }

    public virtual void VisitCustomXmlConflictInsertionRangeEnd(CustomXmlConflictInsertionRangeEnd cxCie, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCustomXmlConflictInsertionRangeEnd(cxCie, d);
    }

    public virtual void VisitCustomXmlConflictInsertionRangeStart(CustomXmlConflictInsertionRangeStart cxCis, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCustomXmlConflictInsertionRangeStart(cxCis, d);
    }

    public virtual void VisitCustomXmlDelRangeEnd(CustomXmlDelRangeEnd cdle, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCustomXmlDelRangeEnd(cdle, d);
    }

    public virtual void VisitCustomXmlDelRangeStart(CustomXmlDelRangeStart cdls, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCustomXmlDelRangeStart(cdls, d);
    }

    public virtual void VisitCustomXmlInsRangeEnd(CustomXmlInsRangeEnd cine, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCustomXmlInsRangeEnd(cine, d);
    }

    public virtual void VisitCustomXmlInsRangeStart(CustomXmlInsRangeStart cins, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCustomXmlInsRangeStart(cins, d);
    }

    public virtual void VisitCustomXmlMoveFromRangeEnd(CustomXmlMoveFromRangeEnd cmfe, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCustomXmlMoveFromRangeEnd(cmfe, d);
    }

    public virtual void VisitCustomXmlMoveFromRangeStart(CustomXmlMoveFromRangeStart cmfs, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCustomXmlMoveFromRangeStart(cmfs, d);
    }

    public virtual void VisitCustomXmlMoveToRangeEnd(CustomXmlMoveToRangeEnd cmte, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCustomXmlMoveToRangeEnd(cmte, d);
    }

    public virtual void VisitCustomXmlMoveToRangeStart(CustomXmlMoveToRangeStart cmts, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitCustomXmlMoveToRangeStart(cmts, d);
    }

    public virtual IDisposable VisitCustomXmlRowBegin(CustomXmlRow cxRow, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitCustomXmlRowBegin(cxRow, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitCustomXmlRunBegin(CustomXmlRun cxr, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitCustomXmlRunBegin(cxr, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitDayLong(DayLong dl, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitDayLong(dl, d);
    }

    public virtual void VisitDayShort(DayShort ds, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitDayShort(ds, d);
    }

    public virtual IDisposable VisitDeletedBegin(Deleted del, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitDeletedBegin(del, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitDeletedFieldCode(DeletedFieldCode dfc, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitDeletedFieldCode(dfc, d);
    }

    public virtual void VisitDeletedParagraphMark(Deleted del, ParagraphProperties pPr, Paragraph? p, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitDeletedParagraphMark(del, pPr, p, d);
    }

    public virtual IDisposable VisitDeletedRunBegin(DeletedRun dr, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitDeletedRunBegin(dr, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitDeletedTableRowMark(Deleted del, TableRowProperties trPr, TableRow? tr, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitDeletedTableRowMark(del, trPr, tr, d);
    }

    public virtual void VisitDeletedText(DeletedText dt, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitDeletedText(dt, d);
    }

    public virtual IDisposable VisitDocumentBodyBegin(Body body, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitDocumentBodyBegin(body, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext documentContext)
    {
        if (!ShouldForwardContent(documentContext))
            return DxpDisposable.Empty;
        return Next?.VisitDocumentBegin(doc, documentContext) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitDrawingBegin(Drawing drw, DxpDrawingInfo? info, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitDrawingBegin(drw, info, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitEmbeddedObject(EmbeddedObject obj, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitEmbeddedObject(obj, d);
    }

    public virtual IDisposable VisitEndnoteBegin(Endnote item1, long item3, int item2, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitEndnoteBegin(item1, item3, item2, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitEndnoteReference(EndnoteReference enr, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitEndnoteReference(enr, d);
    }

    public virtual void VisitEndnoteReferenceMark(EndnoteReferenceMark erm, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitEndnoteReferenceMark(erm, d);
    }

    public virtual void VisitFieldData(FieldData data, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitFieldData(data, d);
    }

    public virtual IDisposable VisitFootnoteBegin(Footnote fn, DxpIFootnoteContext footnote, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitFootnoteBegin(fn, footnote, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitFootnoteReference(FootnoteReference fr, DxpIFootnoteContext footnote, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitFootnoteReference(fr, footnote, d);
    }

    public virtual void VisitFootnoteReferenceMark(FootnoteReferenceMark m, DxpIFootnoteContext footnote, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitFootnoteReferenceMark(m, footnote, d);
    }

    public virtual IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitHyperlinkBegin(link, target, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitInsertedBegin(Inserted ins, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitInsertedBegin(ins, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitInsertedNumbering(Inserted ins, DxpMarker? marker, DxpStyleEffectiveIndentTwips indent, Paragraph? p, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitInsertedNumbering(ins, marker, indent, p, d);
    }

    public virtual void VisitInsertedParagraphMark(Inserted ins, ParagraphProperties pPr2, Paragraph? p, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitInsertedParagraphMark(ins, pPr2, p, d);
    }

    public virtual IDisposable VisitInsertedRunBegin(InsertedRun ir, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitInsertedRunBegin(ir, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitInsertedTableRowMark(Inserted ins, TableRowProperties trPr, TableRow? tr, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitInsertedTableRowMark(ins, trPr, tr, d);
    }

    public virtual void VisitLastRenderedPageBreak(LastRenderedPageBreak pb, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitLastRenderedPageBreak(pb, d);
    }

    public virtual IDisposable VisitLegacyPictureBegin(Picture pict, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitLegacyPictureBegin(pict, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitMonthLong(MonthLong ml, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitMonthLong(ml, d);
    }

    public virtual void VisitMonthShort(MonthShort ms, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitMonthShort(ms, d);
    }

    public virtual void VisitMoveFromRangeEnd(MoveFromRangeEnd mfre, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitMoveFromRangeEnd(mfre, d);
    }

    public virtual void VisitMoveFromRangeStart(MoveFromRangeStart mfrs, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitMoveFromRangeStart(mfrs, d);
    }

    public virtual void VisitMoveFromRun(MoveFromRun mfr, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitMoveFromRun(mfr, d);
    }

    public virtual void VisitMoveToRangeEnd(MoveToRangeEnd mtre, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitMoveToRangeEnd(mtre, d);
    }

    public virtual void VisitMoveToRangeStart(MoveToRangeStart mtrs, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitMoveToRangeStart(mtrs, d);
    }

    public virtual void VisitMoveToRun(MoveToRun mtr, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitMoveToRun(mtr, d);
    }

    public virtual void VisitNoBreakHyphen(NoBreakHyphen h, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitNoBreakHyphen(h, d);
    }

    public virtual void VisitOMath(DocumentFormat.OpenXml.Math.OfficeMath oMath, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMath(oMath, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Accent mAccent, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mAccent, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Bar mBar, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mBar, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.BorderBox mBorderBox, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mBorderBox, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Box mBox, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mBox, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Delimiter mDelim, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mDelim, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.EquationArray mEqArr, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mEqArr, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Fraction mFrac, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mFrac, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.GroupChar mGroupChr, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mGroupChr, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.LimitLower mLimLow, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mLimLow, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.LimitUpper mLimUpp, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mLimUpp, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.MathFunction mFunc, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mFunc, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Matrix mMat, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mMat, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Nary mNary, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mNary, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Phantom mPhant, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mPhant, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.PreSubSuper mPreSubSup, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mPreSubSup, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Radical mRad, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mRad, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Subscript mSub, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mSub, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.SubSuperscript mSubSup, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mSubSup, d);
    }

    public virtual void VisitOMathElement(DocumentFormat.OpenXml.Math.Superscript mSup, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathElement(mSup, d);
    }

    public virtual void VisitOMathParagraph(DocumentFormat.OpenXml.Math.Paragraph oMathPara, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathParagraph(oMathPara, d);
    }

    public virtual void VisitOMathRun(DocumentFormat.OpenXml.Math.Run mMathRun, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitOMathRun(mMathRun, d);
    }

    public virtual void VisitPageNumber(PageNumber pn, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitPageNumber(pn, d);
    }

    public virtual IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitParagraphBegin(p, d, paragraph) ?? DxpDisposable.Empty;
	}

    public virtual void VisitPermEnd(PermEnd pe2, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitPermEnd(pe2, d);
    }

    public virtual void VisitPermStart(PermStart ps, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitPermStart(ps, d);
    }

    public virtual void VisitPositionalTab(PositionalTab ptab, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitPositionalTab(ptab, d);
    }

    public virtual void VisitProofError(ProofError pe, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitProofError(pe, d);
    }

    public virtual IDisposable VisitRubyBegin(Ruby ruby, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitRubyBegin(ruby, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitRubyContentBegin(RubyContentType rc, bool isBase, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitRubyContentBegin(rc, isBase, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitRunBegin(r, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitSectionBegin(SectionProperties properties, SectionLayout layout, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitSectionBegin(properties, layout, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitSectionBodyBegin(SectionProperties properties, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitSectionBodyBegin(properties, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitSectionFooterBegin(Footer ftr, object value, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitSectionFooterBegin(ftr, value, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitSectionHeaderBegin(Header hdr, object value, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitSectionHeaderBegin(hdr, value, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitSeparatorMark(SeparatorMark sep, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitSeparatorMark(sep, d);
    }

    public virtual IDisposable VisitSdtBlockBegin(SdtBlock sdt, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitSdtBlockBegin(sdt, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitSdtCellBegin(SdtCell sdtCell, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitSdtCellBegin(sdtCell, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitSdtContentBlockBegin(SdtContentBlock content, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitSdtContentBlockBegin(content, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitSdtContentRunBegin(SdtContentRun content, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitSdtContentRunBegin(content, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitSdtRowBegin(SdtRow sdtRow, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitSdtRowBegin(sdtRow, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitSdtRunBegin(SdtRun sdtRun, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitSdtRunBegin(sdtRun, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitSimpleFieldBegin(fld, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitSoftHyphen(SoftHyphen sh, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitSoftHyphen(sh, d);
    }

    public virtual IDisposable VisitSmartTagRunBegin(OpenXmlUnknownElement smart, string elementName, string elementUri, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitSmartTagRunBegin(smart, elementName, elementUri, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitSubDocumentReference(SubDocumentReference subDoc, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitSubDocumentReference(subDoc, d);
    }

    public virtual void VisitSymbol(SymbolChar sym, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitSymbol(sym, d);
    }

    public virtual void VisitTab(TabChar tab, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitTab(tab, d);
    }

    public virtual IDisposable VisitTableBegin(Table t, DxpTableModel model, DxpIDocumentContext d, DxpITableContext table)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitTableBegin(t, model, d, table) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitTableCellBegin(TableCell tc, DxpITableCellContext cell, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitTableCellBegin(tc, cell, d) ?? DxpDisposable.Empty;
	}

    public virtual IDisposable VisitTableRowBegin(TableRow tr, DxpITableRowContext row, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitTableRowBegin(tr, row, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitText(Text t, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitText(t, d);
    }

    public virtual IDisposable VisitTextBoxContentBegin(TextBoxContent txbx, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return DxpDisposable.Empty;
        return Next?.VisitTextBoxContentBegin(txbx, d) ?? DxpDisposable.Empty;
	}

    public virtual void VisitUnknown(string context, OpenXmlElement el, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitUnknown(context, el, d);
    }

    public virtual void VisitYearLong(YearLong yl, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitYearLong(yl, d);
    }

    public virtual void VisitYearShort(YearShort ys, DxpIDocumentContext d)
    {
        if (!ShouldForwardContent(d))
            return;
        Next?.VisitYearShort(ys, d);
    }

}
