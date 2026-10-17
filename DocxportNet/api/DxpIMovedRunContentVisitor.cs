using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.API;

public interface DxpIMovedRunContentVisitor
{
    IDisposable? VisitMoveFromRunBegin(MoveFromRun moved, DxpIDocumentContext context);
    IDisposable? VisitMoveToRunBegin(MoveToRun moved, DxpIDocumentContext context);
}
