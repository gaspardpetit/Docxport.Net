using DocxportNet.API;

namespace DocxportNet.Middleware;

internal interface IDxpEmbeddedWalkCompletion
{
    bool HasPendingEmbeddedWork(DxpIDocumentContext documentContext);
    void CompleteEmbeddedWalk(DxpIDocumentContext documentContext);
}
