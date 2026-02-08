namespace DocxportNet.API;

public interface DxpIStyleTracker
{
    void ResetStyle(DxpIDocumentContext d, DxpIVisitor v);
    void ApplyStyle(DxpStyleEffectiveRunStyle style, DxpIDocumentContext d, DxpIVisitor v);
}
