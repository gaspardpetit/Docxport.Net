namespace DocxportNet.api;

public interface IDxpStyleTracker
{
	void ResetStyle(IDxpVisitor v);
	void ApplyStyle(DxpStyleEffectiveRunStyle style, IDxpVisitor v);
}
