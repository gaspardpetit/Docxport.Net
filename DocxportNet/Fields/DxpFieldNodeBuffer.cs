using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Fields;

public sealed class DxpFieldNodeBuffer
{
    private readonly List<Action<DxpIVisitor, DxpIDocumentContext>> _actions = new();

    public static DxpFieldNodeBuffer FromText(string text)
    {
        var buffer = new DxpFieldNodeBuffer();
        buffer.AddRunText(text);
        return buffer;
    }

    public void Replay(DxpIVisitor visitor, DxpIDocumentContext context)
    {
        foreach (var action in _actions)
            action(visitor, context);
    }

    private void AddRunText(string text)
    {
        _actions.Add((visitor, context) => {
            var run = new Run();
            using (visitor.VisitRunBegin(run, context))
            {
                var t = new Text(text);
                if (NeedsPreserveSpace(text))
                    t.Space = SpaceProcessingModeValues.Preserve;
                visitor.VisitText(t, context);
            }
        });
    }

    private static bool NeedsPreserveSpace(string text)
    {
        if (text.Length == 0)
            return false;
        if (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]))
            return true;
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (ch == '\t' || ch == '\r' || ch == '\n')
                return true;
            if (ch == ' ' && i + 1 < text.Length && text[i + 1] == ' ')
                return true;
        }
        return false;
    }
}
