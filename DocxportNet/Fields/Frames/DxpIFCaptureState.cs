using System.Text;
using DocxportNet.Fields.Eval;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpIFCaptureState
{
    public int TokenIndex = 0;
    public bool FieldTypeConsumed;
    public bool InQuote;
    public int BraceDepth;
    public bool JustClosedQuote;
    public readonly StringBuilder CurrentToken = new();
    public readonly DxpFieldNodeBuffer TrueBuffer = new();
    public readonly DxpFieldNodeBuffer FalseBuffer = new();
    public readonly DxpEvalFieldNodeBufferRecorder Recorder = new();

    public DxpFieldNodeBuffer? GetCurrentBuffer()
    {
        return TokenIndex switch
        {
            3 => TrueBuffer,
            4 => FalseBuffer,
            _ => null
        };
    }
}
