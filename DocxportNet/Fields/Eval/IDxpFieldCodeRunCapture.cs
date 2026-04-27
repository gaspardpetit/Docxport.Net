using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Fields.Eval;

internal interface IDxpFieldCodeRunCapture
{
    void TryCaptureCodeRun(Run r);
}
