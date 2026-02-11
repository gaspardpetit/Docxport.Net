using DocxportNet.Fields.Frames;

namespace DocxportNet.Fields;

public sealed partial class DxpFieldEvalContext
{
    internal DxpIFieldEvalFrame? OuterFrame { get; set; }
    internal int FieldDepth { get; set; }
    internal DocumentFormat.OpenXml.Wordprocessing.Run? CurrentRun { get; set; }
}
