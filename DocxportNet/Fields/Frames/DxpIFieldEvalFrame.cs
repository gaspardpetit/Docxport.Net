using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;

namespace DocxportNet.Fields.Frames;

internal interface DxpIFieldEvalFrame
{
    bool SuppressContent { get; set; }
    bool Evaluated { get; set; }
    bool InResult { get; set; }
    string? InstructionText { get; set; }
    RunProperties? CodeRunProperties { get; set; }
    Run? CodeRun { get; set; }
    List<Run?>? CachedResultRuns { get; set; }
    DxpIFCaptureState? IfState { get; set; }
}
