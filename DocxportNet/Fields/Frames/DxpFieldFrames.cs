using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Formatting;
using DocxportNet.Walker;
using DocxportNet.Walker.Context;
using Microsoft.Extensions.Logging;
using System.Linq;

namespace DocxportNet.Fields.Frames;

class DxpFieldFrames
{
	public static List<string> SplitTextByRuns(string text, int count)
	{
		var segments = new List<string>(count);
		if (count <= 1)
		{
			segments.Add(text);
			return segments;
		}

		int length = text.Length;
		int baseSize = length / count;
		int remainder = length % count;
		int offset = 0;

		for (int i = 0; i < count; i++)
		{
			int size = baseSize + (i < remainder ? 1 : 0);
			if (offset >= length)
			{
				segments.Add(string.Empty);
				continue;
			}
			if (offset + size > length)
				size = length - offset;
			segments.Add(text.Substring(offset, size));
			offset += size;
		}

		return segments;
	}

	public static void EmitRun(Run run, DxpIDocumentContext d, DxpIVisitor? sink)
	{
		if (sink == null)
			return;
		if (d is DxpDocumentContext docContext)
			d.Walker.WalkRun(run, docContext, sink);
	}

	public static void EmitTextInRun(string text, DxpIDocumentContext d, Run run, DxpIVisitor? sink)
	{
		if (sink == null)
			return;

        BuildTextInRunBuffer(text, run).Replay(sink, d);
	}

	public static Run NewSyntheticRun(Run? sourceRun, RunProperties? runProperties)
	{
		Run run = sourceRun != null
			? DxpRunCloner.CloneRunWithParagraphAncestor(sourceRun)
			: new Run();

		if (run.RunProperties == null && runProperties != null)
			run.RunProperties = (RunProperties)runProperties.CloneNode(true);

		return run;
	}

	public static DxpFieldNodeBuffer BuildTextInRunBuffer(string text, Run run)
	{
        var buffer = new DxpFieldNodeBuffer();
        var child = buffer.BeginRun(run);
        child.AddTextWithBreaks(text);
        return buffer;
	}

	public static DxpFieldNodeBuffer BuildTextMergeformatWithRuns(
        string text,
        IReadOnlyList<Run?>? runs)
	{
        var buffer = new DxpFieldNodeBuffer();

		if (string.IsNullOrEmpty(text))
			return buffer;

        text = text.Replace("\r\n", "\n").Replace('\r', '\n');

		if (runs == null || runs.Count == 0)
		{
            var child = buffer.BeginRun(NewSyntheticRun(null, null));
            child.AddText(text);
			return buffer;
		}

		int segmentCount = runs.Count;

		var segments = SplitTextByRuns(text, segmentCount);
		for (int i = 0; i < segments.Count; i++)
		{
			Run? segmentRun = runs != null && i < runs.Count ? runs[i] : null;
            var child = buffer.BeginRun(NewSyntheticRun(segmentRun, null));
            child.AddTextWithBreaks(segments[i]);
		}

        return buffer;
	}

    internal static bool EmitTextWithMergeFormat(
		string resultText,
		IReadOnlyList<IDxpFieldFormatSpec> formatSpecs,
		DxpFieldNodeBuffer? cachedResultBuffer,
		Run? codeRun,
		DxpIDocumentContext d,
		DxpIVisitor? sink,
		ILogger? logger)
    {
		if (sink == null)
			return true;

		bool hasMergeFormatting =
			DxpFieldEvalRules.TryGetCharOrMergeFormat(formatSpecs, out var hasCharFormat, out var hasMergeFormat) &&
			(hasCharFormat || hasMergeFormat);

		if (hasMergeFormatting)
		{
			RunProperties? runProps = null;
            List<(string text, RunProperties? props)>? segments = null;

			if (hasMergeFormat && cachedResultBuffer != null && cachedResultBuffer.TryGetRunSegments(out var cachedSegments))
				segments = cachedSegments;
			else if (hasCharFormat && codeRun?.RunProperties != null)
			{
				runProps = codeRun.RunProperties;
			}

			if (hasCharFormat && runProps == null && logger?.IsEnabled(LogLevel.Debug) == true)
				logger.LogDebug("CHARFORMAT requested but no field code run properties captured.");

			if (segments != null)
            {
                var runs = segments.Select(s => {
                    var run = NewSyntheticRun(null, s.props);
                    return run;
                }).Cast<Run?>().ToList();
                BuildTextMergeformatWithRuns(resultText, runs).Replay(sink, d);
            }
			else
                BuildTextInRunBuffer(resultText, NewSyntheticRun(codeRun, runProps ?? codeRun?.RunProperties)).Replay(sink, d);
			return true;
		}
		else
		{
            BuildTextInRunBuffer(resultText, NewSyntheticRun(codeRun, codeRun?.RunProperties)).Replay(sink, d);
			return true;
		}
	}
}
