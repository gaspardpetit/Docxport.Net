using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Formatting;
using DocxportNet.Walker;
using DocxportNet.Walker.Context;
using Microsoft.Extensions.Logging;

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

		var t = new Text(text);
		if (DxpFieldEvalMiddleware.NeedsPreserveSpace(text))
			t.Space = SpaceProcessingModeValues.Preserve;
		run.AppendChild(t);

        EmitRun(run, d, sink);
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

	public static void EmitTextMergeformatWithRuns(
        string text,
        DxpIDocumentContext d,
        IReadOnlyList<Run?>? runs,
		DxpIVisitor sink)
	{
		if (string.IsNullOrEmpty(text))
			return;

		if (runs == null || runs.Count == 0)
		{
            EmitTextInRun(text, d, NewSyntheticRun(null, null), sink);
			return;
		}

		int segmentCount = runs.Count;

		var segments = SplitTextByRuns(text, segmentCount);
		for (int i = 0; i < segments.Count; i++)
		{
			Run? segmentRun = runs != null && i < runs.Count ? runs[i] : null;
            EmitTextInRun(segments[i], d, NewSyntheticRun(segmentRun, null), sink);
		}
	}

    internal static bool EmitTextWithMergeFormat(
		string resultText,
		IReadOnlyList<IDxpFieldFormatSpec> formatSpecs,
		List<Run?>? cachedResultRuns,
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
			IReadOnlyList<Run?>? mergeRuns = null;

			if (hasMergeFormat && cachedResultRuns != null && cachedResultRuns.Count > 0)
			{
				mergeRuns = cachedResultRuns;
			}
			else if (hasCharFormat && codeRun?.RunProperties != null)
			{
				runProps = codeRun.RunProperties;
			}

			if (hasMergeFormat && cachedResultRuns != null && cachedResultRuns.Count > 0)
				mergeRuns = cachedResultRuns;
			else if (hasCharFormat && codeRun?.RunProperties != null)
				runProps = codeRun.RunProperties;

			if (hasCharFormat && runProps == null && logger?.IsEnabled(LogLevel.Debug) == true)
				logger.LogDebug("CHARFORMAT requested but no field code run properties captured.");

			if (mergeRuns != null)
                EmitTextMergeformatWithRuns(resultText, d, mergeRuns, sink);
			else
                EmitTextInRun(resultText, d, NewSyntheticRun(codeRun, runProps ?? codeRun?.RunProperties), sink);
			return true;
		}
		else
		{
            EmitTextInRun(resultText, d, NewSyntheticRun(codeRun, codeRun?.RunProperties), sink);
			return true;
		}
	}
}
