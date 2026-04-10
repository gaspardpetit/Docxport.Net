using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet;

public sealed class DxpCollapseEquivalentRunsTransformer : IDxpNodeTransformer
{
    public DxpTransformDecision Visit(OpenXmlElement node, DxpTransformContext context)
    {
        if (node is not Paragraph paragraph)
            return DxpTransformDecision.Keep();

        var rebuilt = DxpParagraphTransformHelpers.RebuildParagraph(
            paragraph,
            static segments => CollapseEquivalentRuns(segments));
        return DxpTransformDecision.Replace(rebuilt);
    }

    private static IReadOnlyList<Run> CollapseEquivalentRuns(IReadOnlyList<Run> sourceRuns)
    {
        var merged = new List<Run>();
        foreach (var run in sourceRuns)
        {
            if (merged.Count == 0)
            {
                merged.Add(run);
                continue;
            }

            var previous = merged[merged.Count - 1];
            if (!DxpParagraphTransformHelpers.HasEquivalentRunProperties(previous, run))
            {
                merged.Add(run);
                continue;
            }

            foreach (var child in DxpParagraphTransformHelpers.CloneRunPayload(run))
                previous.Append(child);
        }

        return merged;
    }
}

public sealed class DxpSimplifyParagraphRunsTransformer : IDxpNodeTransformer
{
    public DxpTransformDecision Visit(OpenXmlElement node, DxpTransformContext context)
    {
        if (node is not Paragraph paragraph)
            return DxpTransformDecision.Keep();

        var rebuilt = DxpParagraphTransformHelpers.RebuildParagraph(
            paragraph,
            static segments => SimplifyToSingleRun(segments));
        return DxpTransformDecision.Replace(rebuilt);
    }

    private static IReadOnlyList<Run> SimplifyToSingleRun(IReadOnlyList<Run> sourceRuns)
    {
        if (sourceRuns.Count == 0)
            return Array.Empty<Run>();

        var simplified = new Run();
        foreach (var run in sourceRuns)
        {
            foreach (var child in DxpParagraphTransformHelpers.CloneRunPayload(run))
                simplified.Append(child);
        }

        return new[] { simplified };
    }
}

internal static class DxpParagraphTransformHelpers
{
    public static Paragraph RebuildParagraph(Paragraph paragraph, Func<IReadOnlyList<Run>, IReadOnlyList<Run>> rewriteRuns)
    {
        if (paragraph == null)
            throw new ArgumentNullException(nameof(paragraph));
        if (rewriteRuns == null)
            throw new ArgumentNullException(nameof(rewriteRuns));

        var rebuilt = (Paragraph)paragraph.CloneNode(false);
        if (paragraph.ParagraphProperties != null)
            rebuilt.ParagraphProperties = (ParagraphProperties)paragraph.ParagraphProperties.CloneNode(true);

        var extractedRuns = ExtractTextCompatibleRuns(paragraph);
        foreach (var run in rewriteRuns(extractedRuns))
            rebuilt.Append(run);

        return rebuilt;
    }

    public static bool HasEquivalentRunProperties(Run left, Run right)
    {
        if (left.RunProperties == null || right.RunProperties == null)
            return left.RunProperties == null && right.RunProperties == null;

        return string.Equals(left.RunProperties.OuterXml, right.RunProperties.OuterXml, StringComparison.Ordinal);
    }

    public static IEnumerable<OpenXmlElement> CloneRunPayload(Run run)
    {
        foreach (var child in run.ChildElements)
        {
            if (child is RunProperties)
                continue;

            yield return (OpenXmlElement)child.CloneNode(true);
        }
    }

    private static List<Run> ExtractTextCompatibleRuns(Paragraph paragraph)
    {
        var runs = new List<Run>();
        foreach (var sourceRun in paragraph.Descendants<Run>())
        {
            var rebuiltRun = new Run();
            if (sourceRun.RunProperties != null)
                rebuiltRun.RunProperties = (RunProperties)sourceRun.RunProperties.CloneNode(true);

            foreach (var child in sourceRun.ChildElements)
            {
                if (!IsTextCompatibleLeaf(child))
                    continue;

                rebuiltRun.Append(child.CloneNode(true));
            }

            if (rebuiltRun.ChildElements.Count > 0)
                runs.Add(rebuiltRun);
        }

        return runs;
    }

    private static bool IsTextCompatibleLeaf(OpenXmlElement child)
        => child is Text
            || child is TabChar
            || child is Break
            || child is CarriageReturn
            || child is NoBreakHyphen;
}
