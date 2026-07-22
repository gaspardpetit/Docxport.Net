namespace DocxportNet.Fields.Eval;

internal sealed class DxpAutoNumberContext
{
    private readonly Dictionary<string, DxpAutoNumberStoryState> _stories = new(StringComparer.Ordinal);

    internal DxpAutoNumberStoryState GetStory(string storyKey)
    {
        if (!_stories.TryGetValue(storyKey, out var state))
        {
            state = new DxpAutoNumberStoryState();
            _stories[storyKey] = state;
        }
        return state;
    }

    internal void Reset() => _stories.Clear();
}

internal sealed class DxpAutoNumberStoryState
{
    internal DxpAutoNumberFamilyState AutoNum { get; } = new();
    internal DxpAutoNumberFamilyState Legal { get; } = new();
    internal DxpAutoNumberFamilyState Outline { get; } = new();
}

internal sealed class DxpAutoNumberFamilyState
{
    internal int[] HeadingCounters { get; } = new int[9];
    internal int BodyCounter { get; set; }
    internal int HeadingGeneration { get; set; }
    internal int BodyHeadingGeneration { get; set; } = -1;

    internal int AdvanceHeading(int oneBasedLevel)
    {
        int level = Math.Max(1, Math.Min(9, oneBasedLevel)) - 1;
        HeadingCounters[level]++;
        for (int i = level + 1; i < HeadingCounters.Length; i++)
            HeadingCounters[i] = 0;
        HeadingGeneration++;
        BodyCounter = 0;
        BodyHeadingGeneration = HeadingGeneration;
        return HeadingCounters[level];
    }

    internal int AdvanceBody()
    {
        if (BodyHeadingGeneration != HeadingGeneration)
        {
            BodyCounter = 0;
            BodyHeadingGeneration = HeadingGeneration;
        }
        return ++BodyCounter;
    }

    internal IReadOnlyList<int> CurrentPath(int throughLevel)
    {
        int count = Math.Max(0, Math.Min(9, throughLevel));
        return HeadingCounters.Take(count).Select(value => value == 0 ? 1 : value).ToArray();
    }
}
