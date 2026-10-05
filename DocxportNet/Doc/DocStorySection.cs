namespace DocxportNet.Doc;

/// <summary>A section's main-story bounds, formatting, and six header/footer slots.</summary>
public sealed record DocStorySection(int StartCp, int EndCp,
    DocSectionFormatting Formatting, IReadOnlyList<DocStorySectionSlot> Slots,
    long? SourceTextOffsetStart = null, long? SourceTextOffsetEnd = null,
    long? SourceSepxOffset = null)
{
    public void Validate()
    {
        if (StartCp < 0 || EndCp <= StartCp || Slots.Count != 6 ||
            Slots.Select(x => x.Slot).Where((slot, index) => slot != index).Any())
            throw new InvalidDataException("A section has invalid bounds or header/footer slots.");
    }
}

/// <summary>A section slot; source CPs address visible content without its DOC guard mark.</summary>
public sealed record DocStorySectionSlot(int Slot, bool IsPresent,
    uint? SourceCpStart = null, uint? SourceCpEnd = null);
