namespace DocxportNet.Fields.Semantic;

/// <summary>
/// A format-neutral result produced by field evaluation. Export-specific
/// adapters decide how these nodes are represented in DOCX, HTML, Markdown,
/// or plain text.
/// </summary>
public sealed record DxpSemanticFieldResult(
    DxpFieldEvalStatus Status,
    DxpSemanticContent Content,
    DxpFieldValue? Value = null,
    Exception? Error = null,
    DxpSemanticSourceProvenance? Source = null)
{
    public static DxpSemanticFieldResult Empty(DxpFieldEvalStatus status = DxpFieldEvalStatus.Resolved)
        => new(status, DxpSemanticContent.Empty);
}

public sealed record DxpSemanticSourceProvenance(
    string? StoryKey,
    int? DocumentOrder,
    string? ParagraphId);

public sealed record DxpSemanticRunFormat(
    bool? Bold = null,
    bool? Italic = null,
    bool? Strike = null,
    string? Underline = null,
    string? Color = null,
    string? FontSizeHalfPoints = null,
    string? StyleId = null,
    string? Language = null);

public sealed record DxpSemanticParagraphFormat(
    string? StyleId = null,
    string? Alignment = null,
    int? OutlineLevel = null,
    int? NumberingId = null,
    int? NumberingLevel = null);

public sealed record DxpSemanticContent(IReadOnlyList<DxpSemanticNode> Nodes)
{
    public static DxpSemanticContent Empty { get; } = new(Array.Empty<DxpSemanticNode>());
    public bool IsEmpty => Nodes.Count == 0;
    public bool HasBlocks => Nodes.Any(static node =>
        node is DxpSemanticParagraph or DxpSemanticTable or DxpSemanticInclude);
}

public abstract record DxpSemanticNode;

public sealed record DxpSemanticText(string Text, DxpSemanticRunFormat? Format = null) : DxpSemanticNode;
public sealed record DxpSemanticBreak(DxpSemanticRunFormat? Format = null) : DxpSemanticNode;
public sealed record DxpSemanticTab(DxpSemanticRunFormat? Format = null) : DxpSemanticNode;
public sealed record DxpSemanticParagraph(
    DxpSemanticContent Content,
    DxpSemanticParagraphFormat? Format = null) : DxpSemanticNode;
public sealed record DxpSemanticInclude(
    string Path,
    string Identity,
    byte[] Content,
    string? Bookmark = null) : DxpSemanticNode;

public sealed record DxpSemanticTable(
    IReadOnlyList<DxpSemanticTableRow> Rows,
    bool HasHeader = false,
    bool ShowBorders = false,
    bool AutoFit = false) : DxpSemanticNode;

public sealed record DxpSemanticTableRow(IReadOnlyList<DxpSemanticTableCell> Cells);
public sealed record DxpSemanticTableCell(DxpSemanticContent Content);
