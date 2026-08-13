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
    Exception? Error = null)
{
    public static DxpSemanticFieldResult Empty(DxpFieldEvalStatus status = DxpFieldEvalStatus.Resolved)
        => new(status, DxpSemanticContent.Empty);
}

public sealed record DxpSemanticContent(IReadOnlyList<DxpSemanticNode> Nodes)
{
    public static DxpSemanticContent Empty { get; } = new(Array.Empty<DxpSemanticNode>());
    public bool IsEmpty => Nodes.Count == 0;
    public bool HasBlocks => Nodes.Any(static node =>
        node is DxpSemanticParagraph or DxpSemanticTable or DxpSemanticInclude);
}

public abstract record DxpSemanticNode;

public sealed record DxpSemanticText(string Text) : DxpSemanticNode;
public sealed record DxpSemanticBreak : DxpSemanticNode;
public sealed record DxpSemanticTab : DxpSemanticNode;
public sealed record DxpSemanticParagraph(DxpSemanticContent Content) : DxpSemanticNode;
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
