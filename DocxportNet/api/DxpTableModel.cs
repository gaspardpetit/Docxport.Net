namespace DocxportNet.API;

public sealed record DxpTableModel(
    int ColumnCount,
    int RowCount,
    IReadOnlyList<int?> GridColTwips, // from tblGrid if present
    DxpCellModel?[,] Cells               // rectangular matrix
);

