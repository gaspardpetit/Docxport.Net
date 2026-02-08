namespace DocxportNet.Fields.Resolution;

public enum DxpTableRangeDirection
{
    Above,
    Below,
    Left,
    Right
}

public interface IDxpTableValueResolver
{
    Task<IReadOnlyList<double>> ResolveRangeAsync(string range, DxpFieldEvalContext context);
    Task<IReadOnlyList<double>> ResolveDirectionalRangeAsync(DxpTableRangeDirection direction, DxpFieldEvalContext context);
}
