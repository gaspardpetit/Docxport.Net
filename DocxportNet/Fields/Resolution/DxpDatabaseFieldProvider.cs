using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using DocxportNet.Fields;

namespace DocxportNet.Fields.Resolution;

public sealed record DxpDatabaseRequest(
    string QueryText,
    IReadOnlyDictionary<string, DxpFieldValue>? Parameters = null,
    IReadOnlyDictionary<string, string?>? Options = null,
    CultureInfo? Culture = null);

public sealed record DxpDatabaseColumn(string Name, DxpFieldValueKind? Kind = null);

public sealed record DxpDatabaseResult(
    IReadOnlyList<DxpDatabaseColumn> Columns,
    IReadOnlyList<IReadOnlyList<DxpFieldValue?>> Rows);

public interface IDatabaseFieldProvider
{
    Task<DxpDatabaseResult?> ExecuteAsync(
        DxpDatabaseRequest request,
        CancellationToken cancellationToken);
}
