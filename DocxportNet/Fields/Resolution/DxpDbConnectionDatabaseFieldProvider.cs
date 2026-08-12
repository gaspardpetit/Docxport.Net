using System.Data.Common;
using System.Globalization;

namespace DocxportNet.Fields.Resolution;

/// <summary>
/// Executes DATABASE field queries through a caller-supplied ADO.NET connection.
/// Connection selection and credentials remain the caller's responsibility.
/// </summary>
public sealed class DxpDbConnectionDatabaseFieldProvider : IDatabaseFieldProvider
{
    private readonly Func<DxpDatabaseRequest, CancellationToken, Task<DbConnection>> _connectionFactory;

    public DxpDbConnectionDatabaseFieldProvider(
        Func<DxpDatabaseRequest, CancellationToken, Task<DbConnection>> connectionFactory)
        => _connectionFactory = connectionFactory ?? throw new ArgumentNullException(nameof(connectionFactory));

    public async Task<DxpDatabaseResult?> ExecuteAsync(
        DxpDatabaseRequest request,
        CancellationToken cancellationToken)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));

        using DbConnection connection = await _connectionFactory(request, cancellationToken).ConfigureAwait(false)
            ?? throw new InvalidOperationException("The DATABASE connection factory returned null.");
        if (connection.State != System.Data.ConnectionState.Open)
            await connection.OpenAsync(cancellationToken).ConfigureAwait(false);

        using DbCommand command = connection.CreateCommand();
        command.CommandText = request.QueryText;
        AddParameters(command, request.Parameters);

        using DbDataReader reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);
        var columns = new List<DxpDatabaseColumn>(reader.FieldCount);
        for (int i = 0; i < reader.FieldCount; i++)
            columns.Add(new DxpDatabaseColumn(reader.GetName(i), ToFieldKind(reader.GetFieldType(i))));

        var rows = new List<IReadOnlyList<DxpFieldValue?>>();
        while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
        {
            var row = new DxpFieldValue?[reader.FieldCount];
            for (int i = 0; i < reader.FieldCount; i++)
                row[i] = reader.IsDBNull(i) ? null : ToFieldValue(reader.GetValue(i));
            rows.Add(row);
        }

        return new DxpDatabaseResult(columns, rows);
    }

    private static void AddParameters(
        DbCommand command,
        IReadOnlyDictionary<string, DxpFieldValue>? values)
    {
        if (values == null)
            return;

        foreach (var pair in values)
        {
            DbParameter parameter = command.CreateParameter();
            parameter.ParameterName = pair.Key;
            parameter.Value = ToDbValue(pair.Value);
            command.Parameters.Add(parameter);
        }
    }

    private static object ToDbValue(DxpFieldValue value)
        => value.Kind switch {
            DxpFieldValueKind.Number => value.NumberValue.GetValueOrDefault(),
            DxpFieldValueKind.DateTime => value.DateTimeValue?.UtcDateTime ?? (object)DBNull.Value,
            _ => value.StringValue ?? (object)DBNull.Value
        };

    private static DxpFieldValue? ToFieldValue(object value)
    {
        if (value is DateTimeOffset offset)
            return new DxpFieldValue(offset);
        if (value is DateTime dateTime)
            return new DxpFieldValue(new DateTimeOffset(dateTime));
        if (IsNumber(value))
            return new DxpFieldValue(Convert.ToDouble(value, CultureInfo.InvariantCulture));
        return new DxpFieldValue(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
    }

    private static DxpFieldValueKind? ToFieldKind(Type type)
    {
        type = Nullable.GetUnderlyingType(type) ?? type;
        if (type == typeof(DateTime) || type == typeof(DateTimeOffset))
            return DxpFieldValueKind.DateTime;
        if (IsNumberType(type))
            return DxpFieldValueKind.Number;
        return DxpFieldValueKind.String;
    }

    private static bool IsNumber(object value) => IsNumberType(value.GetType());

    private static bool IsNumberType(Type type)
        => type == typeof(byte) || type == typeof(sbyte) ||
           type == typeof(short) || type == typeof(ushort) ||
           type == typeof(int) || type == typeof(uint) ||
           type == typeof(long) || type == typeof(ulong) ||
           type == typeof(float) || type == typeof(double) || type == typeof(decimal);
}
