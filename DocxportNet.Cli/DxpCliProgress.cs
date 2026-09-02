using DocxportNet;

internal sealed class DxpCliProgress : IProgress<DxpExportProgress>
{
    private const int BarWidth = 30;
    private readonly object _gate = new();
    private int _lastLineLength;

    public void Report(DxpExportProgress value)
    {
        lock (_gate)
        {
            string phase = value.Phase switch {
                DxpExportPhase.Opening => "Opening",
                DxpExportPhase.Preparing => "Preparing",
                DxpExportPhase.Converting => "Converting",
                DxpExportPhase.Finalizing => "Finalizing",
                DxpExportPhase.Completed => "Completed",
                _ => value.Phase.ToString()
            };

            int percent = value.Percentage.HasValue
                ? (int)Math.Round(value.Percentage.Value)
                : 0;
            int filled = value.Percentage.HasValue
                ? Math.Clamp(percent * BarWidth / 100, 0, BarWidth)
                : 0;
            string percentage = value.Percentage.HasValue ? $"{percent,3}%" : " --%";
            string units = value.TotalUnits > 0
                ? $" {value.CompletedUnits:N0}/{value.TotalUnits:N0} paragraphs"
                : string.Empty;
            string line = $"[{new string('#', filled)}{new string('-', BarWidth - filled)}] {percentage} {phase}{units}";

            Console.Error.Write('\r');
            Console.Error.Write(line);
            if (line.Length < _lastLineLength)
                Console.Error.Write(new string(' ', _lastLineLength - line.Length));
            _lastLineLength = line.Length;

            if (value.Phase == DxpExportPhase.Completed)
            {
                Console.Error.WriteLine();
                _lastLineLength = 0;
            }
        }
    }
}
