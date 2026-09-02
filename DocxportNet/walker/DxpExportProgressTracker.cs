using System.Diagnostics;

namespace DocxportNet.Walker;

internal sealed class DxpExportProgressTracker
{
    private readonly IProgress<DxpExportProgress> _progress;
    private readonly Stopwatch _reportTimer = Stopwatch.StartNew();
    private long _completed;
    private long _total;
    private int _lastReportedPercent = -1;

    public DxpExportProgressTracker(IProgress<DxpExportProgress> progress)
    {
        _progress = progress;
    }

    public void ReportPhase(DxpExportPhase phase)
    {
        _progress.Report(new DxpExportProgress(phase, _completed, _total));
        _reportTimer.Restart();
    }

    public void BeginConversion(long total)
    {
        _total = Math.Max(0, total);
        _completed = 0;
        _lastReportedPercent = -1;
        ReportPhase(DxpExportPhase.Converting);
    }

    public void ParagraphCompleted()
    {
        if (_completed < _total)
            _completed++;

        int percent = _total == 0 ? 0 : (int)(100 * _completed / _total);
        bool first = _completed == 1;
        if (first || (_reportTimer.ElapsedMilliseconds >= 100 && percent > _lastReportedPercent))
        {
            _lastReportedPercent = percent;
            _progress.Report(new DxpExportProgress(DxpExportPhase.Converting, _completed, _total));
            _reportTimer.Restart();
        }
    }

    public void Finalizing()
    {
        _completed = _total;
        ReportPhase(DxpExportPhase.Finalizing);
    }

    public void Completed()
    {
        _completed = _total;
        ReportPhase(DxpExportPhase.Completed);
    }
}
