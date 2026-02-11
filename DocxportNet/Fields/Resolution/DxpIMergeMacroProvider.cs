using System.Globalization;

namespace DocxportNet.Fields.Resolution;

public interface DxpIMergeMacroProvider
{
    bool CanHandle(CultureInfo culture);
    string? Resolve(string macroName, IDxpMergeRecordCursor cursor, CultureInfo culture);
}
