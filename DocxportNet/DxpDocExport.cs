using DocxportNet.Visitors.Doc;
using DocxportNet.Fields;
using Microsoft.Extensions.Logging;

namespace DocxportNet;

/// <summary>Writes an unformatted binary DOC from DOCX or binary DOC input.</summary>
public static class DxpDocExport
{
    public static byte[] Export(byte[] input, DxpExportOptions? options = null, ILogger? logger = null,
        DxpFieldEval? fieldEval = null)
        => DxpExport.ExportToBytes(input, new DxpDocVisitor(logger, fieldEval), options, logger);

    public static string Export(string inputPath, string outputPath, DxpExportOptions? options = null,
        ILogger? logger = null, DxpFieldEval? fieldEval = null)
        => DxpExport.ExportToFile(inputPath, new DxpDocVisitor(logger, fieldEval), outputPath, options, logger);
}
