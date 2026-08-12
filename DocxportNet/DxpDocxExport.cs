using DocxportNet.Visitors.Docx;
using DocxportNet.Fields;
using Microsoft.Extensions.Logging;

namespace DocxportNet;

/// <summary>Convenience API for rebuilding a DOCX through the existing visitor pipeline.</summary>
public static class DxpDocxExport
{
    public static string Export(string inputPath, string outputPath, DxpExportOptions? options = null, ILogger? logger = null, DxpFieldEval? fieldEval = null)
    {
        if (string.IsNullOrWhiteSpace(inputPath))
            throw new ArgumentException("Input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath))
            throw new ArgumentException("Output path is required.", nameof(outputPath));

        return DxpExport.ExportToFile(inputPath, new DxpDocxVisitor(logger, fieldEval), outputPath, options, logger);
    }

    public static byte[] Export(byte[] docxBytes, DxpExportOptions? options = null, ILogger? logger = null, DxpFieldEval? fieldEval = null)
        => DxpExport.ExportToBytes(docxBytes, new DxpDocxVisitor(logger, fieldEval), options, logger);
}
