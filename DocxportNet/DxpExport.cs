using DocumentFormat.OpenXml.Packaging;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet;

/// <summary>
/// High-level helpers for running a <see cref="DxpIVisitor"/> against a DOCX source and collecting the result.
/// These overloads cover common entry points (file path, in-memory bytes, already-open <see cref="WordprocessingDocument"/>)
/// and common sinks (existing <see cref="TextWriter"/>, returning a <see cref="string"/>, or writing to a file).
/// </summary>
public static class DxpExport
{
	/// <summary>
	/// Export to a text string using a <see cref="DxpITextVisitor"/> and a DOCX file path.
	/// </summary>
	public static string ExportToString(string docxPath, DxpITextVisitor visitor, ILogger? logger = null)
	{
		using var writer = new StringWriter();
		visitor.SetOutput(writer);
		try
		{
			RunWalker(docxPath, visitor, logger);
			return writer.ToString();
		}
		finally
		{
			DisposeVisitor(visitor);
		}
	}

	/// <summary>
	/// Drive a visitor without caring about output (e.g., collectors). A null sink is assigned.
	/// </summary>
	public static void Export(string docxPath, DxpIVisitor visitor, ILogger? logger = null)
	{
		visitor.SetOutput(Stream.Null);
		try
		{
			RunWalker(docxPath, visitor, logger);
		}
		finally
		{
			DisposeVisitor(visitor);
		}
	}

	/// <summary>
	/// Export to a text string using a <see cref="DxpITextVisitor"/> and an already-open <see cref="WordprocessingDocument"/>.
	/// The document remains open; disposal is left to the caller.
	/// </summary>
	public static string ExportToString(WordprocessingDocument document, DxpITextVisitor visitor, ILogger? logger = null)
	{
		using var writer = new StringWriter();
		visitor.SetOutput(writer);
		try
		{
			RunWalker(document, visitor, logger);
			return writer.ToString();
		}
		finally
		{
			DisposeVisitor(visitor);
		}
	}

	/// <summary>
	/// Export to a text string using a <see cref="DxpITextVisitor"/> and in-memory DOCX bytes.
	/// </summary>
	public static string ExportToString(byte[] docxBytes, DxpITextVisitor visitor, ILogger? logger = null)
	{
		using var stream = new MemoryStream(docxBytes, writable: false);
		using var document = WordprocessingDocument.Open(stream, false);
		return ExportToString(document, visitor, logger);
	}

	/// <summary>
	/// Export to a byte array using a <see cref="DxpIVisitor"/> and a DOCX file path.
	/// </summary>
	public static byte[] ExportToBytes(string docxPath, DxpIVisitor visitor, ILogger? logger = null)
	{
		using var ms = new MemoryStream();
		visitor.SetOutput(ms);
		try
		{
			RunWalker(docxPath, visitor, logger);
			return ms.ToArray();
		}
		finally
		{
			DisposeVisitor(visitor);
		}
	}

	/// <summary>
	/// Export to a byte array using a <see cref="DxpIVisitor"/> and an already-open <see cref="WordprocessingDocument"/>.
	/// </summary>
	public static byte[] ExportToBytes(WordprocessingDocument document, DxpIVisitor visitor, ILogger? logger = null)
	{
		using var ms = new MemoryStream();
		visitor.SetOutput(ms);
		try
		{
			RunWalker(document, visitor, logger);
			return ms.ToArray();
		}
		finally
		{
			DisposeVisitor(visitor);
		}
	}

	/// <summary>
	/// Export to a byte array using a <see cref="DxpIVisitor"/> and in-memory DOCX bytes.
	/// </summary>
	public static byte[] ExportToBytes(byte[] docxBytes, DxpIVisitor visitor, ILogger? logger = null)
	{
		using var stream = new MemoryStream(docxBytes, writable: false);
		using var document = WordprocessingDocument.Open(stream, false);
		return ExportToBytes(document, visitor, logger);
	}

	/// <summary>
	/// Walks the DOCX at <paramref name="docxPath"/> with <paramref name="visitor"/> and returns the collected text.
	/// </summary>
	public static string ExportToFile(string docxPath, DxpIVisitor visitor, string outputPath, ILogger? logger = null)
	{
		CreateParentDirectory(outputPath);

		using var fileStream = File.Create(outputPath);
		visitor.SetOutput(fileStream);
		try
		{
			RunWalker(docxPath, visitor, logger);
			fileStream.Flush();
			return outputPath;
		}
		finally
		{
			DisposeVisitor(visitor);
		}
	}

	/// <summary>
	/// Walks in-memory DOCX bytes with <paramref name="visitor"/> and writes the output to <paramref name="outputPath"/>.
	/// </summary>
	public static string ExportToFile(byte[] docxBytes, DxpIVisitor visitor, string outputPath, ILogger? logger = null)
	{
		using var stream = new MemoryStream(docxBytes, writable: false);
		using var document = WordprocessingDocument.Open(stream, false);
		return ExportToFile(document, visitor, outputPath, logger);
	}

	/// <summary>
	/// Walks an already-open <see cref="WordprocessingDocument"/> with <paramref name="visitor"/> and writes the output to <paramref name="outputPath"/>.
	/// The document remains open; disposal is left to the caller.
	/// </summary>
	public static string ExportToFile(WordprocessingDocument document, DxpIVisitor visitor, string outputPath, ILogger? logger = null)
	{
		CreateParentDirectory(outputPath);

		using var fileStream = File.Create(outputPath);
		visitor.SetOutput(fileStream);
		try
		{
			RunWalker(document, visitor, logger);
			fileStream.Flush();
			return outputPath;
		}
		finally
		{
			DisposeVisitor(visitor);
		}
	}

	private static void CreateParentDirectory(string path)
	{
		string? directory = Path.GetDirectoryName(path);
		if (!string.IsNullOrEmpty(directory))
			Directory.CreateDirectory(directory);
	}

	private static void RunWalker(string docxPath, DxpIVisitor visitor, ILogger? logger)
	{
		var walker = new DxpWalker(logger);
		walker.Accept(docxPath, WrapWithFieldEvalMiddleware(visitor));
	}

	private static void RunWalker(WordprocessingDocument document, DxpIVisitor visitor, ILogger? logger)
	{
		var walker = new DxpWalker(logger);
		walker.Accept(document, WrapWithFieldEvalMiddleware(visitor));
	}

	private static DxpIVisitor WrapWithFieldEvalMiddleware(DxpIVisitor visitor)
	{
		if (visitor is IDxpFieldEvalProvider provider)
		{
			return DxpVisitorMiddleware.Chain(
				visitor,
				next => new DxpFieldEvalVisitorMiddleware(next, provider.FieldEval.Context));
		}

		return visitor;
	}

	private static void DisposeVisitor(DxpIVisitor visitor)
	{
		if (visitor is IDisposable disposable)
			disposable.Dispose();
	}
}
