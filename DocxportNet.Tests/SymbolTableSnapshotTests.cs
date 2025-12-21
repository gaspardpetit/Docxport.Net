using Xunit.Abstractions;
using System.Globalization;
using System.Text;

namespace DocxportNet.Tests;

/// <summary>
/// Utility test to generate 256-entry lookup initializers for each symbol font.
/// Run this test to emit the C# literals you can paste into the mapping tables.
/// </summary>
public class SymbolTableSnapshotTests
{
	private readonly ITestOutputHelper _output;

	public SymbolTableSnapshotTests(ITestOutputHelper output) => _output = output;

	[Fact]
	public void Generate_symbol_font_arrays()
	{
		var destDir = Path.Combine("artifacts", "symbol-tables");
		Directory.CreateDirectory(destDir);

		WriteSnapshot("Symbol", destDir);
		WriteSnapshot("Zapf Dingbats", destDir);
		WriteSnapshot("Webdings", destDir);
		WriteSnapshot("Wingdings", destDir);
		WriteSnapshot("Wingdings 2", destDir);
		WriteSnapshot("Wingdings 3", destDir);
	}

	private void WriteSnapshot(string font, string destDir)
	{
		string safeName = font.Replace(' ', '_');
		string path = Path.Combine(destDir, $"{safeName}.txt");
		var snapshot = BuildArrayLiteral(font);
		File.WriteAllText(path, snapshot, Encoding.UTF8);
		_output.WriteLine($"{font}: wrote {path}");
	}

	private static string BuildArrayLiteral(string font)
	{
		var values = new string?[256];
		for (int i = 0; i < 256; i++)
		{
			string mapped = DxpFontSymbols.Substitute(font, (char)i);
			values[i] = string.IsNullOrEmpty(mapped) ? null : mapped;
		}

		var sb = new StringBuilder();
		sb.AppendLine($"// {font} generated {DateTime.UtcNow:O}");
		sb.AppendLine("new string?[] {");
		for (int i = 0; i < values.Length; i++)
		{
			string literal = values[i] is null ? "null" : ToPrintableLiteral(values[i]!);
			string comma = i == values.Length - 1 ? string.Empty : ",";
			sb.Append("    ")
			  .Append(literal)
			  .Append(comma)
			  .Append(" // 0x")
			  .Append(i.ToString("X2", CultureInfo.InvariantCulture))
			  .AppendLine();
		}
		sb.AppendLine("};");
		return sb.ToString();
	}

	private static string ToPrintableLiteral(string value)
	{
		if (!IsPrintable(value))
			return ToCodepointLiteral(value);

		var sb = new StringBuilder(value.Length + 2);
		sb.Append('"');
		foreach (var ch in value)
		{
			sb.Append(ch switch
			{
				'\\' => "\\\\",
				'"' => "\\\"",
				'\n' => "\\n",
				'\r' => "\\r",
				'\t' => "\\t",
				_ => ch.ToString()
			});
		}
		sb.Append('"');
		return sb.ToString();
	}

	private static string ToCodepointLiteral(string value)
	{
		var sb = new StringBuilder();
		sb.Append('"');
		int i = 0;
		while (i < value.Length)
		{
			if (!Rune.TryGetRuneAt(value, i, out var rune))
				break;
			if (rune.Value <= 0xFFFF)
				sb.Append("\\u").Append(rune.Value.ToString("X4", CultureInfo.InvariantCulture));
			else
				sb.Append("\\U").Append(rune.Value.ToString("X8", CultureInfo.InvariantCulture));
			i += rune.Utf16SequenceLength;
		}
		sb.Append('"');
		return sb.ToString();
	}

	private static bool IsPrintable(string value)
	{
		int i = 0;
		while (i < value.Length)
		{
			if (!Rune.TryGetRuneAt(value, i, out var rune))
				return false;

			var cat = Rune.GetUnicodeCategory(rune);
			if (cat == UnicodeCategory.Control ||
			    cat == UnicodeCategory.Format ||
			    cat == UnicodeCategory.OtherNotAssigned ||
			    cat == UnicodeCategory.Surrogate ||
			    cat == UnicodeCategory.PrivateUse)
				return false;

			i += rune.Utf16SequenceLength;
		}
		return true;
	}
}
