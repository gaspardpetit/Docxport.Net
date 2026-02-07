using System.Globalization;

namespace DocxportNet.Formatting;

public interface IDxpNumberToWordsProvider
{
	bool CanHandle(CultureInfo culture);
	string ToCardinal(int number);
	string ToOrdinalWords(int number);
	string ToDollarText(double number);
}
