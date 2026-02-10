using System.Globalization;

namespace DocxportNet.Formatting;

public interface DxpINumberToWordsProvider
{
    bool CanHandle(CultureInfo culture);
    string ToCardinal(int number);
    string ToOrdinalWords(int number);
    string ToDollarText(double number);
}
