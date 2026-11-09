using System.Text;

namespace DocxportNet.Walker;

// Word list spelling and rollover rules; shared field formatting remains separate.
internal static class DxpRegionalListWords
{
    private static readonly string[] HindiSmall =
    {
        "\u0936\u0942\u0928\u094d\u092f",
        "\u090f\u0915",
        "\u0926\u094b",
        "\u0924\u0940\u0928",
        "\u091a\u093e\u0930",
        "\u092a\u093e\u0901\u091a",
        "\u091b\u0903",
        "\u0938\u093e\u0924",
        "\u0906\u0920",
        "\u0928\u094c",
        "\u0926\u0938",
        "\u0917\u094d\u092f\u093e\u0930\u0939",
        "\u092c\u093e\u0930\u0939",
        "\u0924\u0947\u0930\u0939",
        "\u091a\u094c\u0926\u0939",
        "\u092a\u0902\u0926\u094d\u0930\u0939",
        "\u0938\u094b\u0932\u0939",
        "\u0938\u0924\u094d\u0930\u0939",
        "\u0905\u0920\u094d\u0920\u093e\u0930\u0939",
        "\u0909\u0928\u094d\u0928\u0940\u0938",
        "\u092c\u0940\u0938",
        "\u0907\u0915\u094d\u0915\u0940\u0938",
        "\u092c\u093e\u0908\u0938",
        "\u0924\u0947\u0908\u0938",
        "\u091a\u094c\u092c\u0940\u0938",
        "\u092a\u091a\u094d\u091a\u0940\u0938",
        "\u091b\u092c\u094d\u092c\u0940\u0938",
        "\u0938\u0924\u094d\u0924\u093e\u0908\u0938",
        "\u0905\u0920\u094d\u0920\u093e\u0908\u0938",
        "\u0909\u0928\u094d\u0924\u0940\u0938",
        "\u0924\u0940\u0938",
        "\u0907\u0915\u094d\u0924\u0940\u0938",
        "\u092c\u0924\u094d\u0924\u0940\u0938",
        "\u0924\u0948\u0902\u0924\u0940\u0938",
        "\u091a\u094c\u0902\u0924\u0940\u0938",
        "\u092a\u0948\u0902\u0924\u0940\u0938",
        "\u091b\u0924\u0940\u0938",
        "\u0938\u0948\u0902\u0924\u0940\u0938",
        "\u0905\u0921\u093c\u0924\u0940\u0938",
        "\u0909\u0928\u094d\u0924\u093e\u0932\u0940\u0938",
        "\u091a\u093e\u0932\u0940\u0938",
        "\u0907\u0915\u094d\u0924\u093e\u0932\u0940\u0938",
        "\u092c\u092f\u093e\u0932\u0940\u0938",
        "\u0924\u0948\u0924\u093e\u0932\u0940\u0938",
        "\u091a\u0935\u093e\u0932\u0940\u0938",
        "\u092a\u0948\u0902\u0924\u093e\u0932\u0940\u0938",
        "\u091b\u093f\u092f\u093e\u0932\u0940\u0938",
        "\u0938\u0924\u094d\u0924\u093e\u0932\u0940\u0938",
        "\u0905\u095c\u0924\u093e\u0932\u0940\u0938",
        "\u0909\u0928\u094d\u091a\u093e\u0938",
        "\u092a\u091a\u093e\u0938",
        "\u0907\u0915\u093e\u0935\u0928",
        "\u092c\u093e\u0935\u0928",
        "\u0924\u093f\u0930\u092a\u0928",
        "\u091a\u094c\u0935\u0928",
        "\u092a\u091a\u092a\u0928",
        "\u091b\u092a\u094d\u092a\u0928",
        "\u0938\u0924\u094d\u0924\u093e\u0935\u0928",
        "\u0905\u0920\u094d\u0920\u093e\u0935\u0928",
        "\u0909\u0928\u0938\u0920",
        "\u0938\u093e\u0920",
        "\u0907\u0915\u0938\u0920",
        "\u092c\u093e\u0938\u0920",
        "\u0924\u093f\u0930\u0938\u0920",
        "\u091a\u094c\u0902\u0938\u0920",
        "\u092a\u0948\u0902\u0938\u0920",
        "\u091b\u093f\u092f\u093e\u0938\u0920",
        "\u0938\u095c\u0938\u0920",
        "\u0905\u095c\u0938\u0920",
        "\u0909\u0928\u0939\u0924\u094d\u0924\u0930",
        "\u0938\u0924\u094d\u0924\u0930",
        "\u090f\u0915\u0939\u0924\u094d\u0924\u0930",
        "\u092c\u0939\u0924\u094d\u0924\u0930",
        "\u0924\u093f\u0939\u0924\u094d\u0924\u0930",
        "\u091a\u094c\u0939\u0924\u094d\u0924\u0930",
        "\u092a\u091b\u0924\u094d\u0924\u0930",
        "\u091b\u093f\u0939\u0924\u094d\u0924\u0930",
        "\u0938\u0924\u0939\u0924\u094d\u0924\u0930",
        "\u0905\u0920\u0939\u0924\u094d\u0924\u0930",
        "\u0909\u0928\u094d\u0928\u093e\u0938\u0940",
        "\u0905\u0938\u094d\u0938\u0940",
        "\u0907\u0915\u094d\u092f\u093e\u0938\u0940",
        "\u092c\u092f\u093e\u0938\u0940",
        "\u0924\u093f\u0930\u093e\u0938\u0940",
        "\u091a\u094c\u0930\u093e\u0938\u0940",
        "\u092a\u091a\u093e\u0938\u0940",
        "\u091b\u093f\u092f\u093e\u0938\u0940",
        "\u0938\u0924\u094d\u0924\u093e\u0938\u0940",
        "\u0905\u0920\u094d\u0920\u093e\u0938\u0940",
        "\u0928\u0935\u093e\u0938\u0940",
        "\u0928\u092c\u094d\u092c\u0947",
        "\u0907\u0915\u094d\u092f\u093e\u0928\u0935\u0947",
        "\u092c\u093e\u0928\u0935\u0947",
        "\u0924\u093f\u0930\u093e\u0928\u0935\u0947",
        "\u091a\u094c\u0930\u093e\u0928\u0935\u0947",
        "\u092a\u091b\u093e\u0928\u0935\u0947",
        "\u091b\u094d\u092f\u093e\u0928\u0935\u0947",
        "\u0938\u0924\u094d\u0924\u093e\u0928\u0935\u0947",
        "\u0905\u0920\u094d\u0920\u093e\u0928\u0935\u0947",
        "\u0928\u093f\u0928\u094d\u092f\u093e\u0928\u0935\u0947",
    };

    internal static string Hindi(int number)
    {
        if (number < 0) return "";
        if (number == 0) return HindiSmall[0];
        var value = (number - 1) % 9999 + 1;
        var parts = new List<string>();
        if (value >= 1000) parts.Add(HindiSmall[value / 1000] + " \u0939\u091c\u093e\u0930");
        value %= 1000;
        if (value >= 100) parts.Add(HindiSmall[value / 100] + " \u0938\u094c");
        value %= 100;
        if (value > 0) parts.Add(HindiSmall[value]);
        return string.Join(" ", parts);
    }

    private static readonly string[] ThaiDigits =
    {
        "\u0e28\u0e39\u0e19\u0e22\u0e4c",
        "\u0e2b\u0e19\u0e36\u0e48\u0e07",
        "\u0e2a\u0e2d\u0e07",
        "\u0e2a\u0e32\u0e21",
        "\u0e2a\u0e35\u0e48",
        "\u0e2b\u0e49\u0e32",
        "\u0e2b\u0e01",
        "\u0e40\u0e08\u0e47\u0e14",
        "\u0e41\u0e1b\u0e14",
        "\u0e40\u0e01\u0e49\u0e32",
    };

    internal static string Thai(int number)
    {
        if (number == int.MinValue) return "";
        // Word displays zero beyond this range, including negative values.
        if (number <= -1000000000 || number >= 1000000000 || number == 0) return ThaiDigits[0];
        var negative = number < 0;
        var value = negative ? -number : number;
        var prefix = negative ? "\u0e25\u0e1a" : "";
        if (value < 1000000) return prefix + ThaiGroup(value, value > 10);
        var remainder = value % 1000000;
        return prefix + ThaiGroup(value / 1000000, value / 1000000 > 10) + "\u0e25\u0e49\u0e32\u0e19"
            + (remainder == 0 ? "" : ThaiGroup(remainder, true));
    }

    private static string ThaiGroup(int number, bool finalOneAsEd)
    {
        var text = new StringBuilder();
        var remaining = number;
        int[] divisors = { 100000, 10000, 1000, 100, 10, 1 };
        string[] units = { "\u0e41\u0e2a\u0e19", "\u0e2b\u0e21\u0e37\u0e48\u0e19", "\u0e1e\u0e31\u0e19", "\u0e23\u0e49\u0e2d\u0e22", "", "" };
        for (var i = 0; i < divisors.Length; i++)
        {
            var digit = remaining / divisors[i];
            remaining %= divisors[i];
            if (digit == 0) continue;
            if (divisors[i] == 10)
                text.Append(digit == 1 ? "\u0e2a\u0e34\u0e1a" : digit == 2 ? "\u0e22\u0e35\u0e48\u0e2a\u0e34\u0e1a" : ThaiDigits[digit] + "\u0e2a\u0e34\u0e1a");
            else if (divisors[i] == 1)
                text.Append(digit == 1 && finalOneAsEd ? "\u0e40\u0e2d\u0e47\u0e14" : ThaiDigits[digit]);
            else text.Append(ThaiDigits[digit]).Append(units[i]);
        }
        return text.ToString();
    }
}
