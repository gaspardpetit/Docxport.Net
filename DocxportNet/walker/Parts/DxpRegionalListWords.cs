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

    private static readonly string[] VietnameseSmall =
    {
        "kh\u00f4ng",
        "m\u00f4\u0323t",
        "hai",
        "ba",
        "b\u00f4\u0301n",
        "n\u0103m",
        "sa\u0301u",
        "ba\u0309y",
        "ta\u0301m",
        "chi\u0301n",
        "m\u01b0\u01a1\u0300i",
        "m\u01b0\u01a1\u0300i m\u00f4\u0323t",
        "m\u01b0\u01a1\u0300i hai",
        "m\u01b0\u01a1\u0300i ba",
        "m\u01b0\u01a1\u0300i b\u00f4\u0301n",
        "m\u01b0\u01a1\u0300i l\u0103m",
        "m\u01b0\u01a1\u0300i sa\u0301u",
        "m\u01b0\u01a1\u0300i ba\u0309y",
        "m\u01b0\u01a1\u0300i ta\u0301m",
        "m\u01b0\u01a1\u0300i chi\u0301n",
        "hai m\u01b0\u01a1i",
        "hai m\u01b0\u01a1i m\u00f4\u0301t",
        "hai m\u01b0\u01a1i hai",
        "hai m\u01b0\u01a1i ba",
        "hai m\u01b0\u01a1i b\u00f4\u0301n",
        "hai m\u01b0\u01a1i l\u0103m",
        "hai m\u01b0\u01a1i sa\u0301u",
        "hai m\u01b0\u01a1i ba\u0309y",
        "hai m\u01b0\u01a1i ta\u0301m",
        "hai m\u01b0\u01a1i chi\u0301n",
        "ba m\u01b0\u01a1i",
        "ba m\u01b0\u01a1i m\u00f4\u0301t",
        "ba m\u01b0\u01a1i hai",
        "ba m\u01b0\u01a1i ba",
        "ba m\u01b0\u01a1i b\u00f4\u0301n",
        "ba m\u01b0\u01a1i l\u0103m",
        "ba m\u01b0\u01a1i sa\u0301u",
        "ba m\u01b0\u01a1i ba\u0309y",
        "ba m\u01b0\u01a1i ta\u0301m",
        "ba m\u01b0\u01a1i chi\u0301n",
        "b\u00f4\u0301n m\u01b0\u01a1i",
        "b\u00f4\u0301n m\u01b0\u01a1i m\u00f4\u0301t",
        "b\u00f4\u0301n m\u01b0\u01a1i hai",
        "b\u00f4\u0301n m\u01b0\u01a1i ba",
        "b\u00f4\u0301n m\u01b0\u01a1i b\u00f4\u0301n",
        "b\u00f4\u0301n m\u01b0\u01a1i l\u0103m",
        "b\u00f4\u0301n m\u01b0\u01a1i sa\u0301u",
        "b\u00f4\u0301n m\u01b0\u01a1i ba\u0309y",
        "b\u00f4\u0301n m\u01b0\u01a1i ta\u0301m",
        "b\u00f4\u0301n m\u01b0\u01a1i chi\u0301n",
        "n\u0103m m\u01b0\u01a1i",
        "n\u0103m m\u01b0\u01a1i m\u00f4\u0301t",
        "n\u0103m m\u01b0\u01a1i hai",
        "n\u0103m m\u01b0\u01a1i ba",
        "n\u0103m m\u01b0\u01a1i b\u00f4\u0301n",
        "n\u0103m m\u01b0\u01a1i l\u0103m",
        "n\u0103m m\u01b0\u01a1i sa\u0301u",
        "n\u0103m m\u01b0\u01a1i ba\u0309y",
        "n\u0103m m\u01b0\u01a1i ta\u0301m",
        "n\u0103m m\u01b0\u01a1i chi\u0301n",
        "sa\u0301u m\u01b0\u01a1i",
        "sa\u0301u m\u01b0\u01a1i m\u00f4\u0301t",
        "sa\u0301u m\u01b0\u01a1i hai",
        "sa\u0301u m\u01b0\u01a1i ba",
        "sa\u0301u m\u01b0\u01a1i b\u00f4\u0301n",
        "sa\u0301u m\u01b0\u01a1i l\u0103m",
        "sa\u0301u m\u01b0\u01a1i sa\u0301u",
        "sa\u0301u m\u01b0\u01a1i ba\u0309y",
        "sa\u0301u m\u01b0\u01a1i ta\u0301m",
        "sa\u0301u m\u01b0\u01a1i chi\u0301n",
        "ba\u0309y m\u01b0\u01a1i",
        "ba\u0309y m\u01b0\u01a1i m\u00f4\u0301t",
        "ba\u0309y m\u01b0\u01a1i hai",
        "ba\u0309y m\u01b0\u01a1i ba",
        "ba\u0309y m\u01b0\u01a1i b\u00f4\u0301n",
        "ba\u0309y m\u01b0\u01a1i l\u0103m",
        "ba\u0309y m\u01b0\u01a1i sa\u0301u",
        "ba\u0309y m\u01b0\u01a1i ba\u0309y",
        "ba\u0309y m\u01b0\u01a1i ta\u0301m",
        "ba\u0309y m\u01b0\u01a1i chi\u0301n",
        "ta\u0301m m\u01b0\u01a1i",
        "ta\u0301m m\u01b0\u01a1i m\u00f4\u0301t",
        "ta\u0301m m\u01b0\u01a1i hai",
        "ta\u0301m m\u01b0\u01a1i ba",
        "ta\u0301m m\u01b0\u01a1i b\u00f4\u0301n",
        "ta\u0301m m\u01b0\u01a1i l\u0103m",
        "ta\u0301m m\u01b0\u01a1i sa\u0301u",
        "ta\u0301m m\u01b0\u01a1i ba\u0309y",
        "ta\u0301m m\u01b0\u01a1i ta\u0301m",
        "ta\u0301m m\u01b0\u01a1i chi\u0301n",
        "chi\u0301n m\u01b0\u01a1i",
        "chi\u0301n m\u01b0\u01a1i m\u00f4\u0301t",
        "chi\u0301n m\u01b0\u01a1i hai",
        "chi\u0301n m\u01b0\u01a1i ba",
        "chi\u0301n m\u01b0\u01a1i b\u00f4\u0301n",
        "chi\u0301n m\u01b0\u01a1i l\u0103m",
        "chi\u0301n m\u01b0\u01a1i sa\u0301u",
        "chi\u0301n m\u01b0\u01a1i ba\u0309y",
        "chi\u0301n m\u01b0\u01a1i ta\u0301m",
        "chi\u0301n m\u01b0\u01a1i chi\u0301n",
    };

    internal static string Vietnamese(int number)
    {
        if (number < 0) return "";
        if (number == 0) return VietnameseSmall[0];
        var value = (number - 1) % 1000 + 1;
        if (value == 1000) return "m\u00f4\u0323t nga\u0300n";
        if (value < 100) return VietnameseSmall[value];
        var text = VietnameseSmall[value / 100] + " tr\u0103m";
        var remainder = value % 100;
        return remainder == 0 ? text : text + (remainder < 10 ? " le\u0309 " : " ") + VietnameseSmall[remainder];
    }
}
