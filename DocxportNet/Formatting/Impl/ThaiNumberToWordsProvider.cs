using System.Globalization;

namespace DocxportNet.Formatting.Impl;

public sealed class ThaiNumberToWordsProvider : DxpINumberToWordsProvider
{
    public bool CanHandle(CultureInfo culture) => culture.TwoLetterISOLanguageName.Equals("th", StringComparison.OrdinalIgnoreCase);

    public string ToCardinal(int number)
    {
        if (number == 0)
            return "ศูนย์";
        if (number < 0)
            return "ลบ " + ToCardinal(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);

        return ToThaiNumber(number);
    }

    public string ToOrdinalWords(int number)
    {
        if (number == 0)
            return "ที่ศูนย์";
        if (number < 0)
            return "ที่ลบ " + ToCardinal(Math.Abs(number));
        if (number > 32767)
            return number.ToString(CultureInfo.InvariantCulture);

        return "ที่" + ToCardinal(number);
    }

    public string ToDollarText(double number)
    {
        double abs = Math.Abs(number);
        long dollars = (long)Math.Floor(abs);
        int cents = (int)Math.Round((abs - dollars) * 100, MidpointRounding.AwayFromZero);
        if (cents == 100)
        {
            dollars += 1;
            cents = 0;
        }

        string words = ToCardinal((int)dollars);
        string centsText = cents.ToString("00", CultureInfo.InvariantCulture) + "/100";
        string result = words + " และ " + centsText;
        if (number < 0)
            result = "ลบ " + result;
        return result;
    }

    private static string ToThaiNumber(int number)
    {
        if (number == 0)
            return "ศูนย์";

        var parts = new List<string>();
        int millionGroup = 0;
        int remaining = number;
        while (remaining > 0)
        {
            int group = remaining % 1000000;
            if (group > 0)
            {
                string groupWords = ToThaiBelowMillion(group);
                if (millionGroup > 0)
                    groupWords += "ล้าน";
                parts.Insert(0, groupWords);
            }
            remaining /= 1000000;
            millionGroup++;
        }
        return string.Join(string.Empty, parts);
    }

    private static string ToThaiBelowMillion(int number)
    {
        string[] digits =
        {
            "ศูนย์","หนึ่ง","สอง","สาม","สี่","ห้า","หก","เจ็ด","แปด","เก้า"
        };

        string[] units =
        {
            "", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"
        };

        var result = new List<string>();
        int position = 0;
        int n = number;
        while (n > 0)
        {
            int digit = n % 10;
            if (digit != 0)
            {
                string word;
                if (position == 0)
                {
                    word = digit == 1 && number > 10 ? "เอ็ด" : digits[digit];
                }
                else if (position == 1)
                {
                    if (digit == 1)
                        word = "สิบ";
                    else if (digit == 2)
                        word = "ยี่สิบ";
                    else
                        word = digits[digit] + units[position];
                }
                else
                {
                    word = digits[digit] + units[position];
                }
                result.Insert(0, word);
            }
            n /= 10;
            position++;
        }
        return string.Join(string.Empty, result);
    }
}
