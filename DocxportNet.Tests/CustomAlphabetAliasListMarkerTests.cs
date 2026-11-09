using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Walker;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class CustomAlphabetAliasListMarkerTests
{
    // Word captures for regional alphabet aliases, including logical Unicode controls.
    [Theory]
    [InlineData("arabicAbjad", -2147483648, 0, "")]
    [InlineData("arabicAbjad", -1, 0, "")]
    [InlineData("arabicAbjad", 0, 0, "")]
    [InlineData("arabicAbjad", 1, 0, "\u200c\u0623")]
    [InlineData("arabicAbjad", 9, 0, "\u200c\u0637")]
    [InlineData("arabicAbjad", 10, 0, "\u200c\u064a")]
    [InlineData("arabicAbjad", 99, 0, "\u200c\u0633\u200c\u0633\u200c\u0633\u200c\u0633")]
    [InlineData("arabicAbjad", 100, 0, "\u200c\u0639\u200c\u0639\u200c\u0639\u200c\u0639")]
    [InlineData("arabicAbjad", 999, 0, "\u200c\u0642\u200c\u0642\u200c\u0642\u200c\u0642\u200c\u0642\u200c\u0642\u200c\u0642\u200c\u0642")]
    [InlineData("arabicAbjad", 1000, 0, "\u200c\u0631\u200c\u0631\u200c\u0631\u200c\u0631\u200c\u0631\u200c\u0631\u200c\u0631\u200c\u0631")]
    [InlineData("arabicAbjad", 9999, 0, "\u200c\u062c\u200c\u062c\u200c\u062c\u200c\u062c\u200c\u062c\u200c\u062c\u200c\u062c\u200c\u062c")]
    [InlineData("arabicAbjad", 10000, 0, "\u200c\u062f\u200c\u062f\u200c\u062f\u200c\u062f\u200c\u062f\u200c\u062f\u200c\u062f\u200c\u062f")]
    [InlineData("arabicAbjad", 32767, 0, "\u200c\u0632\u200c\u0632\u200c\u0632\u200c\u0632\u200c\u0632\u200c\u0632\u200c\u0632\u200c\u0632\u200c\u0632")]
    [InlineData("arabicAbjad", 32768, 0, "\u200c\u062d\u200c\u062d\u200c\u062d\u200c\u062d\u200c\u062d\u200c\u062d\u200c\u062d\u200c\u062d\u200c\u062d")]
    [InlineData("arabicAbjad", 65535, 0, "\u200c\u0633\u200c\u0633\u200c\u0633")]
    [InlineData("arabicAbjad", 999999, 0, "\u200c\u0632")]
    [InlineData("arabicAbjad", 2147483647, 0, "\u200c\u0633\u200c\u0633\u200c\u0633\u200c\u0633\u200c\u0633\u200c\u0633\u200c\u0633\u200c\u0633\u200c\u0633")]
    [InlineData("arabicAlpha", -2147483648, 0, "")]
    [InlineData("arabicAlpha", -1, 0, "")]
    [InlineData("arabicAlpha", 0, 0, "")]
    [InlineData("arabicAlpha", 1, 0, "\u0623\u200c")]
    [InlineData("arabicAlpha", 9, 0, "\u0630\u200c")]
    [InlineData("arabicAlpha", 10, 0, "\u0631\u200c")]
    [InlineData("arabicAlpha", 99, 0, "\u0636\u200c\u0636\u200c\u0636\u200c\u0636\u200c")]
    [InlineData("arabicAlpha", 100, 0, "\u0637\u200c\u0637\u200c\u0637\u200c\u0637\u200c")]
    [InlineData("arabicAlpha", 999, 0, "\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c")]
    [InlineData("arabicAlpha", 1000, 0, "\u0641\u200c\u0641\u200c\u0641\u200c\u0641\u200c\u0641\u200c\u0641\u200c\u0641\u200c\u0641\u200c")]
    [InlineData("arabicAlpha", 9999, 0, "\u062a\u200c\u062a\u200c\u062a\u200c\u062a\u200c\u062a\u200c\u062a\u200c\u062a\u200c\u062a\u200c")]
    [InlineData("arabicAlpha", 10000, 0, "\u062b\u200c\u062b\u200c\u062b\u200c\u062b\u200c\u062b\u200c\u062b\u200c\u062b\u200c\u062b\u200c")]
    [InlineData("arabicAlpha", 32767, 0, "\u062e\u200c\u062e\u200c\u062e\u200c\u062e\u200c\u062e\u200c\u062e\u200c\u062e\u200c\u062e\u200c\u062e\u200c")]
    [InlineData("arabicAlpha", 32768, 0, "\u062f\u200c\u062f\u200c\u062f\u200c\u062f\u200c\u062f\u200c\u062f\u200c\u062f\u200c\u062f\u200c\u062f\u200c")]
    [InlineData("arabicAlpha", 65535, 0, "\u0636\u200c\u0636\u200c\u0636\u200c")]
    [InlineData("arabicAlpha", 999999, 0, "\u062e\u200c")]
    [InlineData("arabicAlpha", 2147483647, 0, "\u0636\u200c\u0636\u200c\u0636\u200c\u0636\u200c\u0636\u200c\u0636\u200c\u0636\u200c\u0636\u200c\u0636\u200c")]
    [InlineData("hebrew1", -2147483648, 0, "")]
    [InlineData("hebrew1", -1, 0, "")]
    [InlineData("hebrew1", 0, 0, "")]
    [InlineData("hebrew1", 1, 0, "\u05d0")]
    [InlineData("hebrew1", 9, 0, "\u05d8")]
    [InlineData("hebrew1", 10, 0, "\u05d9")]
    [InlineData("hebrew1", 99, 0, "\u05e6\u05d8")]
    [InlineData("hebrew1", 100, 0, "\u05e7")]
    [InlineData("hebrew1", 999, 0, "\u05e8\u05d8\u05d5")]
    [InlineData("hebrew1", 1000, 0, "\u05e8\u05d8\u05d6")]
    [InlineData("hebrew1", 9999, 0, "\u05e7\u05e6\u05d8")]
    [InlineData("hebrew1", 10000, 0, "\u05e8")]
    [InlineData("hebrew1", 32767, 0, "\u05e8\u05dc\u05d0")]
    [InlineData("hebrew1", 32768, 0, "\u05e8\u05dc\u05d1")]
    [InlineData("hebrew1", 65535, 0, "\u05e2\u05d0")]
    [InlineData("hebrew1", 999999, 0, "\u05d6")]
    [InlineData("hebrew1", 2147483647, 0, "\u05e8\u05dc\u05d8")]
    [InlineData("hebrew2", -2147483648, 0, "")]
    [InlineData("hebrew2", -1, 0, "")]
    [InlineData("hebrew2", 0, 0, "")]
    [InlineData("hebrew2", 1, 0, "\u200f\u05d0")]
    [InlineData("hebrew2", 9, 0, "\u200f\u05d8")]
    [InlineData("hebrew2", 10, 0, "\u200f\u05d9")]
    [InlineData("hebrew2", 99, 0, "\u200f\u05ea\u05ea\u05ea\u05ea\u05db")]
    [InlineData("hebrew2", 100, 0, "\u200f\u05ea\u05ea\u05ea\u05ea\u05dc")]
    [InlineData("hebrew2", 999, 0, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e4")]
    [InlineData("hebrew2", 1000, 0, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e6")]
    [InlineData("hebrew2", 9999, 0, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05d0")]
    [InlineData("hebrew2", 10000, 0, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05d1")]
    [InlineData("hebrew2", 32767, 0, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05db")]
    [InlineData("hebrew2", 32768, 0, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05dc")]
    [InlineData("hebrew2", 65535, 0, "\u200f\u05ea\u05ea\u05ea\u05d4")]
    [InlineData("hebrew2", 999999, 0, "\u200f\u05d6")]
    [InlineData("hebrew2", 2147483647, 0, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e7")]
    [InlineData("hindiConsonants", -2147483648, 0, "")]
    [InlineData("hindiConsonants", -1, 0, "")]
    [InlineData("hindiConsonants", 0, 0, " ")]
    [InlineData("hindiConsonants", 1, 0, "\u0905")]
    [InlineData("hindiConsonants", 9, 0, "\u090d")]
    [InlineData("hindiConsonants", 10, 0, "\u090e")]
    [InlineData("hindiConsonants", 99, 0, "\u090d\u090d\u090d\u090d\u090d\u090d")]
    [InlineData("hindiConsonants", 100, 0, "\u090e\u090e\u090e\u090e\u090e\u090e")]
    [InlineData("hindiConsonants", 999, 0, "\u0913\u0913\u0913\u0913\u0913")]
    [InlineData("hindiConsonants", 1000, 0, "\u0914\u0914\u0914\u0914\u0914")]
    [InlineData("hindiConsonants", 9999, 0, "\u0905\u0905\u0905\u0905\u0905\u0905\u0905\u0905\u0905\u0905\u0905\u0905\u0905\u0905\u0905\u0905")]
    [InlineData("hindiConsonants", 10000, 0, "\u0906\u0906\u0906\u0906\u0906\u0906\u0906\u0906\u0906\u0906\u0906\u0906\u0906\u0906\u0906\u0906")]
    [InlineData("hindiConsonants", 32767, 0, "\u0909\u0909\u0909\u0909\u0909\u0909\u0909\u0909\u0909\u0909\u0909\u0909\u0909\u0909")]
    [InlineData("hindiConsonants", 32768, 0, "\u090a\u090a\u090a\u090a\u090a\u090a\u090a\u090a\u090a\u090a\u090a\u090a\u090a\u090a")]
    [InlineData("hindiConsonants", 65535, 0, "\u0911\u0911\u0911\u0911\u0911\u0911\u0911\u0911\u0911\u0911")]
    [InlineData("hindiConsonants", 999999, 0, "\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902")]
    [InlineData("hindiConsonants", 2147483647, 0, "\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902\u0905\u0902")]
    [InlineData("hindiVowels", -2147483648, 0, "")]
    [InlineData("hindiVowels", -1, 0, "")]
    [InlineData("hindiVowels", 0, 0, " ")]
    [InlineData("hindiVowels", 1, 0, "\u0915")]
    [InlineData("hindiVowels", 9, 0, "\u091d")]
    [InlineData("hindiVowels", 10, 0, "\u091e")]
    [InlineData("hindiVowels", 99, 0, "\u092d\u092d\u092d")]
    [InlineData("hindiVowels", 100, 0, "\u092e\u092e\u092e")]
    [InlineData("hindiVowels", 999, 0, "\u0921\u0921\u0921")]
    [InlineData("hindiVowels", 1000, 0, "\u0922\u0922\u0922")]
    [InlineData("hindiVowels", 9999, 0, "\u0920\u0920\u0920\u0920\u0920\u0920\u0920\u0920")]
    [InlineData("hindiVowels", 10000, 0, "\u0921\u0921\u0921\u0921\u0921\u0921\u0921\u0921")]
    [InlineData("hindiVowels", 32767, 0, "\u0925\u0925\u0925\u0925\u0925\u0925\u0925")]
    [InlineData("hindiVowels", 32768, 0, "\u0926\u0926\u0926\u0926\u0926\u0926\u0926")]
    [InlineData("hindiVowels", 65535, 0, "\u092f\u092f\u092f\u092f\u092f")]
    [InlineData("hindiVowels", 999999, 0, "\u0934\u0934\u0934\u0934")]
    [InlineData("hindiVowels", 2147483647, 0, "\u0930\u0930\u0930\u0930\u0930\u0930\u0930\u0930")]
    [InlineData("russianLower", -2147483648, 0, "")]
    [InlineData("russianLower", -1, 0, "")]
    [InlineData("russianLower", 0, 0, " ")]
    [InlineData("russianLower", 1, 0, "\u0430")]
    [InlineData("russianLower", 9, 0, "\u0438")]
    [InlineData("russianLower", 10, 0, "\u043a")]
    [InlineData("russianLower", 99, 0, "\u043c\u043c\u043c\u043c")]
    [InlineData("russianLower", 100, 0, "\u043d\u043d\u043d\u043d")]
    [InlineData("russianLower", 999, 0, "\u043d\u043d\u043d\u043d\u043d")]
    [InlineData("russianLower", 1000, 0, "\u043e\u043e\u043e\u043e\u043e")]
    [InlineData("russianLower", 9999, 0, "\u0447\u0447\u0447\u0447\u0447\u0447\u0447\u0447\u0447\u0447\u0447\u0447\u0447\u0447\u0447")]
    [InlineData("russianLower", 10000, 0, "\u0448\u0448\u0448\u0448\u0448\u0448\u0448\u0448\u0448\u0448\u0448\u0448\u0448\u0448\u0448")]
    [InlineData("russianLower", 32767, 0, "\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b\u044b")]
    [InlineData("russianLower", 32768, 0, "\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d\u044d")]
    [InlineData("russianLower", 65535, 0, "\u0448\u0448\u0448\u0448\u0448\u0448\u0448\u0448\u0448\u0448")]
    [InlineData("russianLower", 999999, 0, "\u0445\u0445\u0445\u0445\u0445\u0445\u0445\u0445\u0445\u0445\u0445\u0445\u0445")]
    [InlineData("russianLower", 2147483647, 0, "\u0436")]
    [InlineData("russianUpper", -2147483648, 0, "")]
    [InlineData("russianUpper", -1, 0, "")]
    [InlineData("russianUpper", 0, 0, " ")]
    [InlineData("russianUpper", 1, 0, "\u0410")]
    [InlineData("russianUpper", 9, 0, "\u0418")]
    [InlineData("russianUpper", 10, 0, "\u041a")]
    [InlineData("russianUpper", 99, 0, "\u041c\u041c\u041c\u041c")]
    [InlineData("russianUpper", 100, 0, "\u041d\u041d\u041d\u041d")]
    [InlineData("russianUpper", 999, 0, "\u041d\u041d\u041d\u041d\u041d")]
    [InlineData("russianUpper", 1000, 0, "\u041e\u041e\u041e\u041e\u041e")]
    [InlineData("russianUpper", 9999, 0, "\u0427\u0427\u0427\u0427\u0427\u0427\u0427\u0427\u0427\u0427\u0427\u0427\u0427\u0427\u0427")]
    [InlineData("russianUpper", 10000, 0, "\u0428\u0428\u0428\u0428\u0428\u0428\u0428\u0428\u0428\u0428\u0428\u0428\u0428\u0428\u0428")]
    [InlineData("russianUpper", 32767, 0, "\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b\u042b")]
    [InlineData("russianUpper", 32768, 0, "\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d\u042d")]
    [InlineData("russianUpper", 65535, 0, "\u0428\u0428\u0428\u0428\u0428\u0428\u0428\u0428\u0428\u0428")]
    [InlineData("russianUpper", 999999, 0, "\u0425\u0425\u0425\u0425\u0425\u0425\u0425\u0425\u0425\u0425\u0425\u0425\u0425")]
    [InlineData("russianUpper", 2147483647, 0, "\u0416")]
    [InlineData("thaiLetters", -2147483648, 0, "")]
    [InlineData("thaiLetters", -1, 0, "")]
    [InlineData("thaiLetters", 0, 0, " ")]
    [InlineData("thaiLetters", 1, 0, "\u0e01")]
    [InlineData("thaiLetters", 9, 0, "\u0e0c")]
    [InlineData("thaiLetters", 10, 0, "\u0e0d")]
    [InlineData("thaiLetters", 99, 0, "\u0e14\u0e14\u0e14")]
    [InlineData("thaiLetters", 100, 0, "\u0e15\u0e15\u0e15")]
    [InlineData("thaiLetters", 999, 0, "\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12")]
    [InlineData("thaiLetters", 1000, 0, "\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13")]
    [InlineData("thaiLetters", 9999, 0, "\u0e29\u0e29\u0e29\u0e29")]
    [InlineData("thaiLetters", 10000, 0, "\u0e2a\u0e2a\u0e2a\u0e2a")]
    [InlineData("thaiLetters", 32767, 0, "\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b")]
    [InlineData("thaiLetters", 32768, 0, "\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c")]
    [InlineData("thaiLetters", 65535, 0, "\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14")]
    [InlineData("thaiLetters", 999999, 0, "\u0e0c")]
    [InlineData("thaiLetters", 2147483647, 0, "\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var abstractXml = name switch
        {
            "arabicAbjad" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u0623, \u0628, \u062c, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "arabicAlpha" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u0623, \u0628, \u062a, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "hebrew1" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u05d0, \u05d9, \u05e7, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "hebrew2" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u05d0, \u05d1, \u05d2, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "hindiConsonants" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u0905, \u0906, \u0907, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "hindiVowels" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u0915, \u0916, \u0917, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "russianLower" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u0430, \u0431, \u0432, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "russianUpper" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u0410, \u0411, \u0412, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "thaiLetters" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u0e01, \u0e02, \u0e04, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            _ => throw new ArgumentException(name)
        };
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var numbering = new Numbering(new AbstractNum(abstractXml),
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 },
                new LevelOverride(new StartOverrideNumberingValue { Val = start }) { LevelIndex = level }) { NumberID = 2 });
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
        var item = new Paragraph(new ParagraphProperties(new NumberingProperties(
            new NumberingLevelReference { Val = level }, new NumberingId { Val = 2 })), new Run(new Text("Body")));
        main.Document = new Document(new Body(item));
        var original = numbering.OuterXml;
        DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
        foreach (var visitor in visitors)
        {
            var output = DxpExport.ExportToString(doc, visitor);
            if (visitor is DxpHtmlVisitor) output = new HtmlParser().ParseDocument(output).Body!.TextContent;
            var matches = Regex.Matches(output, "BEGIN(.*?)END");
            Assert.Equal(expected, matches.Count == 0 ? null : matches[matches.Count - 1].Groups[1].Value);
            Assert.Equal(original, numbering.OuterXml);
        }
    }
}
