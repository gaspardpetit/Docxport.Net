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

public sealed class ThaiLetterListMarkerTests
{
    [Fact]
    public void SharedFieldFormatterRetainsDecimalFallback()
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, NumberFormatValues.ThaiLetters));
        Assert.Equal("1231", DxpListMarkerFormatter.FormatNumber(1231, NumberFormatValues.ThaiLetters));
    }

    // Word ListString captures: preserve logical Unicode including invisible bidi and joining controls.
    [Theory]
    [InlineData("thaiLetters", -2147483648, true, "")]
    [InlineData("thaiLetters", -1, true, "")]
    [InlineData("thaiLetters", 0, true, " ")]
    [InlineData("thaiLetters", 1, true, "\u0e01")]
    [InlineData("thaiLetters", 2, true, "\u0e02")]
    [InlineData("thaiLetters", 3, true, "\u0e04")]
    [InlineData("thaiLetters", 4, true, "\u0e07")]
    [InlineData("thaiLetters", 5, true, "\u0e08")]
    [InlineData("thaiLetters", 6, true, "\u0e09")]
    [InlineData("thaiLetters", 7, true, "\u0e0a")]
    [InlineData("thaiLetters", 8, true, "\u0e0b")]
    [InlineData("thaiLetters", 9, true, "\u0e0c")]
    [InlineData("thaiLetters", 10, true, "\u0e0d")]
    [InlineData("thaiLetters", 11, true, "\u0e0e")]
    [InlineData("thaiLetters", 12, true, "\u0e0f")]
    [InlineData("thaiLetters", 13, true, "\u0e10")]
    [InlineData("thaiLetters", 14, true, "\u0e11")]
    [InlineData("thaiLetters", 15, true, "\u0e12")]
    [InlineData("thaiLetters", 16, true, "\u0e13")]
    [InlineData("thaiLetters", 17, true, "\u0e14")]
    [InlineData("thaiLetters", 18, true, "\u0e15")]
    [InlineData("thaiLetters", 19, true, "\u0e16")]
    [InlineData("thaiLetters", 20, true, "\u0e17")]
    [InlineData("thaiLetters", 21, true, "\u0e18")]
    [InlineData("thaiLetters", 22, true, "\u0e19")]
    [InlineData("thaiLetters", 23, true, "\u0e1a")]
    [InlineData("thaiLetters", 24, true, "\u0e1b")]
    [InlineData("thaiLetters", 25, true, "\u0e1c")]
    [InlineData("thaiLetters", 26, true, "\u0e1d")]
    [InlineData("thaiLetters", 27, true, "\u0e1e")]
    [InlineData("thaiLetters", 28, true, "\u0e1f")]
    [InlineData("thaiLetters", 29, true, "\u0e20")]
    [InlineData("thaiLetters", 30, true, "\u0e21")]
    [InlineData("thaiLetters", 31, true, "\u0e22")]
    [InlineData("thaiLetters", 32, true, "\u0e23")]
    [InlineData("thaiLetters", 33, true, "\u0e25")]
    [InlineData("thaiLetters", 34, true, "\u0e27")]
    [InlineData("thaiLetters", 35, true, "\u0e28")]
    [InlineData("thaiLetters", 36, true, "\u0e29")]
    [InlineData("thaiLetters", 37, true, "\u0e2a")]
    [InlineData("thaiLetters", 38, true, "\u0e2b")]
    [InlineData("thaiLetters", 39, true, "\u0e2c")]
    [InlineData("thaiLetters", 40, true, "\u0e2d")]
    [InlineData("thaiLetters", 41, true, "\u0e2e")]
    [InlineData("thaiLetters", 42, true, "\u0e01\u0e01")]
    [InlineData("thaiLetters", 43, true, "\u0e02\u0e02")]
    [InlineData("thaiLetters", 44, true, "\u0e04\u0e04")]
    [InlineData("thaiLetters", 45, true, "\u0e07\u0e07")]
    [InlineData("thaiLetters", 46, true, "\u0e08\u0e08")]
    [InlineData("thaiLetters", 47, true, "\u0e09\u0e09")]
    [InlineData("thaiLetters", 48, true, "\u0e0a\u0e0a")]
    [InlineData("thaiLetters", 49, true, "\u0e0b\u0e0b")]
    [InlineData("thaiLetters", 50, true, "\u0e0c\u0e0c")]
    [InlineData("thaiLetters", 51, true, "\u0e0d\u0e0d")]
    [InlineData("thaiLetters", 52, true, "\u0e0e\u0e0e")]
    [InlineData("thaiLetters", 53, true, "\u0e0f\u0e0f")]
    [InlineData("thaiLetters", 54, true, "\u0e10\u0e10")]
    [InlineData("thaiLetters", 55, true, "\u0e11\u0e11")]
    [InlineData("thaiLetters", 56, true, "\u0e12\u0e12")]
    [InlineData("thaiLetters", 57, true, "\u0e13\u0e13")]
    [InlineData("thaiLetters", 58, true, "\u0e14\u0e14")]
    [InlineData("thaiLetters", 59, true, "\u0e15\u0e15")]
    [InlineData("thaiLetters", 60, true, "\u0e16\u0e16")]
    [InlineData("thaiLetters", 61, true, "\u0e17\u0e17")]
    [InlineData("thaiLetters", 62, true, "\u0e18\u0e18")]
    [InlineData("thaiLetters", 63, true, "\u0e19\u0e19")]
    [InlineData("thaiLetters", 64, true, "\u0e1a\u0e1a")]
    [InlineData("thaiLetters", 65, true, "\u0e1b\u0e1b")]
    [InlineData("thaiLetters", 66, true, "\u0e1c\u0e1c")]
    [InlineData("thaiLetters", 67, true, "\u0e1d\u0e1d")]
    [InlineData("thaiLetters", 68, true, "\u0e1e\u0e1e")]
    [InlineData("thaiLetters", 69, true, "\u0e1f\u0e1f")]
    [InlineData("thaiLetters", 70, true, "\u0e20\u0e20")]
    [InlineData("thaiLetters", 71, true, "\u0e21\u0e21")]
    [InlineData("thaiLetters", 72, true, "\u0e22\u0e22")]
    [InlineData("thaiLetters", 73, true, "\u0e23\u0e23")]
    [InlineData("thaiLetters", 74, true, "\u0e25\u0e25")]
    [InlineData("thaiLetters", 75, true, "\u0e27\u0e27")]
    [InlineData("thaiLetters", 76, true, "\u0e28\u0e28")]
    [InlineData("thaiLetters", 77, true, "\u0e29\u0e29")]
    [InlineData("thaiLetters", 78, true, "\u0e2a\u0e2a")]
    [InlineData("thaiLetters", 79, true, "\u0e2b\u0e2b")]
    [InlineData("thaiLetters", 80, true, "\u0e2c\u0e2c")]
    [InlineData("thaiLetters", 81, true, "\u0e2d\u0e2d")]
    [InlineData("thaiLetters", 82, true, "\u0e2e\u0e2e")]
    [InlineData("thaiLetters", 83, true, "\u0e01\u0e01\u0e01")]
    [InlineData("thaiLetters", 84, true, "\u0e02\u0e02\u0e02")]
    [InlineData("thaiLetters", 85, true, "\u0e04\u0e04\u0e04")]
    [InlineData("thaiLetters", 86, true, "\u0e07\u0e07\u0e07")]
    [InlineData("thaiLetters", 87, true, "\u0e08\u0e08\u0e08")]
    [InlineData("thaiLetters", 88, true, "\u0e09\u0e09\u0e09")]
    [InlineData("thaiLetters", 89, true, "\u0e0a\u0e0a\u0e0a")]
    [InlineData("thaiLetters", 90, true, "\u0e0b\u0e0b\u0e0b")]
    [InlineData("thaiLetters", 91, true, "\u0e0c\u0e0c\u0e0c")]
    [InlineData("thaiLetters", 92, true, "\u0e0d\u0e0d\u0e0d")]
    [InlineData("thaiLetters", 93, true, "\u0e0e\u0e0e\u0e0e")]
    [InlineData("thaiLetters", 94, true, "\u0e0f\u0e0f\u0e0f")]
    [InlineData("thaiLetters", 95, true, "\u0e10\u0e10\u0e10")]
    [InlineData("thaiLetters", 96, true, "\u0e11\u0e11\u0e11")]
    [InlineData("thaiLetters", 97, true, "\u0e12\u0e12\u0e12")]
    [InlineData("thaiLetters", 98, true, "\u0e13\u0e13\u0e13")]
    [InlineData("thaiLetters", 99, true, "\u0e14\u0e14\u0e14")]
    [InlineData("thaiLetters", 100, true, "\u0e15\u0e15\u0e15")]
    [InlineData("thaiLetters", 114, true, "\u0e23\u0e23\u0e23")]
    [InlineData("thaiLetters", 115, true, "\u0e25\u0e25\u0e25")]
    [InlineData("thaiLetters", 116, true, "\u0e27\u0e27\u0e27")]
    [InlineData("thaiLetters", 299, true, "\u0e0f\u0e0f\u0e0f\u0e0f\u0e0f\u0e0f\u0e0f\u0e0f")]
    [InlineData("thaiLetters", 300, true, "\u0e10\u0e10\u0e10\u0e10\u0e10\u0e10\u0e10\u0e10")]
    [InlineData("thaiLetters", 301, true, "\u0e11\u0e11\u0e11\u0e11\u0e11\u0e11\u0e11\u0e11")]
    [InlineData("thaiLetters", 303, true, "\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13")]
    [InlineData("thaiLetters", 304, true, "\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14")]
    [InlineData("thaiLetters", 305, true, "\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15")]
    [InlineData("thaiLetters", 390, true, "\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18")]
    [InlineData("thaiLetters", 391, true, "\u0e19\u0e19\u0e19\u0e19\u0e19\u0e19\u0e19\u0e19\u0e19\u0e19")]
    [InlineData("thaiLetters", 392, true, "\u0e1a\u0e1a\u0e1a\u0e1a\u0e1a\u0e1a\u0e1a\u0e1a\u0e1a\u0e1a")]
    [InlineData("thaiLetters", 393, true, "\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b")]
    [InlineData("thaiLetters", 399, true, "\u0e21\u0e21\u0e21\u0e21\u0e21\u0e21\u0e21\u0e21\u0e21\u0e21")]
    [InlineData("thaiLetters", 400, true, "\u0e22\u0e22\u0e22\u0e22\u0e22\u0e22\u0e22\u0e22\u0e22\u0e22")]
    [InlineData("thaiLetters", 401, true, "\u0e23\u0e23\u0e23\u0e23\u0e23\u0e23\u0e23\u0e23\u0e23\u0e23")]
    [InlineData("thaiLetters", 539, true, "\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09")]
    [InlineData("thaiLetters", 540, true, "\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a")]
    [InlineData("thaiLetters", 541, true, "\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b")]
    [InlineData("thaiLetters", 557, true, "\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b")]
    [InlineData("thaiLetters", 558, true, "\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c")]
    [InlineData("thaiLetters", 559, true, "\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d")]
    [InlineData("thaiLetters", 607, true, "\u0e25\u0e25\u0e25\u0e25\u0e25\u0e25\u0e25\u0e25\u0e25\u0e25\u0e25\u0e25\u0e25\u0e25\u0e25")]
    [InlineData("thaiLetters", 608, true, "\u0e27\u0e27\u0e27\u0e27\u0e27\u0e27\u0e27\u0e27\u0e27\u0e27\u0e27\u0e27\u0e27\u0e27\u0e27")]
    [InlineData("thaiLetters", 609, true, "\u0e28\u0e28\u0e28\u0e28\u0e28\u0e28\u0e28\u0e28\u0e28\u0e28\u0e28\u0e28\u0e28\u0e28\u0e28")]
    [InlineData("thaiLetters", 659, true, "\u0e04\u0e04\u0e04\u0e04\u0e04\u0e04\u0e04\u0e04\u0e04\u0e04\u0e04\u0e04\u0e04\u0e04\u0e04\u0e04\u0e04")]
    [InlineData("thaiLetters", 660, true, "\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07")]
    [InlineData("thaiLetters", 661, true, "\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08")]
    [InlineData("thaiLetters", 681, true, "\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c")]
    [InlineData("thaiLetters", 682, true, "\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d")]
    [InlineData("thaiLetters", 683, true, "\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e\u0e1e")]
    [InlineData("thaiLetters", 783, true, "\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07\u0e07")]
    [InlineData("thaiLetters", 784, true, "\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08\u0e08")]
    [InlineData("thaiLetters", 785, true, "\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09")]
    [InlineData("thaiLetters", 839, true, "\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16")]
    [InlineData("thaiLetters", 840, true, "\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17")]
    [InlineData("thaiLetters", 841, true, "\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18\u0e18")]
    [InlineData("thaiLetters", 867, true, "\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09\u0e09")]
    [InlineData("thaiLetters", 868, true, "\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a\u0e0a")]
    [InlineData("thaiLetters", 869, true, "\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b")]
    [InlineData("thaiLetters", 999, true, "\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12\u0e12")]
    [InlineData("thaiLetters", 1000, true, "\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13\u0e13")]
    [InlineData("thaiLetters", 1001, true, "\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14")]
    [InlineData("thaiLetters", 1049, true, "\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b\u0e1b")]
    [InlineData("thaiLetters", 1050, true, "\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c\u0e1c")]
    [InlineData("thaiLetters", 1051, true, "\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d\u0e1d")]
    [InlineData("thaiLetters", 1084, true, "\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15\u0e15")]
    [InlineData("thaiLetters", 1085, true, "\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16\u0e16")]
    [InlineData("thaiLetters", 1086, true, "\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17\u0e17")]
    [InlineData("thaiLetters", 1229, true, "\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d")]
    [InlineData("thaiLetters", 1230, true, "\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e")]
    [InlineData("thaiLetters", 1231, true, "\u0e01")]
    [InlineData("thaiLetters", 1270, true, "\u0e2d")]
    [InlineData("thaiLetters", 1271, true, "\u0e2e")]
    [InlineData("thaiLetters", 1272, true, "\u0e01\u0e01")]
    [InlineData("thaiLetters", 2459, true, "\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d")]
    [InlineData("thaiLetters", 2460, true, "\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e")]
    [InlineData("thaiLetters", 2461, true, "\u0e01")]
    [InlineData("thaiLetters", 32767, true, "\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b\u0e0b")]
    [InlineData("thaiLetters", 32768, true, "\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c\u0e0c")]
    [InlineData("thaiLetters", 65535, true, "\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14\u0e14")]
    [InlineData("thaiLetters", 1000000, true, "\u0e0d")]
    [InlineData("thaiLetters", 2147483647, true, "\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b\u0e2b")]
    public void RegionalAlphabetsMatchWordAtBoundaries(string format, int start, bool startOverride, string? expected)
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var numbering = new Numbering(new AbstractNum(new MultiLevelType { Val = MultiLevelValues.Multilevel },
            new Level(new StartNumberingValue { Val = startOverride ? 1 : start },
                new NumberingFormat { Val = new NumberFormatValues(format) },
                new LevelText { Val = "BEGIN%1END" }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
        if (startOverride)
            numbering.Append(new NumberingInstance(new AbstractNumId { Val = 1 },
                new LevelOverride(new StartOverrideNumberingValue { Val = start }) { LevelIndex = 0 }) { NumberID = 2 });
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
        Paragraph Item(int numId) => new(new ParagraphProperties(new NumberingProperties(
            new NumberingLevelReference { Val = 0 }, new NumberingId { Val = numId })), new Run(new Text("Body")));
        var body = new Body(Item(1));
        if (startOverride) body.Append(Item(2));
        main.Document = new Document(body);
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
