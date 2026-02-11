using System.Collections.Generic;
using System.Globalization;

namespace DocxportNet.Fields.Resolution.Impl;

public sealed class JapaneseMergeMacroProvider : SimpleMergeMacroProvider
{
    protected override string LanguageCode => "ja";
    protected override string Salutation => "親愛なる";
    protected override GreetingStyle Greeting => GreetingStyle.NameOnly;
    protected override bool IncludeTitleInGreeting => false;
    protected override string? NameSuffix => "様";
    protected override string GreetingPunctuation => string.Empty;
    protected override string DefaultGreeting => "ご担当者様";
    protected override bool LastNameFirst => true;
    protected override LocalityOrder LocalityLayout => LocalityOrder.PostalStateCity;

    protected override string? ResolveAddressBlock(IDxpMergeRecordCursor cursor, CultureInfo culture)
    {
        var company = Normalize(GetFieldString(cursor, "Company", culture));
        var address1 = Normalize(GetFieldString(cursor, "Address1", culture));
        var address2 = Normalize(GetFieldString(cursor, "Address2", culture));
        var city = Normalize(GetFieldString(cursor, "City", culture));
        var state = Normalize(GetFieldString(cursor, "State", culture));
        var postalCode = Normalize(GetFieldString(cursor, "PostalCode", culture));
        var country = Normalize(GetFieldString(cursor, "Country", culture));

        var lines = new List<string>();
        if (company is { Length: > 0 })
            lines.Add(company);
        if (postalCode is { Length: > 0 })
            lines.Add($"〒{postalCode}");

        var locality = JoinNonEmpty(string.Empty, state, city);
        if (locality is { Length: > 0 })
            lines.Add(locality);
        if (address1 is { Length: > 0 })
            lines.Add(address1);
        if (address2 is { Length: > 0 })
            lines.Add(address2);
        if (country is { Length: > 0 })
            lines.Add(country);

        if (lines.Count == 0)
            return null;

        return string.Join("\n", lines);
    }
}
