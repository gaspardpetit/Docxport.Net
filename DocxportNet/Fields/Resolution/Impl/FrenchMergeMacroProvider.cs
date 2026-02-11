using System.Collections.Generic;
using System.Globalization;

namespace DocxportNet.Fields.Resolution.Impl;

public sealed class FrenchMergeMacroProvider : SimpleMergeMacroProvider
{
    protected override string LanguageCode => "fr";
    protected override string Salutation => "Cher";
    protected override string? DualFormSalutation => "Cher·e";
    protected override GreetingStyle Greeting => GreetingStyle.DualFormSalutation;
    protected override bool IncludeTitleInGreeting => false;
    protected override string DefaultGreeting => "Bonjour,";
    protected override LocalityOrder LocalityLayout => LocalityOrder.PostalCityState;

    protected override string? ResolveAddressBlock(IDxpMergeRecordCursor cursor, CultureInfo culture)
    {
        var company = Normalize(GetFieldString(cursor, "Company", culture));
        var address1 = Normalize(GetFieldString(cursor, "Address1", culture));
        var address2 = Normalize(GetFieldString(cursor, "Address2", culture));
        var city = Normalize(GetFieldString(cursor, "City", culture));
        var postalCode = Normalize(GetFieldString(cursor, "PostalCode", culture));
        var cedex = Normalize(GetFieldString(cursor, "Cedex", culture))
            ?? Normalize(GetFieldString(cursor, "CEDEX", culture));
        var country = Normalize(GetFieldString(cursor, "Country", culture));

        var lines = new List<string>();
        if (company is { Length: > 0 })
            lines.Add(company);
        if (address1 is { Length: > 0 })
            lines.Add(address1);
        if (address2 is { Length: > 0 })
            lines.Add(address2);

        var locality = JoinNonEmpty(" ", postalCode, city);
        if (cedex is { Length: > 0 })
            locality = JoinNonEmpty(" ", locality, $"CEDEX {cedex}");
        if (locality is { Length: > 0 })
            lines.Add(locality);

        if (country is { Length: > 0 })
            lines.Add(country);

        if (lines.Count == 0)
            return null;

        return string.Join("\n", lines);
    }
}
