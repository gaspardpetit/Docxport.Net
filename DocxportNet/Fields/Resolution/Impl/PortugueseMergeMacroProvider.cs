using System.Collections.Generic;
using System.Globalization;

namespace DocxportNet.Fields.Resolution.Impl;

public sealed class PortugueseMergeMacroProvider : SimpleMergeMacroProvider
{
    protected override string LanguageCode => "pt";
    protected override string Salutation => "Prezado";
    protected override string? DualFormSalutation => "Prezado(a)";
    protected override GreetingStyle Greeting => GreetingStyle.DualFormSalutation;
    protected override bool IncludeTitleInGreeting => false;
    protected override string DefaultGreeting => "Olá,";
    protected override LocalityOrder LocalityLayout => LocalityOrder.PostalCityState;

    protected override string? ResolveAddressBlock(IDxpMergeRecordCursor cursor, CultureInfo culture)
    {
        var company = Normalize(GetFieldString(cursor, "Company", culture));
        var address1 = Normalize(GetFieldString(cursor, "Address1", culture));
        var address2 = Normalize(GetFieldString(cursor, "Address2", culture));
        var city = Normalize(GetFieldString(cursor, "City", culture));
        var postalCode = Normalize(GetFieldString(cursor, "PostalCode", culture));
        var country = Normalize(GetFieldString(cursor, "Country", culture));

        if (city is { Length: > 0 })
            city = city.ToUpperInvariant();
        if (country is { Length: > 0 })
            country = country.ToUpperInvariant();

        var lines = new List<string>();
        if (company is { Length: > 0 })
            lines.Add(company);
        if (address1 is { Length: > 0 })
            lines.Add(address1);
        if (address2 is { Length: > 0 })
            lines.Add(address2);

        var locality = JoinNonEmpty(" ", postalCode, city);
        if (locality is { Length: > 0 })
            lines.Add(locality);

        if (country is { Length: > 0 })
            lines.Add(country);

        if (lines.Count == 0)
            return null;

        return string.Join("\n", lines);
    }
}
