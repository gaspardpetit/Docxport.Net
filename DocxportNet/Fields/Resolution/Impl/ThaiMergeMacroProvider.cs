namespace DocxportNet.Fields.Resolution.Impl;

public sealed class ThaiMergeMacroProvider : SimpleMergeMacroProvider
{
    protected override string LanguageCode => "th";
    protected override string Salutation => "เรียน";
    protected override string DefaultGreeting => "เรียน ผู้เกี่ยวข้อง,";
    protected override LocalityOrder LocalityLayout => LocalityOrder.PostalCityState;
    protected override string CityStateSeparator => " ";
    protected override string LocalityPostalSeparator => " ";
}
