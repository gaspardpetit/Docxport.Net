namespace DocxportNet.Fields.Resolution.Impl;

public sealed class ChineseSimplifiedMergeMacroProvider : SimpleMergeMacroProvider
{
    protected override string LanguageCode => "zh";
    protected override string Salutation => "尊敬的";
    protected override string GreetingPunctuation => "，";
    protected override string DefaultGreeting => "您好，";
    protected override bool LastNameFirst => true;
    protected override LocalityOrder LocalityLayout => LocalityOrder.PostalCityState;
    protected override string CityStateSeparator => " ";
    protected override string LocalityPostalSeparator => " ";
}
