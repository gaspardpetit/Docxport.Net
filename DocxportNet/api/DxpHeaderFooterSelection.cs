namespace DocxportNet.API;

public enum DxpHeaderFooterSelection
{
    None,
    First,
    Last
}

public interface IDxpHeaderFooterSelectionProvider
{
    DxpHeaderFooterSelection HeaderSelection { get; }
    DxpHeaderFooterSelection FooterSelection { get; }
}

