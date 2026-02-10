using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Walker;

public class DxpSdtContext : DxpISdtContext
{
    public SdtElement Sdt { get; }
    public SdtProperties? Properties { get; }
    public SdtEndCharProperties? EndCharProperties { get; }

    public DxpSdtContext(SdtElement sdt, SdtProperties? properties, SdtEndCharProperties? endCharProperties)
    {
        Sdt = sdt;
        Properties = properties;
        EndCharProperties = endCharProperties;
    }
}
