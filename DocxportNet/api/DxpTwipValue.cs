namespace DocxportNet.API;

public readonly struct DxpTwipValue
{
    public DxpTwipValue(int twips)
    {
        Twips = twips;
    }

    public int Twips { get; }
    public double Inches => Twips / 1440.0;
    public double Millimeters => Twips * 25.4 / 1440.0;
    public double Points => Twips / 20.0;

    public static double ToInches(int twips) => new DxpTwipValue(twips).Inches;
    public static double ToMillimeters(int twips) => new DxpTwipValue(twips).Millimeters;
    public static double ToPoints(int twips) => new DxpTwipValue(twips).Points;
}
