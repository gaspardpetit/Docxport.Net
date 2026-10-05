using System.Text;
using System.Runtime.CompilerServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Doc;
using DocxportNet.Fields;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Visitors.Doc;

/// <summary>Collects body and explicit header/footer stories from a DOCX walk.</summary>
public sealed class DxpDocVisitor : DxpVisitor, IDxpAllHeaderFooterVisitor,
    IDxpCoveredTableCellVisitor
{
    public bool IncludeCoveredTableCells => true;
    private readonly StringBuilder _text = new();
    private readonly List<SectionCapture> _sections = new();
    private readonly Dictionary<StringBuilder, Dictionary<int, (string Font, ushort Character)>>
        _symbols = new(BuilderReferenceComparer.Instance);
    private readonly Dictionary<StringBuilder, List<DocPlainTextFormatRun>> _runs =
        new(BuilderReferenceComparer.Instance);
    private readonly Dictionary<StringBuilder, List<DocPlainTextParagraphStyleRun>>
        _paragraphStyles = new(BuilderReferenceComparer.Instance);
    private readonly List<DocStyleDefinition> _styles = new();
    private readonly Dictionary<string, int> _styleById = new(StringComparer.Ordinal);
    private readonly Dictionary<string, BookmarkCapture> _bookmarks = new(StringComparer.Ordinal);
    private readonly List<PictureCapture> _pictures = new();
    private readonly List<FloatingPictureCapture> _floatingPictures = new();
    private readonly OpenXmlPart?[] _lastHeaderFooterParts = new OpenXmlPart?[6];
    private readonly HashSet<OpenXmlPart> _seenHeaderFooterParts = new();
    private DocListIndex _lists = new([], []);
    private IReadOnlyDictionary<int, short> _listNumberIds = new Dictionary<int, short>();
    private Stream? _output;
    private MainDocumentPart? _mainPart;
    private SectionCapture? _section;
    private StringBuilder? _storyText;
    private StringBuilder? _paragraphText;
    private List<DocTabStop>? _positionalTabs;
    private DocParagraphFormatting? _positionalTabParagraphFormatting;
    private List<short>? _positionalExistingTabPositions;
    private int _positionalLineStart;
    private int? _positionalCellContentWidth;
    private int _currentParagraphStyleIndex;
    private bool _suppressHorizontalContinuation;
    private DocCharacterFormatting _defaultCharacterFormatting = DocCharacterFormatting.Empty;
    private bool _deletedRevisionRun;
    private bool _insertedRevisionRun;
    private string? _deletedRevisionAuthor;
    private string? _insertedRevisionAuthor;
    private DateTime? _deletedRevisionAt;
    private DateTime? _insertedRevisionAt;
    private DocParagraphFormatting _defaultParagraphFormatting = DocParagraphFormatting.Empty;
    private string? _themeMajorLatinFont;
    private string? _themeMinorLatinFont;
    private string? _themeMajorEastAsiaFont;
    private string? _themeMinorEastAsiaFont;
    private string? _themeMajorComplexFont;
    private string? _themeMinorComplexFont;
    private string? _themeEastAsiaLanguage;
    private string? _themeLatinLanguage;
    private string? _themeBidiLanguage;
    private IReadOnlyDictionary<string, string> _themeMajorScriptFonts =
        new Dictionary<string, string>();
    private IReadOnlyDictionary<string, string> _themeMinorScriptFonts =
        new Dictionary<string, string>();
    private IReadOnlyDictionary<string, string> _themeColors =
        new Dictionary<string, string>();
    private IReadOnlyDictionary<string, string> _themeColorMapping =
        new Dictionary<string, string>();
    private bool _evenAndOddHeaders;
    private bool _mirrorMargins;
    private bool _autoHyphenation;
    private bool _hyphenateCaps = true;
    private short _hyphenationZoneTwips;
    private short _consecutiveHyphenLimit;
    private bool _gutterAtTop;
    private bool _balanceSingleByteDoubleByteWidth;
    private bool _growAutofit;
    private IReadOnlyList<DocFontDefinition> _fontMetadata = [];
    private short _defaultTabStopTwips = 720;
    private bool _suppressBookmarkCapture;
    private DocCharacterFormatting? _activeConditionalTableRunFormatting;
    private DocParagraphFormatting? _activeConditionalTableParagraphFormatting;
    private bool _activeCellHasConditionalStyle;
    private TableCapture? _activeTable;
    private RowCapture? _activeRow;
    private readonly Stack<(TableCapture Table, RowCapture? Row)> _tableStack = new();
    private uint _nextTableGroupId;

    private sealed class TableCapture
    {
        public TableCapture(StringBuilder story, IReadOnlyList<short> gridWidths,
            int rowCount, DocTableBorders? borders,
            DocCellShading? backgroundShading, bool autoFit,
            DocTablePreferredWidth? preferredWidth, short? indentTwips,
            byte? justification,
            DocCellMargins? defaultMargins,
            bool rightToLeft, ushort? cellSpacingTwips, Shading? defaultShading,
            Shading? firstRowShading, Shading? lastRowShading,
            TableCellBorders? firstRowBorders, TableCellBorders? lastRowBorders,
            TableCellBorders? firstColumnBorders, TableCellBorders? lastColumnBorders,
            TableCellBorders? northWestBorders, TableCellBorders? northEastBorders,
            TableCellBorders? southWestBorders, TableCellBorders? southEastBorders,
            TableCellBorders? band1HorizontalBorders, TableCellBorders? band2HorizontalBorders,
            TableCellBorders? band1VerticalBorders, TableCellBorders? band2VerticalBorders,
            Shading? band1HorizontalShading, Shading? band2HorizontalShading,
            int horizontalBandSize, int horizontalBandOffset,
            Shading? band1VerticalShading, Shading? band2VerticalShading,
            int verticalBandSize, int verticalBandOffset,
            Shading? firstColumnShading, Shading? lastColumnShading,
            Shading? northWestShading, Shading? northEastShading,
            Shading? southWestShading, Shading? southEastShading,
            ushort? styleIndex, uint groupId)
        {
            Story = story;
            GridWidths = gridWidths;
            RowCount = rowCount;
            Borders = borders;
            BackgroundShading = backgroundShading;
            AutoFit = autoFit;
            PreferredWidth = preferredWidth;
            IndentTwips = indentTwips;
            Justification = justification;
            DefaultMargins = defaultMargins;
            RightToLeft = rightToLeft;
            CellSpacingTwips = cellSpacingTwips;
            DefaultShading = defaultShading;
            FirstRowShading = firstRowShading;
            LastRowShading = lastRowShading;
            FirstRowBorders = firstRowBorders;
            LastRowBorders = lastRowBorders;
            FirstColumnBorders = firstColumnBorders;
            LastColumnBorders = lastColumnBorders;
            NorthWestBorders = northWestBorders;
            NorthEastBorders = northEastBorders;
            SouthWestBorders = southWestBorders;
            SouthEastBorders = southEastBorders;
            Band1HorizontalBorders = band1HorizontalBorders;
            Band2HorizontalBorders = band2HorizontalBorders;
            Band1VerticalBorders = band1VerticalBorders;
            Band2VerticalBorders = band2VerticalBorders;
            Band1HorizontalShading = band1HorizontalShading;
            Band2HorizontalShading = band2HorizontalShading;
            HorizontalBandSize = horizontalBandSize;
            HorizontalBandOffset = horizontalBandOffset;
            Band1VerticalShading = band1VerticalShading;
            Band2VerticalShading = band2VerticalShading;
            VerticalBandSize = verticalBandSize;
            VerticalBandOffset = verticalBandOffset;
            FirstColumnShading = firstColumnShading;
            LastColumnShading = lastColumnShading;
            NorthWestShading = northWestShading;
            NorthEastShading = northEastShading;
            SouthWestShading = southWestShading;
            SouthEastShading = southEastShading;
            StyleIndex = styleIndex;
            GroupId = groupId;
        }
        public StringBuilder Story { get; }
        public IReadOnlyList<short> GridWidths { get; }
        public int RowCount { get; }
        public int RowIndex { get; set; }
        public DocTableBorders? Borders { get; }
        public DocCellShading? BackgroundShading { get; }
        public bool AutoFit { get; }
        public DocTablePreferredWidth? PreferredWidth { get; }
        public short? IndentTwips { get; }
        public byte? Justification { get; }
        public DocCellMargins? DefaultMargins { get; }
        public TableVerticalAlignmentValues? DefaultVerticalAlignment { get; set; }
        public bool? DefaultNoWrap { get; set; }
        public bool? FirstRowNoWrap { get; set; }
        public bool? LastRowNoWrap { get; set; }
        public bool? FirstColumnNoWrap { get; set; }
        public bool? LastColumnNoWrap { get; set; }
        public bool? NorthWestNoWrap { get; set; }
        public bool? NorthEastNoWrap { get; set; }
        public bool? SouthWestNoWrap { get; set; }
        public bool? SouthEastNoWrap { get; set; }
        public bool? Band1HorizontalNoWrap { get; set; }
        public bool? Band2HorizontalNoWrap { get; set; }
        public bool? Band1VerticalNoWrap { get; set; }
        public bool? Band2VerticalNoWrap { get; set; }
        public TableVerticalAlignmentValues? FirstRowVerticalAlignment { get; set; }
        public TableVerticalAlignmentValues? LastRowVerticalAlignment { get; set; }
        public TableVerticalAlignmentValues? FirstColumnVerticalAlignment { get; set; }
        public TableVerticalAlignmentValues? LastColumnVerticalAlignment { get; set; }
        public TableVerticalAlignmentValues? NorthWestVerticalAlignment { get; set; }
        public TableVerticalAlignmentValues? NorthEastVerticalAlignment { get; set; }
        public TableVerticalAlignmentValues? SouthWestVerticalAlignment { get; set; }
        public TableVerticalAlignmentValues? SouthEastVerticalAlignment { get; set; }
        public TableVerticalAlignmentValues? Band1HorizontalVerticalAlignment { get; set; }
        public TableVerticalAlignmentValues? Band2HorizontalVerticalAlignment { get; set; }
        public TableVerticalAlignmentValues? Band1VerticalVerticalAlignment { get; set; }
        public TableVerticalAlignmentValues? Band2VerticalVerticalAlignment { get; set; }
        public bool RightToLeft { get; }
        public ushort? CellSpacingTwips { get; }
        public Shading? DefaultShading { get; }
        public Shading? FirstRowShading { get; }
        public DocCharacterFormatting? FirstRowRunFormatting { get; set; }
        public DocCharacterFormatting? LastRowRunFormatting { get; set; }
        public DocCharacterFormatting? FirstColumnRunFormatting { get; set; }
        public DocCharacterFormatting? LastColumnRunFormatting { get; set; }
        public DocCharacterFormatting? Band1HorizontalRunFormatting { get; set; }
        public DocCharacterFormatting? Band2HorizontalRunFormatting { get; set; }
        public DocCharacterFormatting? Band1VerticalRunFormatting { get; set; }
        public DocCharacterFormatting? Band2VerticalRunFormatting { get; set; }
        public DocCharacterFormatting? NorthWestRunFormatting { get; set; }
        public DocCharacterFormatting? NorthEastRunFormatting { get; set; }
        public DocCharacterFormatting? SouthWestRunFormatting { get; set; }
        public DocCharacterFormatting? SouthEastRunFormatting { get; set; }
        public Dictionary<TableStyleOverrideValues, DocParagraphFormatting>
            ConditionalParagraphFormatting { get; } = new();
        public Shading? LastRowShading { get; }
        public TableCellBorders? FirstRowBorders { get; }
        public TableCellBorders? LastRowBorders { get; }
        public TableCellBorders? FirstColumnBorders { get; }
        public TableCellBorders? LastColumnBorders { get; }
        public TableCellBorders? NorthWestBorders { get; }
        public TableCellBorders? NorthEastBorders { get; }
        public TableCellBorders? SouthWestBorders { get; }
        public TableCellBorders? SouthEastBorders { get; }
        public TableCellBorders? Band1HorizontalBorders { get; }
        public TableCellBorders? Band2HorizontalBorders { get; }
        public TableCellBorders? Band1VerticalBorders { get; }
        public TableCellBorders? Band2VerticalBorders { get; }
        public Shading? Band1HorizontalShading { get; }
        public Shading? Band2HorizontalShading { get; }
        public int HorizontalBandSize { get; }
        public int HorizontalBandOffset { get; }
        public Shading? Band1VerticalShading { get; }
        public Shading? Band2VerticalShading { get; }
        public int VerticalBandSize { get; }
        public int VerticalBandOffset { get; }
        public Shading? FirstColumnShading { get; }
        public Shading? LastColumnShading { get; }
        public Shading? NorthWestShading { get; }
        public Shading? NorthEastShading { get; }
        public Shading? SouthWestShading { get; }
        public Shading? SouthEastShading { get; }
        public ushort? StyleIndex { get; }
        public ushort? LookMask { get; set; }
        public uint GroupId { get; }
        public bool HasFitText { get; set; }
    }
    private sealed class RowCapture
    {
        public DocTableBorders? TableBorders { get; set; }
        public int CellCount { get; set; }
        public List<DocCellShading?> Shadings { get; } = new();
        public List<byte?> VerticalAlignments { get; } = new();
        public List<ushort?> TextFlows { get; } = new();
        public List<bool?> HideMarks { get; } = new();
        public List<byte?> VerticalMerges { get; } = new();
        public List<byte?> HorizontalMerges { get; } = new();
        public List<short> CellWidths { get; } = new();
        public List<DocTablePreferredWidth?> PreferredCellWidths { get; } = new();
        public List<bool?> NoWraps { get; } = new();
        public List<bool?> FitTexts { get; } = new();
        public List<DocCellMargins?> CellMargins { get; } = new();
        public int PendingHorizontalContinuations { get; set; }
        public bool UsesHorizontalMerge { get; set; }
        public bool PreserveHorizontalMergeCells { get; set; }
        public List<DocCellBorders?> Borders { get; } = new();
    }

    private sealed class SectionCapture
    {
        public int EndCp { get; set; }
        public DocSectionFormatting? Formatting { get; set; }
        public StringBuilder?[] Stories { get; } = new StringBuilder?[6];
    }

    private sealed class BookmarkCapture
    {
        public BookmarkCapture(string name, StringBuilder story, int start)
        { Name = name; Story = story; Start = start; }
        public string Name { get; }
        public StringBuilder Story { get; }
        public int Start { get; }
        public int? End { get; set; }
    }

    private sealed record PictureCapture(StringBuilder Story, int Cp,
        DocInlinePicture Picture);
    private sealed record FloatingPictureCapture(StringBuilder Story, int Cp,
        DocInlinePicture Picture, int LeftTwips, int TopTwips,
        byte WrapCode, bool BehindText, byte WrapSide,
        byte HorizontalOrigin, byte VerticalOrigin,
        byte HorizontalAlignment, byte VerticalAlignment,
        int DistanceTopEmu, int DistanceBottomEmu,
        int DistanceLeftEmu, int DistanceRightEmu);

    private sealed class BuilderReferenceComparer : IEqualityComparer<StringBuilder>
    {
        public static BuilderReferenceComparer Instance { get; } = new();
        public bool Equals(StringBuilder? x, StringBuilder? y) => ReferenceEquals(x, y);
        public int GetHashCode(StringBuilder value) => RuntimeHelpers.GetHashCode(value);
    }

    public DxpFieldEval FieldEval { get; }

    public DxpDocVisitor(ILogger? logger = null, DxpFieldEval? fieldEval = null) : base(logger)
        => FieldEval = fieldEval ?? new DxpFieldEval(logger: logger);

    public override void SetOutput(Stream stream) =>
        _output = stream ?? throw new ArgumentNullException(nameof(stream));

    public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext context)
    {
        if (_output == null)
            throw new InvalidOperationException("An output stream must be assigned before walking the document.");
        _text.Clear();
        _sections.Clear();
        _runs.Clear();
        _symbols.Clear();
        _paragraphStyles.Clear();
        _styles.Clear();
        _styleById.Clear();
        _bookmarks.Clear();
        _pictures.Clear();
        _floatingPictures.Clear();
        Array.Clear(_lastHeaderFooterParts, 0, _lastHeaderFooterParts.Length);
        _seenHeaderFooterParts.Clear();
        _section = null;
        _storyText = null;
        _paragraphText = null;
        _currentParagraphStyleIndex = 0;
        _suppressHorizontalContinuation = false;
        _activeTable = null;
        _tableStack.Clear();
        _nextTableGroupId = 0;
        _activeRow = null;
        _suppressBookmarkCapture = false;
        _mainPart = doc.MainDocumentPart;
        _fontMetadata = DocFontTable.ReadOpenXml(_mainPart);
        (_themeMajorLatinFont, _themeMinorLatinFont,
            _themeMajorEastAsiaFont, _themeMinorEastAsiaFont,
            _themeMajorComplexFont, _themeMinorComplexFont,
            _themeMajorScriptFonts, _themeMinorScriptFonts) =
            ReadThemeFonts(_mainPart?.ThemePart);
        _themeEastAsiaLanguage = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<ThemeFontLanguages>()?.EastAsia?.Value;
        _themeLatinLanguage = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<ThemeFontLanguages>()?.Val?.Value;
        _themeBidiLanguage = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<ThemeFontLanguages>()?.Bidi?.Value;
        _themeColors = ReadThemeColors(_mainPart?.ThemePart);
        _themeColorMapping = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<ColorSchemeMapping>()?.GetAttributes()
            .ToDictionary(x => x.LocalName, x => x.Value,
                StringComparer.OrdinalIgnoreCase) ??
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        (_lists, _listNumberIds) = DocOpenXmlListReader.Read(_mainPart,
            properties => ReadCharacterFormatting(properties),
            properties => ReadParagraphFormatting(properties));
        var evenOddSetting = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<EvenAndOddHeaders>();
        _evenAndOddHeaders = evenOddSetting != null &&
            (evenOddSetting.Val?.Value ?? true);
        var mirrorSetting = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<MirrorMargins>();
        _mirrorMargins = mirrorSetting != null &&
            (mirrorSetting.Val?.Value ?? true);
        var autoHyphenationSetting = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<AutoHyphenation>();
        _autoHyphenation = autoHyphenationSetting != null &&
            (autoHyphenationSetting.Val?.Value ?? true);
        var doNotHyphenateCaps = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<DoNotHyphenateCaps>();
        _hyphenateCaps = doNotHyphenateCaps == null ||
            !(doNotHyphenateCaps.Val?.Value ?? true);
        var hyphenationZone = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<HyphenationZone>()?.Val?.Value;
        _hyphenationZoneTwips = short.TryParse(hyphenationZone,
            System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var zoneTwips) &&
            zoneTwips >= 0 ? zoneTwips : (short)0;
        var consecutiveHyphenLimit = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<ConsecutiveHyphenLimit>()?.Val?.Value;
        _consecutiveHyphenLimit = consecutiveHyphenLimit is <= (ushort)short.MaxValue
            ? checked((short)consecutiveHyphenLimit.Value) : (short)0;
        var topGutterSetting = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<GutterAtTop>();
        _gutterAtTop = topGutterSetting != null &&
            (topGutterSetting.Val?.Value ?? true);
        var balanceSetting = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<Compatibility>()?
            .GetFirstChild<BalanceSingleByteDoubleByteWidth>();
        _balanceSingleByteDoubleByteWidth = balanceSetting != null &&
            (balanceSetting.Val?.Value ?? true);
        var growAutofitSetting = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<Compatibility>()?.GetFirstChild<GrowAutofit>();
        _growAutofit = growAutofitSetting != null &&
            (growAutofitSetting.Val?.Value ?? true);
        var defaultTabStop = _mainPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<DefaultTabStop>()?.Val?.Value;
        _defaultTabStopTwips = defaultTabStop is > 0 and <= short.MaxValue
            ? checked((short)defaultTabStop.Value) : (short)720;
        var documentDefaults = _mainPart?.StyleDefinitionsPart?.Styles?
            .Elements<DocDefaults>().FirstOrDefault();
        var defaultRuns = documentDefaults?.GetFirstChild<RunPropertiesDefault>()?
            .RunPropertiesBaseStyle;
        var hasDefinedRunDefaults = defaultRuns?.ChildElements.Count > 0;
        var hasNormalStyle = _mainPart?.StyleDefinitionsPart?.Styles?
            .Elements<Style>().Any(x => x.Type?.Value == StyleValues.Paragraph &&
                x.StyleId?.Value == "Normal") == true;
        _defaultParagraphFormatting = ReadParagraphFormatting(documentDefaults?
            .GetFirstChild<ParagraphPropertiesDefault>()?.ParagraphPropertiesBaseStyle);
        _defaultCharacterFormatting = ReadCharacterFormatting(defaultRuns);
        // Word's implicit kerning threshold depends on whether an authored
        // Normal style is present. A document with explicit run defaults but
        // no Normal style keeps the zero-point default when saved as DOC.
        _defaultCharacterFormatting = _defaultCharacterFormatting with
        {
            KerningThresholdHalfPoints =
                _defaultCharacterFormatting.KerningThresholdHalfPoints ??
                (hasDefinedRunDefaults && !hasNormalStyle ? (ushort)0 : (ushort)2)
        };
        if (!hasDefinedRunDefaults)
            // Word's current implicit Normal font is Aptos when the DOCX
            // supplies neither run defaults nor a theme font definition.
            _defaultCharacterFormatting = _defaultCharacterFormatting with
            {
                AsciiFontName = _defaultCharacterFormatting.AsciiFontName ??
                    _themeMinorLatinFont ?? "Aptos",
                HighAnsiFontName = _defaultCharacterFormatting.HighAnsiFontName ??
                    _themeMinorLatinFont ?? "Aptos",
                EastAsiaFontName = _defaultCharacterFormatting.EastAsiaFontName ??
                    _themeMinorEastAsiaFont,
                ComplexScriptFontName = _defaultCharacterFormatting.ComplexScriptFontName ??
                    _themeMinorComplexFont ?? "Cordia New",
                ComplexScriptSizeHalfPoints = hasNormalStyle
                    ? _defaultCharacterFormatting.ComplexScriptSizeHalfPoints
                    : _defaultCharacterFormatting.ComplexScriptSizeHalfPoints ?? 30
            };
        var sourceStyles = _mainPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()
            .Where(x => x.Type?.Value == StyleValues.Paragraph ||
                x.Type?.Value == StyleValues.Character ||
                x.Type?.Value == StyleValues.Table)
            .ToArray() ?? [];
        if (!hasDefinedRunDefaults && _defaultCharacterFormatting.EastAsiaFontName == null)
        {
            // Word promotes a sole authored East Asian paragraph font to the
            // implicit DOC default. Without it, mixed-script DOC text uses the
            // machine's Normal font even when its paragraph style has the font.
            var eastAsiaFonts = sourceStyles.Where(x => x.Type?.Value ==
                StyleValues.Paragraph).Select(x => x.StyleRunProperties?
                .GetFirstChild<RunFonts>()?.EastAsia?.Value)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase).Take(2).ToArray();
            if (eastAsiaFonts.Length == 1)
            {
                var selected = eastAsiaFonts[0];
                // Word's implicit East Asian Normal font remains the fallback
                // when the sole authored font cannot draw document CJK text.
                // It may itself need glyph substitution (for example Hangul),
                // but promoting the authored font changes nearby space widths.
                var authoredFace = DocSystemFontAdvances.Find(selected);
                var fallbackFace = DocSystemFontAdvances.Find("Yu Mincho");
                if (authoredFace != null && fallbackFace != null)
                {
                    var storyText = (_mainPart?.Document?.Descendants<Text>()
                        .Select(x => x.Text) ?? Enumerable.Empty<string>())
                        .Concat(_mainPart?.HeaderParts.SelectMany(x =>
                            x.Header?.Descendants<Text>().Select(t => t.Text) ??
                            Enumerable.Empty<string>()) ?? Enumerable.Empty<string>())
                        .Concat(_mainPart?.FooterParts.SelectMany(x =>
                            x.Footer?.Descendants<Text>().Select(t => t.Text) ??
                            Enumerable.Empty<string>()) ?? Enumerable.Empty<string>());
                    bool HasMissingEastAsianGlyph(string content)
                    {
                        for (var i = 0; i < content.Length; i++)
                        {
                            var start = i;
                            int scalar = content[i];
                            if (char.IsHighSurrogate(content[i]) &&
                                i + 1 < content.Length &&
                                char.IsLowSurrogate(content[i + 1]))
                                scalar = char.ConvertToUtf32(content[i], content[++i]);
                            if ((scalar is >= 0x3000 and <= 0x9FFF or
                                >= 0xAC00 and <= 0xD7AF or
                                >= 0x20000 and <= 0x323AF) &&
                                !authoredFace.TryMeasure(content.Substring(start,
                                    i - start + 1), 12, out _))
                                return true;
                        }
                        return false;
                    }
                    if (storyText.Any(HasMissingEastAsianGlyph))
                        selected = "Yu Mincho";
                }
                _defaultCharacterFormatting = _defaultCharacterFormatting with
                { EastAsiaFontName = selected };
            }
        }
        var normalWithExplicitSize = sourceStyles.Any(x =>
            x.Type?.Value == StyleValues.Paragraph && x.StyleId?.Value == "Normal" &&
            x.StyleRunProperties?.GetFirstChild<FontSize>() != null);
        var useBuiltInNormalSlot = normalWithExplicitSize &&
            (!hasDefinedRunDefaults ||
                _defaultCharacterFormatting.SizeHalfPoints != null);
        if (!hasDefinedRunDefaults && normalWithExplicitSize)
        {
            var normalSize = sourceStyles.First(x => x.Type?.Value ==
                StyleValues.Paragraph && x.StyleId?.Value == "Normal")
                .StyleRunProperties!.GetFirstChild<FontSize>()!.Val?.Value;
            if (ushort.TryParse(normalSize, out var halfPoints))
                _defaultCharacterFormatting = _defaultCharacterFormatting with
                { SizeHalfPoints = halfPoints };
        }
        var normalWithParagraphOverrides = sourceStyles.Any(x =>
            x.Type?.Value == StyleValues.Paragraph && x.StyleId?.Value == "Normal" &&
            x.StyleParagraphProperties?.ChildElements.Count > 0);
        var needsImplicitParagraphDefaults = _defaultParagraphFormatting.IsEmpty &&
            !normalWithParagraphOverrides;
        if (!hasDefinedRunDefaults && !normalWithExplicitSize)
        {
            // Word uses 12-point Normal text when docDefaults is absent.
            _defaultCharacterFormatting = _defaultCharacterFormatting with
            { SizeHalfPoints = _defaultCharacterFormatting.SizeHalfPoints ?? 24 };
        }
        if (needsImplicitParagraphDefaults && !normalWithExplicitSize)
        {
            // An explicit run default does not suppress Word's implicit
            // Normal paragraph spacing. A Normal paragraph override does.
            // Word lays out the DOC form of the implicit 1.15-line default
            // slightly shorter with 276; 278 retains the DOCX line height.
            _defaultParagraphFormatting = _defaultParagraphFormatting with
            {
                AfterTwips = _defaultParagraphFormatting.AfterTwips ?? 160,
                LineValue = _defaultParagraphFormatting.LineValue ?? 278,
                LineIsMultiple = _defaultParagraphFormatting.LineIsMultiple ?? true
            };
        }
        for (var i = 0; i < sourceStyles.Length; i++)
            if (sourceStyles[i].StyleId?.Value is string id)
                _styleById[id] = useBuiltInNormalSlot && id == "Normal" ? 0 : 15 + i;
        for (var i = 0; i < sourceStyles.Length; i++)
        {
            var source = sourceStyles[i];
            var id = source.StyleId?.Value;
            if (id == null) continue;
            var basedOnId = source.BasedOn?.Val?.Value;
            int? basedOn = null;
            if (basedOnId != null && _styleById.TryGetValue(basedOnId, out var parent))
                basedOn = parent;
            else if (source.Type?.Value == StyleValues.Paragraph && id != "Normal")
            {
                if (!hasDefinedRunDefaults &&
                    _styleById.TryGetValue("Normal", out var normalIndex))
                    basedOn = normalIndex;
                else if (!hasDefinedRunDefaults || !hasNormalStyle)
                    basedOn = 0;
            }
            var nextId = source.NextParagraphStyle?.Val?.Value;
            var next = nextId != null && _styleById.TryGetValue(nextId, out var foundNext)
                ? foundNext : _styleById[id];
            var linkedId = source.LinkedStyle?.Val?.Value;
            var linked = linkedId != null && _styleById.TryGetValue(linkedId, out var linkedIndex)
                ? (int?)linkedIndex : null;
            var characterFormatting = ReadCharacterFormatting(source.StyleRunProperties);
            if (id == "Normal")
            {
                characterFormatting = MergeCharacterDefaults(characterFormatting,
                    _defaultCharacterFormatting);
                if (!hasDefinedRunDefaults)
                    characterFormatting = characterFormatting with
                    {
                        ComplexScriptSizeHalfPoints =
                            characterFormatting.ComplexScriptSizeHalfPoints ?? 30
                    };
            }
            var paragraphFormatting = ReadParagraphFormatting(source.StyleParagraphProperties);
            if (id == "Normal")
                paragraphFormatting = MergeParagraphDefaults(paragraphFormatting,
                    _defaultParagraphFormatting);
            else if (source.Type?.Value == StyleValues.Paragraph && basedOn == null)
            {
                // An unbased DOCX paragraph style still inherits docDefaults.
                // DOC has no separate style-default layer for that style.
                characterFormatting = MergeCharacterDefaults(characterFormatting,
                    _defaultCharacterFormatting);
                paragraphFormatting = MergeParagraphDefaults(paragraphFormatting,
                    _defaultParagraphFormatting);
            }
            if (id == "Normal" && needsImplicitParagraphDefaults && normalWithExplicitSize)
                paragraphFormatting = MergeParagraphDefaults(paragraphFormatting,
                    DocParagraphFormatting.Empty with
                    {
                        AfterTwips = 160,
                        LineValue = 278,
                        LineIsMultiple = true
                    });
            var tableShading = source.StyleTableCellProperties?.GetFirstChild<Shading>();
            DocParagraphFormatting? tableFormatting = null;
            if (source.Type?.Value == StyleValues.Table)
            {
                var styleBorders = ReadTableBorders(source.StyleTableProperties?
                    .GetFirstChild<TableBorders>());
                DocCellShading? styleShading = null;
                // ShdNil in a table style is ignored by Word. It does not
                // clear inherited cell shading like a direct cell ShdNil.
                if (tableShading != null &&
                    tableShading.Val?.Value != ShadingPatternValues.Nil)
                {
                    var pattern = DocShadingPatterns.ToDoc(tableShading.Val?.Value ??
                        ShadingPatternValues.Clear);
                    if (pattern == null)
                        throw new NotSupportedException("The DOC table-style shading pattern is unsupported.");
                    styleShading = new DocCellShading(ReadShadingFill(tableShading),
                        ReadShadingForeground(tableShading), pattern.Value);
                }
                DocCellShading? backgroundShading = null;
                if (source.StyleTableProperties?.GetFirstChild<Shading>() is { } background &&
                    background.Val?.Value != ShadingPatternValues.Nil)
                {
                    var pattern = DocShadingPatterns.ToDoc(background.Val?.Value ??
                        ShadingPatternValues.Clear);
                    if (pattern == null)
                        throw new NotSupportedException("The DOC table-background shading pattern is unsupported.");
                    backgroundShading = new DocCellShading(ReadShadingFill(background),
                        ReadShadingForeground(background), pattern.Value);
                }
                tableFormatting = DocParagraphFormatting.Empty with
                {
                    TableStyleShading = styleShading,
                    TableBackgroundShading = backgroundShading,
                    TableStyleNoWrap = source.StyleTableCellProperties?
                        .GetFirstChild<NoWrap>() is { } baseNoWrap
                            ? baseNoWrap.Val == null ||
                                baseNoWrap.Val.Value == OnOffOnlyValues.On
                            : null,
                    TableStyleVerticalAlignment = source.StyleTableCellProperties?
                        .GetFirstChild<TableCellVerticalAlignment>()?.Val?.Value
                        is { } baseVerticalAlignment
                            ? baseVerticalAlignment ==
                                TableVerticalAlignmentValues.Center ? (byte)1 :
                                baseVerticalAlignment == TableVerticalAlignmentValues.Bottom
                                    ? (byte)2 : (byte)0
                            : null,
                    TableBorders = styleBorders,
                    TableCellSpacingTwips = ReadCellSpacing(source.StyleTableProperties?
                        .GetFirstChild<TableCellSpacing>()),
                    TableDefaultCellMargins = ReadCellMargins(source.StyleTableProperties?
                        .GetFirstChild<TableCellMarginDefault>()),
                    TableIndentTwips = ReadTableIndent(source.StyleTableProperties?
                        .GetFirstChild<TableIndentation>()),
                    TableJustification = ReadTableJustification(source.StyleTableProperties?
                        .GetFirstChild<TableJustification>()),
                    TableHorizontalBandSize = source.StyleTableProperties?
                        .GetFirstChild<TableStyleRowBandSize>()?.Val?.Value is int rowSize
                            and >= 1 and <= 3 ? checked((byte)rowSize) : null,
                    TableVerticalBandSize = source.StyleTableProperties?
                        .GetFirstChild<TableStyleColumnBandSize>()?.Val?.Value is int columnSize
                            and >= 1 and <= 3 ? checked((byte)columnSize) : null
                };
            }
            var styleName = source.StyleName?.Val?.Value ?? id;
            if (source.Aliases?.Val?.Value is string aliases && aliases.Length != 0)
                styleName += "," + aliases;
            Dictionary<ushort, DocParagraphFormatting>? conditionalParagraphs = null;
            Dictionary<ushort, DocCharacterFormatting>? conditionalCharacters = null;
            Dictionary<ushort, DocCellShading>? conditionalShadings = null;
            Dictionary<ushort, DocConditionalTableBorders>? conditionalBorders = null;
            Dictionary<ushort, byte>? conditionalVerticalAlignments = null;
            Dictionary<ushort, bool>? conditionalNoWraps = null;
            if (source.Type?.Value == StyleValues.Table)
                foreach (var condition in source.Elements<TableStyleProperties>())
                {
                    var code = condition.Type?.Value is { } kind
                        ? DocTableStyleCondition.FromOpenXml(kind) : (ushort)0;
                    if (code == 0) continue;
                    if (condition.GetFirstChild<StyleParagraphProperties>() is { } paragraph)
                    {
                        conditionalParagraphs ??= new();
                        conditionalParagraphs[code] = ReadParagraphFormatting(paragraph);
                    }
                    if (condition.RunPropertiesBaseStyle is { } run)
                    {
                        conditionalCharacters ??= new();
                        conditionalCharacters[code] = ReadCharacterFormatting(run);
                    }
                    if (condition.TableStyleConditionalFormattingTableCellProperties?
                        .GetFirstChild<Shading>() is { } shading &&
                        shading.Val?.Value != ShadingPatternValues.Nil)
                    {
                        var pattern = DocShadingPatterns.ToDoc(shading.Val?.Value ??
                            ShadingPatternValues.Clear);
                        if (pattern == null)
                            throw new NotSupportedException("The conditional table-style shading pattern is unsupported.");
                        conditionalShadings ??= new();
                        conditionalShadings[code] = new DocCellShading(
                            ReadShadingFill(shading), ReadShadingForeground(shading),
                            pattern.Value);
                    }
                    if (DocConditionalTableBorders.FromOpenXml(condition
                        .TableStyleConditionalFormattingTableCellProperties?
                        .GetFirstChild<TableCellBorders>()) is { } borders)
                    {
                        conditionalBorders ??= new();
                        conditionalBorders[code] = borders;
                    }
                    if (condition.TableStyleConditionalFormattingTableCellProperties?
                        .GetFirstChild<TableCellVerticalAlignment>()?.Val?.Value
                        is { } verticalAlignment)
                    {
                        conditionalVerticalAlignments ??= new();
                        conditionalVerticalAlignments[code] = verticalAlignment ==
                            TableVerticalAlignmentValues.Center ? (byte)1 :
                            verticalAlignment == TableVerticalAlignmentValues.Bottom
                                ? (byte)2 : (byte)0;
                    }
                    if (condition.TableStyleConditionalFormattingTableCellProperties?
                        .GetFirstChild<NoWrap>() is { } conditionalNoWrap)
                    {
                        conditionalNoWraps ??= new();
                        conditionalNoWraps[code] = conditionalNoWrap.Val == null ||
                            conditionalNoWrap.Val.Value == OnOffOnlyValues.On;
                    }
                }
            _styles.Add(new DocStyleDefinition(_styleById[id],
                styleName,
                source.Type?.Value == StyleValues.Paragraph ? 1 :
                    source.Type?.Value == StyleValues.Character ? 2 : 3,
                basedOn, next,
                characterFormatting, paragraphFormatting, tableFormatting,
                source.CustomStyle?.Value == true ? 0x0FFE : id switch
                {
                    "Normal" => 0,
                    "Heading1" => 1,
                    "Heading2" => 2,
                    "Heading3" => 3,
                    "Heading4" => 4,
                    "Heading5" => 5,
                    "Heading6" => 6,
                    "Heading7" => 7,
                    "Heading8" => 8,
                    "Heading9" => 9,
                    "Title" => 62,
                    "Subtitle" => 74,
                    "Hyperlink" => 85,
                    "Header" => 31,
                    "Footer" => 32,
                    "DefaultParagraphFont" => 65,
                    "TableNormal" => 105,
                    "NoList" => 107,
                    "ListParagraph" => 179,
                    "Quote" => 180,
                    "IntenseQuote" => 181,
                    "IntenseEmphasis" => 261,
                    "IntenseReference" => 263,
                    "UnresolvedMention" => 374,
                    _ => source.StyleName?.Val?.Value?.ToLowerInvariant() switch
                    {
                        "normal" => 0,
                        "heading 1" => 1,
                        "heading 2" => 2,
                        "heading 3" => 3,
                        "heading 4" => 4,
                        "heading 5" => 5,
                        "heading 6" => 6,
                        "heading 7" => 7,
                        "heading 8" => 8,
                        "heading 9" => 9,
                        "title" => 62,
                        "subtitle" => 74,
                        "hyperlink" => 85,
                        "header" => 31,
                        "footer" => 32,
                        "default paragraph font" => 65,
                        "table normal" => 105,
                        "no list" => 107,
                        "list paragraph" => 179,
                        "quote" => 180,
                        "intense quote" => 181,
                        "intense emphasis" => 261,
                        "intense reference" => 263,
                        "unresolved mention" => 374,
                        _ => null
                    }
                }, LinkedStyleIndex: linked,
                ConditionalParagraphFormatting: conditionalParagraphs,
                ConditionalCharacterFormatting: conditionalCharacters,
                ConditionalTableShading: conditionalShadings,
                ConditionalTableBorders: conditionalBorders,
                ConditionalTableVerticalAlignment: conditionalVerticalAlignments,
                ConditionalTableNoWrap: conditionalNoWraps));
        }
        if (useBuiltInNormalSlot && hasDefinedRunDefaults)
            // Word needs the authored Normal at index 0 for character-unit
            // indents; retain separate DOCX run defaults in an unused style.
            _styles.Add(new DocStyleDefinition(15 + sourceStyles.Length,
                DocDefaultStructures.PreservedRunDefaultsStyleName, 2, null, null,
                _defaultCharacterFormatting));
        return DxpDisposable.Create(() =>
        {
            if (_text.Length == 0 || _text[_text.Length - 1] != '\r') _text.Append('\r');
            if (_sections.Count == 0) _sections.Add(new SectionCapture());
            _sections[_sections.Count - 1].EndCp = _text.Length;
            for (var i = 0; i < _sections.Count - 1; i++)
            {
                var end = _sections[i].EndCp;
                if (end == 0 || _text[end - 1] != '\r')
                    throw new InvalidDataException("A DOCX section does not end at a paragraph mark.");
                _text[end - 1] = '\f';
            }
            IReadOnlyList<DocPlainTextFormatRun> RunsFor(StringBuilder? story) =>
                story != null && _runs.TryGetValue(story, out var runs) ? runs : [];
            IReadOnlyList<DocPlainTextParagraphStyleRun> ParagraphStylesFor(StringBuilder? story) =>
                story != null && _paragraphStyles.TryGetValue(story, out var runs) ? runs : [];
            var capturedStories = new HashSet<StringBuilder>(BuilderReferenceComparer.Instance)
                { _text };
            foreach (var section in _sections)
            foreach (var story in section.Stories)
                if (story != null) capturedStories.Add(story);
            if (_bookmarks.Values.Any(x => !capturedStories.Contains(x.Story)) ||
                _pictures.Any(x => !capturedStories.Contains(x.Story)) ||
                _floatingPictures.Any(x => !capturedStories.Contains(x.Story)))
                throw new InvalidDataException(
                    "A DOCX bookmark or drawing has no captured document story.");
            DocPlainTextStory CaptureStory(StringBuilder story)
            {
                var storyText = story.ToString();
                var styles = ParagraphStylesFor(story);
                return new DocPlainTextStory(storyText, RunsFor(story)
                    .Select(x => new DocStoryCharacterRun(x.Start, x.End,
                        x.Formatting)).ToArray(),
                    styles.Where(x => !DocStoryTableRow.IsRow(x.Formatting))
                        .Select(x => new DocStoryParagraphStyleRun(x.Start, x.End,
                            x.StyleIndex, x.Formatting)).ToArray())
                {
                    Paragraphs = DocStoryParagraphRange.Capture(storyText, styles,
                        ReferenceEquals(story, _text) ? _sections.Take(
                            Math.Max(0, _sections.Count - 1)).Select(x => x.EndCp)
                            .ToArray() : null),
                    TableRows = styles.Where(x => DocStoryTableRow.IsRow(x.Formatting))
                        .Select(x => new DocStoryTableRow(x.Start, x.End, x.StyleIndex,
                            x.Formatting!)).ToArray(),
                    ListParagraphs = DocStoryListParagraph.Capture(
                        styles.Where(x => !DocStoryTableRow.IsRow(x.Formatting))
                            .Select(x => new DocStoryParagraphStyleRun(x.Start, x.End,
                                x.StyleIndex, x.Formatting)),
                        styles.Where(x => DocStoryTableRow.IsRow(x.Formatting))
                            .Select(x => new DocStoryTableRow(x.Start, x.End,
                                x.StyleIndex, x.Formatting!))),
                    FieldMarks = DocPlainTextWriter.ReadFieldMarks(storyText),
                    Bookmarks = _bookmarks.Values.Where(x => ReferenceEquals(x.Story, story))
                        .Select(x =>
                        {
                            var end = x.End ?? throw new InvalidDataException(
                                $"DOCX bookmark '{x.Name}' has no end.");
                            if (x.Start < 0 || end < x.Start || end > storyText.Length)
                                throw new InvalidDataException(
                                    $"DOCX bookmark '{x.Name}' exceeds its story.");
                            return new DocStoryBookmark(x.Name, x.Start, end);
                        }).ToArray(),
                    Pictures = _pictures.Where(x => ReferenceEquals(x.Story, story))
                        .Select(x =>
                        {
                            if (x.Cp < 0 || x.Cp >= storyText.Length ||
                                storyText[x.Cp] != '\u0001')
                                throw new InvalidDataException(
                                    "A DOCX inline picture exceeds its story.");
                            return new DocStoryInlinePicture(x.Cp)
                                { Payload = x.Picture };
                        }).ToArray(),
                    FloatingPictures = _floatingPictures.Where(x => ReferenceEquals(x.Story, story))
                        .Select(x =>
                        {
                            if (x.Cp < 0 || x.Cp >= storyText.Length ||
                                storyText[x.Cp] != '\u0008')
                                throw new InvalidDataException(
                                    "A DOCX floating picture exceeds its story.");
                            return new DocStoryFloatingPicture(x.Cp, x.Picture.Bytes,
                                x.Picture.ContentType, x.Picture.WidthEmu,
                                x.Picture.HeightEmu, x.LeftTwips, x.TopTwips,
                                x.WrapCode, x.BehindText, x.WrapSide,
                                x.HorizontalOrigin, x.VerticalOrigin,
                                x.HorizontalAlignment, x.VerticalAlignment,
                                x.DistanceTopEmu, x.DistanceBottomEmu,
                                x.DistanceLeftEmu, x.DistanceRightEmu)
                            {
                                Crop = x.Picture.Crop,
                                FlipHorizontal = x.Picture.FlipHorizontal,
                                FlipVertical = x.Picture.FlipVertical,
                                RotationDegrees = x.Picture.RotationDegrees
                            };
                        }).ToArray()
                };
            }
            var sectionStart = 0;
            var sections = _sections.Select(x =>
            {
                var model = new DocStorySection(sectionStart, x.EndCp,
                    x.Formatting ?? new DocSectionFormatting(),
                    Enumerable.Range(0, 6).Select(slot => new DocStorySectionSlot(
                        slot, x.Stories[slot] != null)).ToArray());
                sectionStart = x.EndCp;
                return new DocPlainTextSection(x.EndCp,
                    x.Stories.Select(y => y == null ? null : CaptureStory(y)).ToArray(),
                    x.Formatting) { Model = model };
            }).ToArray();
            DocPlainTextWriter.Write(_output, new DocPlainTextDocument(
                CaptureStory(_text), sections,
                Styles: _styles,
                DefaultCharacterFormatting: _defaultCharacterFormatting,
                EvenAndOddHeaders: _evenAndOddHeaders,
                Lists: _lists,
                DefaultParagraphFormatting: _defaultParagraphFormatting,
                DefaultTabStopTwips: _defaultTabStopTwips,
                MirrorMargins: _mirrorMargins,
                GutterAtTop: _gutterAtTop,
                FontMetadata: _fontMetadata,
                BalanceSingleByteDoubleByteWidth: _balanceSingleByteDoubleByteWidth,
                AutoHyphenation: _autoHyphenation,
                HyphenationZoneTwips: _hyphenationZoneTwips,
                ConsecutiveHyphenLimit: _consecutiveHyphenLimit,
                HyphenateCaps: _hyphenateCaps,
                Title: context.CoreProperties?.Title,
                Subject: context.CoreProperties?.Subject,
                Author: context.CoreProperties?.Creator,
                Keywords: context.CoreProperties?.Keywords,
                Comments: context.CoreProperties?.Description,
                LastAuthor: context.CoreProperties?.LastModifiedBy,
                PageCount: ReadStatistic(context.ExtendedProperties?.Pages?.Text),
                WordCount: ReadStatistic(context.ExtendedProperties?.Words?.Text),
                CharacterCount: ReadStatistic(context.ExtendedProperties?.Characters?.Text),
                GrowAutofit: _growAutofit,
                RevisionNumber: context.CoreProperties?.Revision));
        });
    }

    public override IDisposable VisitSectionBegin(SectionProperties properties, SectionLayout layout,
        DxpIDocumentContext context)
    {
        var previous = _section;
        _section = new SectionCapture();
        _section.Formatting = DocSectionFormatting.FromOpenXml(properties);
        _sections.Add(_section);
        return DxpDisposable.Create(() =>
        {
            _section.EndCp = _text.Length;
            _section = previous;
        });
    }

    public override IDisposable VisitSectionHeaderBegin(Header header, object value,
        DxpIDocumentContext context) => BeginHeaderFooter(value, footer: false);

    public override IDisposable VisitSectionFooterBegin(Footer footer, object value,
        DxpIDocumentContext context) => BeginHeaderFooter(value, footer: true);

    private IDisposable BeginHeaderFooter(object value, bool footer)
    {
        if (_section == null || value is not DxpHeaderFooterContext story)
            return DxpDisposable.Empty;
        var slot = story.Kind == HeaderFooterValues.Even ? (footer ? 2 : 0) :
            story.Kind == HeaderFooterValues.First ? (footer ? 5 : 4) :
            footer ? 3 : 1;
        var previous = _storyText;
        var previousSuppression = _suppressBookmarkCapture;
        var reusedPart = story.Part != null && !_seenHeaderFooterParts.Add(story.Part);
        if (story.Part != null && ReferenceEquals(_lastHeaderFooterParts[slot], story.Part))
        {
            _storyText = new StringBuilder();
            _suppressBookmarkCapture = true;
            return DxpDisposable.Create(() =>
            {
                _storyText = previous;
                _suppressBookmarkCapture = previousSuppression;
            });
        }
        _lastHeaderFooterParts[slot] = story.Part;
        _storyText = _section.Stories[slot] = new StringBuilder();
        _suppressBookmarkCapture = reusedPart;
        return DxpDisposable.Create(() =>
        {
            _storyText = previous;
            _suppressBookmarkCapture = previousSuppression;
        });
    }

    public override IDisposable VisitParagraphBegin(Paragraph paragraph, DxpIDocumentContext context,
        DxpIParagraphContext paragraphContext)
    {
        if (_suppressHorizontalContinuation) return DxpDisposable.Empty;
        var target = _storyText ?? (context.CurrentPart == _mainPart ? _text : null);
        if (target == null) return DxpDisposable.Empty;
        var previous = _paragraphText;
        var previousPositionalTabs = _positionalTabs;
        var previousPositionalFormatting = _positionalTabParagraphFormatting;
        var previousTabPositions = _positionalExistingTabPositions;
        var previousLineStart = _positionalLineStart;
        _positionalLineStart = target.Length;
        _positionalTabs = new List<DocTabStop>();
        _positionalExistingTabPositions = new List<short>();
        var previousStyleIndex = _currentParagraphStyleIndex;
        var start = target.Length;
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        _currentParagraphStyleIndex = styleId != null &&
            _styleById.TryGetValue(styleId, out var paragraphStyleIndex)
            ? paragraphStyleIndex : _styleById.TryGetValue("Normal", out var normalIndex)
                ? normalIndex : 0;
        var formatting = ReadParagraphFormatting(paragraph.ParagraphProperties);
        if (_activeConditionalTableParagraphFormatting is { } conditionalFormatting)
        {
            var styleFormatting = DocParagraphFormatting.Empty;
            var visited = new HashSet<int>();
            var index = _currentParagraphStyleIndex;
            while (index != 0 && visited.Add(index))
            {
                var style = _styles.FirstOrDefault(x => x.Index == index);
                if (style == null) break;
                styleFormatting = MergeParagraphDefaults(styleFormatting,
                    style.ParagraphFormatting);
                index = style.BasedOnIndex ?? 0;
            }
            // Without cell cnfStyle or style alignment, Word displays the
            // conditional table alignment for an explicit direct left. Its
            // saved cnfStyle or an aligned paragraph style makes left win.
            if (!_activeCellHasConditionalStyle &&
                styleFormatting.Justification == null &&
                conditionalFormatting.Justification != null &&
                paragraph.ParagraphProperties?.Justification?.Val?.Value ==
                    JustificationValues.Left)
                formatting = formatting with { Justification = null };
            conditionalFormatting = ExcludeParagraphStyleProperties(
                conditionalFormatting, styleFormatting);
            formatting = MergeParagraphDefaults(formatting, conditionalFormatting);
        }
        if (_activeTable != null)
            formatting = formatting with { InTable = true,
                TableDepth = _tableStack.Count + 1 };
        _positionalTabParagraphFormatting = formatting;
        if (formatting.TabStops != null)
            _positionalExistingTabPositions.AddRange(formatting.TabStops.Select(x =>
                x.PositionTwips));
        var positionalStyle = _currentParagraphStyleIndex;
        var positionalVisited = new HashSet<int>();
        while (positionalStyle != 0 && positionalVisited.Add(positionalStyle))
        {
            var source = _styles.FirstOrDefault(x => x.Index == positionalStyle);
            if (source == null) break;
            if (source.ParagraphFormatting.TabStops != null)
                _positionalExistingTabPositions.AddRange(
                    source.ParagraphFormatting.TabStops.Select(x => x.PositionTwips));
            _positionalTabParagraphFormatting = MergeParagraphDefaults(
                _positionalTabParagraphFormatting, source.ParagraphFormatting);
            positionalStyle = source.BasedOnIndex ?? 0;
        }
        _positionalTabParagraphFormatting = MergeParagraphDefaults(
            _positionalTabParagraphFormatting, _defaultParagraphFormatting);
        if (_defaultParagraphFormatting.TabStops != null)
            _positionalExistingTabPositions.AddRange(
                _defaultParagraphFormatting.TabStops.Select(x => x.PositionTwips));
        var markProperties = paragraph.ParagraphProperties?.ParagraphMarkRunProperties;
        var markFormatting = ReadCharacterFormatting(markProperties) with
        {
            DeletedRevision = markProperties?.GetFirstChild<Deleted>() != null ? true : null,
            InsertedRevision = markProperties?.GetFirstChild<Inserted>() != null ? true : null,
            DeletedRevisionAuthor = markProperties?.GetFirstChild<Deleted>()?.Author?.Value,
            InsertedRevisionAuthor = markProperties?.GetFirstChild<Inserted>()?.Author?.Value,
            DeletedRevisionAt = markProperties?.GetFirstChild<Deleted>()?.Date?.Value,
            InsertedRevisionAt = markProperties?.GetFirstChild<Inserted>()?.Date?.Value
        };
        _paragraphText = target;
        return DxpDisposable.Create(() =>
        {
            var markCp = target.Length;
            target.Append('\r');
            if (!markFormatting.IsEmpty)
            {
                if (!_runs.TryGetValue(target, out var characterRuns))
                    _runs[target] = characterRuns = new List<DocPlainTextFormatRun>();
                characterRuns.Add(new DocPlainTextFormatRun(markCp, target.Length,
                    markFormatting));
            }
            var styleIndex = styleId != null && _styleById.TryGetValue(styleId, out var found)
                ? found : _styleById.TryGetValue("Normal", out var normal)
                    ? normal : 0;
            if (_positionalTabs.Count != 0)
                formatting = formatting with { TabStops =
                    _positionalTabs.OrderBy(x => x.PositionTwips).ToArray(),
                    ClearedTabPositions = _positionalExistingTabPositions?
                        .Distinct().Where(position => !_positionalTabs.Any(x =>
                            x.PositionTwips == position)).OrderBy(x => x).ToArray() };
            if (styleIndex != 0 || !formatting.IsEmpty)
            {
                if (!_paragraphStyles.TryGetValue(target, out var runs))
                    _paragraphStyles[target] = runs = new List<DocPlainTextParagraphStyleRun>();
                runs.Add(new DocPlainTextParagraphStyleRun(start, target.Length, styleIndex,
                    formatting));
            }
            _paragraphText = previous;
            _positionalTabs = previousPositionalTabs;
            _positionalTabParagraphFormatting = previousPositionalFormatting;
            _positionalExistingTabPositions = previousTabPositions;
            _positionalLineStart = previousLineStart;
            _currentParagraphStyleIndex = previousStyleIndex;
        });
    }

    private bool TryResolveMeasuredRun(Run run, Paragraph paragraph,
        Styles? styles, out DocOpenTypeAdvances? advances, out double sizePoints,
        out bool useKerning)
    {
        advances = null;
        sizePoints = 0;
        useKerning = false;
        string? family = null;
        int? halfPoints = null;
        int? kerningThresholdHalfPoints = null;
        var sources = new List<OpenXmlElement>();
        if (run.RunProperties is { } direct) sources.Add(direct);
        void AddStyleChain(string? id, StyleValues kind)
        {
            var visited = new HashSet<string>(StringComparer.Ordinal);
            while (id != null && visited.Add(id))
            {
                var style = styles?.Elements<Style>().FirstOrDefault(x =>
                    x.StyleId?.Value == id && x.Type?.Value == kind);
                if (style == null) break;
                if (style.StyleRunProperties is { } properties)
                    sources.Add(properties);
                id = style.BasedOn?.Val?.Value;
            }
        }
        AddStyleChain(run.RunProperties?.RunStyle?.Val?.Value, StyleValues.Character);
        AddStyleChain(paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ??
            "Normal", StyleValues.Paragraph);
        if (styles?.GetFirstChild<DocDefaults>()?.GetFirstChild<RunPropertiesDefault>()?
            .RunPropertiesBaseStyle is { } documentDefaults)
            sources.Add(documentDefaults);
        foreach (var source in sources)
        {
            if (source.GetFirstChild<Bold>() is { } bold &&
                bold.Val?.Value != false ||
                source.GetFirstChild<Italic>() is { } italic &&
                italic.Val?.Value != false)
                return false;
            var fonts = source.GetFirstChild<RunFonts>();
            family ??= ResolveThemeLatinFont(fonts?.AsciiTheme?.Value) ??
                fonts?.Ascii?.Value;
            if (halfPoints == null && int.TryParse(
                source.GetFirstChild<FontSize>()?.Val?.Value, out var parsed))
                halfPoints = parsed;
            if (kerningThresholdHalfPoints == null && int.TryParse(
                source.GetFirstChild<Kern>()?.Val?.Value.ToString(), out var threshold))
                kerningThresholdHalfPoints = threshold;
        }
        family ??= _defaultCharacterFormatting.AsciiFontName ?? "Aptos";
        halfPoints ??= _defaultCharacterFormatting.SizeHalfPoints ?? 24;
        if (halfPoints is not > 0) return false;
        advances = DocSystemFontAdvances.Find(family);
        sizePoints = halfPoints.Value / 2.0;
        kerningThresholdHalfPoints ??= _defaultCharacterFormatting
            .KerningThresholdHalfPoints ?? 2;
        useKerning = kerningThresholdHalfPoints > 0 &&
            halfPoints >= kerningThresholdHalfPoints;
        return advances != null;
    }

    // In a simple two- or three-column auto-fit table with occupied
    // horizontal-merge continuations, Word's DOC save fits each story to
    // its text, even when tblGrid carries much wider nominal columns.
    // Only use a measured fit when every
    // contributing run resolves to a locally available font face and size.
    private short[]? TryFitOccupiedMergeGrid(Table table, short[] grid,
        DocCellMargins? margins, Styles? styles)
    {
        if (grid.Length is < 2 or > 3 ||
            table.TableProperties?.GetFirstChild<TableStyle>() != null ||
            table.TableProperties?.GetFirstChild<TableLayout>()?.Type?.Value ==
                TableLayoutValues.Fixed ||
            (table.TableProperties?.GetFirstChild<TableWidth>()?.Type?.Value ==
                TableWidthUnitValues.Dxa ||
             table.TableProperties?.GetFirstChild<TableWidth>()?.Type?.Value ==
                TableWidthUnitValues.Pct))
            return null;
        var rows = table.Elements<TableRow>().ToArray();
        if (rows.Length < 2 || rows.Any(row =>
            row.Elements<TableCell>().Count() != grid.Length))
            return null;
        var baseWidths = new double[grid.Length];
        double mergedMinimum = 0;
        var hasOccupiedMerge = false;
        foreach (var row in rows)
        {
            if (row.TablePropertyExceptions?.GetFirstChild<TableCellMarginDefault>() != null)
                return null;
            var cells = row.Elements<TableCell>().ToArray();
            var merges = cells.Select(cell => cell.TableCellProperties?
                .GetFirstChild<HorizontalMerge>()).ToArray();
            var merged = merges[0]?.Val?.Value == MergedCellValues.Restart &&
                merges.Skip(1).All(merge =>
                    merge?.Val?.Value == MergedCellValues.Continue);
            if (!merged && merges.Any(merge => merge != null)) return null;
            for (var i = 0; i < cells.Length; i++)
            {
                var cell = cells[i];
                var cellProperties = cell.TableCellProperties;
                if (cell.Descendants<Table>().Any() || cell.Elements<Paragraph>().Count() != 1 ||
                    cellProperties?.GetFirstChild<GridSpan>() != null ||
                    cellProperties?.GetFirstChild<TableCellMargin>() != null ||
                    cellProperties?.GetFirstChild<NoWrap>() is { } noWrap &&
                        (noWrap.Val == null ||
                         noWrap.Val.Value == OnOffOnlyValues.On) ||
                    cellProperties?.GetFirstChild<TableCellWidth>()?.Type?.Value ==
                        TableWidthUnitValues.Dxa ||
                    cellProperties?.GetFirstChild<TableCellWidth>()?.Type?.Value ==
                        TableWidthUnitValues.Pct)
                    return null;
                var paragraph = cell.GetFirstChild<Paragraph>()!;
                double content = 0;
                foreach (var run in paragraph.Elements<Run>())
                {
                    if (run.Elements<Break>().Any() || run.Elements<TabChar>().Any())
                        return null;
                    var text = string.Concat(run.Elements<Text>().Select(x => x.Text));
                    if (text.Length == 0) continue;
                    if (text.Any(ch => ch > 0x7f) ||
                        !TryResolveMeasuredRun(run, paragraph, styles,
                            out var advances, out var sizePoints,
                            out var useKerning) ||
                        !advances!.TryMeasure(text, sizePoints, out var width,
                            useKerning))
                        return null;
                    content += width * 20;
                }
                var sideInsets = (margins?.Left ?? 10) + (margins?.Right ?? 10);
                if (merged)
                {
                    if (i > 0 && content > 0) hasOccupiedMerge = true;
                    mergedMinimum = Math.Max(mergedMinimum, content + sideInsets);
                }
                else baseWidths[i] = Math.Max(baseWidths[i], content + sideInsets);
            }
        }
        if (!hasOccupiedMerge || baseWidths.Any(x => x <= 0)) return null;
        var total = baseWidths.Sum();
        if (mergedMinimum > total)
        {
            var scale = mergedMinimum / total;
            for (var i = 0; i < baseWidths.Length; i++)
                baseWidths[i] *= scale;
        }
        if (baseWidths.Sum() > short.MaxValue) return null;
        return baseWidths.Select(x => checked((short)Math.Ceiling(x))).ToArray();
    }

    public override IDisposable VisitTableBegin(Table table, DxpTableModel model,
        DxpIDocumentContext context, DxpITableContext tableContext)
    {
        var story = _storyText ?? (context.CurrentPart == _mainPart ? _text : null);
        if (story == null) return DxpDisposable.Empty;
        if (_activeTable != null)
        {
            _tableStack.Push((_activeTable, _activeRow));
            _activeTable = null;
            _activeRow = null;
        }
        var bidi = ResolveTableProperty<BiDiVisual>(table);
        var rightToLeft = bidi != null && bidi.Val?.Value != OnOffOnlyValues.Off;
        var cellSpacingTwips = ReadCellSpacing(ResolveTableProperty<TableCellSpacing>(table));
        var widths = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>()
            .Select(x => short.TryParse(x.Width?.Value, out var width) && width > 0
                ? width : (short)1440).ToArray() ?? [];
        var tableWidth = ResolveTableProperty<TableWidth>(table);
        DocTablePreferredWidth? preferredWidth = null;
        var widthUnit = tableWidth?.Type?.Value;
        var widthValue = tableWidth?.Width?.Value;
        if (widthUnit == TableWidthUnitValues.Auto)
            preferredWidth = new DocTablePreferredWidth(1, 0);
        else if ((widthUnit == TableWidthUnitValues.Pct ||
            widthUnit == TableWidthUnitValues.Dxa) &&
            ushort.TryParse(widthValue, out var parsedWidth))
            preferredWidth = new DocTablePreferredWidth(
                widthUnit == TableWidthUnitValues.Pct ? (byte)2 : (byte)3,
                parsedWidth);
        var justification = ReadTableJustification(
            ResolveTableProperty<TableJustification>(table));
        var defaultMargins = ResolveTableDefaultMargins(table);
        // The calculated widths belong to this story's table instance.
        widths = TryFitOccupiedMergeGrid(table, widths, defaultMargins,
            _mainPart?.StyleDefinitionsPart?.Styles) ?? widths;
        var tableStyleId = table.TableProperties?.GetFirstChild<TableStyle>()?.Val?.Value;
        if (defaultMargins == null && tableStyleId == null)
        {
            var builtInTableStyle = _mainPart?.StyleDefinitionsPart?.Styles?
                .Elements<Style>().FirstOrDefault(x =>
                    x.Type?.Value == StyleValues.Table &&
                    x.StyleId?.Value == "TableNormal");
            defaultMargins = ReadCellMargins(builtInTableStyle?
                .StyleTableProperties?.GetFirstChild<TableCellMarginDefault>()) ??
                new DocCellMargins(Top: 0, Left: 10, Bottom: 0, Right: 10);
        }
        if (defaultMargins == null && tableStyleId != null)
            defaultMargins = new DocCellMargins(Top: 0, Left: 0,
                Bottom: 0, Right: 0);
        var hasVaryingHorizontalMargins = table.Elements<TableRow>().Any(row =>
        {
            var margins = ReadCellMargins(row.TablePropertyExceptions?
                .GetFirstChild<TableCellMarginDefault>());
            return margins?.Left is ushort left &&
                    left != (defaultMargins?.Left ?? 108) ||
                margins?.Right is ushort right &&
                    right != (defaultMargins?.Right ?? 108);
        });
        var cells = table.Elements<TableRow>()
            .SelectMany(row => row.Elements<TableCell>()).ToArray();
        var hasPartialVerticalCellBorders = cells.Any(cell =>
        {
            var borders = cell.TableCellProperties?.GetFirstChild<TableCellBorders>();
            return borders?.GetFirstChild<LeftBorder>() != null ||
                borders?.GetFirstChild<RightBorder>() != null ||
                borders?.GetFirstChild<StartBorder>() != null ||
                borders?.GetFirstChild<EndBorder>() != null;
        }) &&
            cells.Any(cell => cell.TableCellProperties?
                .GetFirstChild<TableCellBorders>() == null);
        // Word's DOC save supplies a zero indent for tables with partial
        // vertical cell borders or complete direct table borders when their
        // horizontal cell insets are uniform.
        var explicitIndent = ReadTableIndent(ResolveTableProperty<TableIndentation>(table));
        // Word's DOC save omits zero indent on styled tables without direct
        // borders; retaining it shifts the cell text left by the side inset.
        if (explicitIndent == 0 && tableStyleId != null &&
            table.TableProperties?.GetFirstChild<TableBorders>() == null)
            explicitIndent = null;
        var indentTwips = explicitIndent
            ?? ((tableStyleId == null && hasPartialVerticalCellBorders &&
                !hasVaryingHorizontalMargins) ||
                (table.TableProperties?.GetFirstChild<TableBorders>() is
                    { TopBorder: not null, LeftBorder: not null,
                      BottomBorder: not null, RightBorder: not null,
                      InsideHorizontalBorder: not null,
                      InsideVerticalBorder: not null } &&
                    !hasVaryingHorizontalMargins &&
                    (defaultMargins?.Left ?? 108) ==
                        (defaultMargins?.Right ?? 108))
                ? (short)0 : null);
        ushort? tableStyleIndex = tableStyleId != null &&
            _styleById.TryGetValue(tableStyleId, out var foundStyleIndex)
            ? checked((ushort)foundStyleIndex) : null;
        var look = table.TableProperties?.GetFirstChild<TableLook>();
        ushort.TryParse(look?.Val?.Value, System.Globalization.NumberStyles.HexNumber,
            System.Globalization.CultureInfo.InvariantCulture, out var lookMask);
        if (look == null) lookMask = 0x04A0;
        // Word fills omitted flags on an attribute-only tblLook with its
        // default column emphasis and disabled vertical banding.
        var attributeOnlyLook = look != null && look.Val?.Value == null;
        var firstRowLook = look?.FirstRow?.Value ?? (lookMask & 0x0020) != 0;
        var lastRowLook = look?.LastRow?.Value ?? (lookMask & 0x0040) != 0;
        var firstColumnLook = look?.FirstColumn?.Value ??
            (attributeOnlyLook || (lookMask & 0x0080) != 0);
        var lastColumnLook = look?.LastColumn?.Value ??
            (attributeOnlyLook || (lookMask & 0x0100) != 0);
        var noHorizontalBand = look?.NoHorizontalBand?.Value ?? (lookMask & 0x0200) != 0;
        var noVerticalBand = look?.NoVerticalBand?.Value ??
            (attributeOnlyLook || (lookMask & 0x0400) != 0);
        var horizontalBandSize = ResolveTableProperty<TableStyleRowBandSize>(table)?
            .Val?.Value ?? 0;
        if (noHorizontalBand) horizontalBandSize = 0;
        var verticalBandSize = ResolveTableProperty<TableStyleColumnBandSize>(table)?
            .Val?.Value ?? 0;
        if (noVerticalBand) verticalBandSize = 0;
        var firstColumnBandOffset = firstColumnLook &&
            (ResolveConditionalTableCellStyleProperty<Shading>(table,
                 TableStyleOverrideValues.FirstColumn) != null ||
             ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                 TableStyleOverrideValues.FirstColumn) != null ||
             ResolveConditionalTableRunProperties(table,
                 TableStyleOverrideValues.FirstColumn) != null ||
             ResolveConditionalTableParagraphProperties(table,
                 TableStyleOverrideValues.FirstColumn) != null) ? 1 : 0;
        _activeTable = new TableCapture(story, widths,
            table.Elements<TableRow>().Count(),
            ResolveEffectiveTableBorders(table),
            ReadTableBackgroundShading(table.TableProperties?.GetFirstChild<Shading>()),
            ResolveTableProperty<TableLayout>(table)?.Type?.Value !=
                TableLayoutValues.Fixed, preferredWidth, indentTwips,
            justification,
            defaultMargins,
            rightToLeft, cellSpacingTwips,
            ResolveTableCellStyleProperty<Shading>(table),
            firstRowLook
                ? ResolveConditionalTableCellStyleProperty<Shading>(table,
                    TableStyleOverrideValues.FirstRow) : null,
            lastRowLook
                ? ResolveConditionalTableCellStyleProperty<Shading>(table,
                    TableStyleOverrideValues.LastRow) : null,
            firstRowLook
                ? ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                    TableStyleOverrideValues.FirstRow) : null,
            lastRowLook
                ? ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                    TableStyleOverrideValues.LastRow) : null,
            firstColumnLook
                ? ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                    TableStyleOverrideValues.FirstColumn) : null,
            lastColumnLook
                ? ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                    TableStyleOverrideValues.LastColumn) : null,
            firstRowLook && firstColumnLook
                ? ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                    TableStyleOverrideValues.NorthWestCell) : null,
            firstRowLook && lastColumnLook
                ? ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                    TableStyleOverrideValues.NorthEastCell) : null,
            lastRowLook && firstColumnLook
                ? ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                    TableStyleOverrideValues.SouthWestCell) : null,
            lastRowLook && lastColumnLook
                ? ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                    TableStyleOverrideValues.SouthEastCell) : null,
            horizontalBandSize == 0
                ? null : ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                    TableStyleOverrideValues.Band1Horizontal),
            horizontalBandSize == 0
                ? null : ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                    TableStyleOverrideValues.Band2Horizontal),
            verticalBandSize == 0
                ? null : ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                    TableStyleOverrideValues.Band1Vertical),
            verticalBandSize == 0
                ? null : ResolveConditionalTableCellStyleProperty<TableCellBorders>(table,
                    TableStyleOverrideValues.Band2Vertical),
            horizontalBandSize == 0
                ? null : ResolveConditionalTableCellStyleProperty<Shading>(table,
                    TableStyleOverrideValues.Band1Horizontal),
            horizontalBandSize == 0
                ? null : ResolveConditionalTableCellStyleProperty<Shading>(table,
                    TableStyleOverrideValues.Band2Horizontal),
            horizontalBandSize,
            firstRowLook ? 1 : 0,
            verticalBandSize == 0
                ? null : ResolveConditionalTableCellStyleProperty<Shading>(table,
                    TableStyleOverrideValues.Band1Vertical),
            verticalBandSize == 0
                ? null : ResolveConditionalTableCellStyleProperty<Shading>(table,
                    TableStyleOverrideValues.Band2Vertical),
            verticalBandSize,
            firstColumnBandOffset,
            firstColumnLook
                ? ResolveConditionalTableCellStyleProperty<Shading>(table,
                    TableStyleOverrideValues.FirstColumn) : null,
            lastColumnLook
                ? ResolveConditionalTableCellStyleProperty<Shading>(table,
                    TableStyleOverrideValues.LastColumn) : null,
            firstRowLook && firstColumnLook
                ? ResolveConditionalTableCellStyleProperty<Shading>(table,
                    TableStyleOverrideValues.NorthWestCell) : null,
            firstRowLook && lastColumnLook
                ? ResolveConditionalTableCellStyleProperty<Shading>(table,
                    TableStyleOverrideValues.NorthEastCell) : null,
            lastRowLook && firstColumnLook
                ? ResolveConditionalTableCellStyleProperty<Shading>(table,
                    TableStyleOverrideValues.SouthWestCell) : null,
            lastRowLook && lastColumnLook
                ? ResolveConditionalTableCellStyleProperty<Shading>(table,
                    TableStyleOverrideValues.SouthEastCell) : null,
            tableStyleIndex, checked(++_nextTableGroupId));
        _activeTable.DefaultVerticalAlignment = ResolveTableCellStyleProperty<
            TableCellVerticalAlignment>(table)?.Val?.Value;
        var styleNoWrap = ResolveTableCellStyleProperty<NoWrap>(table);
        static bool? NoWrapValue(NoWrap? value) => value == null ? null :
            value.Val == null || value.Val.Value == OnOffOnlyValues.On;
        _activeTable.DefaultNoWrap = NoWrapValue(styleNoWrap);
        _activeTable.FirstRowNoWrap = firstRowLook ? NoWrapValue(
            ResolveConditionalTableCellStyleProperty<NoWrap>(table,
                TableStyleOverrideValues.FirstRow)) : null;
        _activeTable.LastRowNoWrap = lastRowLook ? NoWrapValue(
            ResolveConditionalTableCellStyleProperty<NoWrap>(table,
                TableStyleOverrideValues.LastRow)) : null;
        _activeTable.FirstColumnNoWrap = firstColumnLook ? NoWrapValue(
            ResolveConditionalTableCellStyleProperty<NoWrap>(table,
                TableStyleOverrideValues.FirstColumn)) : null;
        _activeTable.LastColumnNoWrap = lastColumnLook ? NoWrapValue(
            ResolveConditionalTableCellStyleProperty<NoWrap>(table,
                TableStyleOverrideValues.LastColumn)) : null;
        _activeTable.NorthWestNoWrap = firstRowLook && firstColumnLook
            ? NoWrapValue(ResolveConditionalTableCellStyleProperty<NoWrap>(table,
                TableStyleOverrideValues.NorthWestCell)) : null;
        _activeTable.NorthEastNoWrap = firstRowLook && lastColumnLook
            ? NoWrapValue(ResolveConditionalTableCellStyleProperty<NoWrap>(table,
                TableStyleOverrideValues.NorthEastCell)) : null;
        _activeTable.SouthWestNoWrap = lastRowLook && firstColumnLook
            ? NoWrapValue(ResolveConditionalTableCellStyleProperty<NoWrap>(table,
                TableStyleOverrideValues.SouthWestCell)) : null;
        _activeTable.SouthEastNoWrap = lastRowLook && lastColumnLook
            ? NoWrapValue(ResolveConditionalTableCellStyleProperty<NoWrap>(table,
                TableStyleOverrideValues.SouthEastCell)) : null;
        if (horizontalBandSize > 0)
        {
            _activeTable.Band1HorizontalNoWrap = NoWrapValue(
                ResolveConditionalTableCellStyleProperty<NoWrap>(table,
                    TableStyleOverrideValues.Band1Horizontal));
            _activeTable.Band2HorizontalNoWrap = NoWrapValue(
                ResolveConditionalTableCellStyleProperty<NoWrap>(table,
                    TableStyleOverrideValues.Band2Horizontal));
        }
        if (verticalBandSize > 0)
        {
            _activeTable.Band1VerticalNoWrap = NoWrapValue(
                ResolveConditionalTableCellStyleProperty<NoWrap>(table,
                    TableStyleOverrideValues.Band1Vertical));
            _activeTable.Band2VerticalNoWrap = NoWrapValue(
                ResolveConditionalTableCellStyleProperty<NoWrap>(table,
                    TableStyleOverrideValues.Band2Vertical));
        }
        if (firstRowLook)
            _activeTable.FirstRowVerticalAlignment =
                ResolveConditionalTableCellStyleProperty<TableCellVerticalAlignment>(
                    table, TableStyleOverrideValues.FirstRow)?.Val?.Value;
        if (lastRowLook)
            _activeTable.LastRowVerticalAlignment =
                ResolveConditionalTableCellStyleProperty<TableCellVerticalAlignment>(
                    table, TableStyleOverrideValues.LastRow)?.Val?.Value;
        if (firstColumnLook)
            _activeTable.FirstColumnVerticalAlignment =
                ResolveConditionalTableCellStyleProperty<TableCellVerticalAlignment>(
                    table, TableStyleOverrideValues.FirstColumn)?.Val?.Value;
        if (lastColumnLook)
            _activeTable.LastColumnVerticalAlignment =
                ResolveConditionalTableCellStyleProperty<TableCellVerticalAlignment>(
                    table, TableStyleOverrideValues.LastColumn)?.Val?.Value;
        if (firstRowLook && firstColumnLook)
            _activeTable.NorthWestVerticalAlignment =
                ResolveConditionalTableCellStyleProperty<TableCellVerticalAlignment>(
                    table, TableStyleOverrideValues.NorthWestCell)?.Val?.Value;
        if (firstRowLook && lastColumnLook)
            _activeTable.NorthEastVerticalAlignment =
                ResolveConditionalTableCellStyleProperty<TableCellVerticalAlignment>(
                    table, TableStyleOverrideValues.NorthEastCell)?.Val?.Value;
        if (lastRowLook && firstColumnLook)
            _activeTable.SouthWestVerticalAlignment =
                ResolveConditionalTableCellStyleProperty<TableCellVerticalAlignment>(
                    table, TableStyleOverrideValues.SouthWestCell)?.Val?.Value;
        if (lastRowLook && lastColumnLook)
            _activeTable.SouthEastVerticalAlignment =
                ResolveConditionalTableCellStyleProperty<TableCellVerticalAlignment>(
                    table, TableStyleOverrideValues.SouthEastCell)?.Val?.Value;
        if (horizontalBandSize > 0)
        {
            _activeTable.Band1HorizontalVerticalAlignment =
                ResolveConditionalTableCellStyleProperty<TableCellVerticalAlignment>(
                    table, TableStyleOverrideValues.Band1Horizontal)?.Val?.Value;
            _activeTable.Band2HorizontalVerticalAlignment =
                ResolveConditionalTableCellStyleProperty<TableCellVerticalAlignment>(
                    table, TableStyleOverrideValues.Band2Horizontal)?.Val?.Value;
        }
        if (verticalBandSize > 0)
        {
            _activeTable.Band1VerticalVerticalAlignment =
                ResolveConditionalTableCellStyleProperty<TableCellVerticalAlignment>(
                    table, TableStyleOverrideValues.Band1Vertical)?.Val?.Value;
            _activeTable.Band2VerticalVerticalAlignment =
                ResolveConditionalTableCellStyleProperty<TableCellVerticalAlignment>(
                    table, TableStyleOverrideValues.Band2Vertical)?.Val?.Value;
        }
        if (look != null)
            _activeTable.LookMask = (ushort)((firstRowLook ? 0x0020 : 0) |
                (lastRowLook ? 0x0040 : 0) |
                (firstColumnLook ? 0x0080 : 0) |
                (lastColumnLook ? 0x0100 : 0) |
                (noHorizontalBand ? 0x0200 : 0) |
                (noVerticalBand ? 0x0400 : 0));
        _activeTable.HasFitText = table.Elements<TableRow>()
            .SelectMany(row => row.Elements<TableCell>())
            .Any(cell => cell.TableCellProperties?
                .GetFirstChild<TableCellFitText>() is { } fitText &&
                (fitText.Val == null || fitText.Val.Value == OnOffOnlyValues.On));
        if (firstRowLook && ResolveConditionalTableRunProperties(table,
            TableStyleOverrideValues.FirstRow) is { } firstRowRun)
            _activeTable.FirstRowRunFormatting = ReadCharacterFormatting(firstRowRun);
        if (lastRowLook && ResolveConditionalTableRunProperties(table,
            TableStyleOverrideValues.LastRow) is { } lastRowRun)
            _activeTable.LastRowRunFormatting = ReadCharacterFormatting(lastRowRun);
        if (firstColumnLook && ResolveConditionalTableRunProperties(table,
            TableStyleOverrideValues.FirstColumn) is { } firstColumnRun)
            _activeTable.FirstColumnRunFormatting = ReadCharacterFormatting(firstColumnRun);
        if (lastColumnLook && ResolveConditionalTableRunProperties(table,
            TableStyleOverrideValues.LastColumn) is { } lastColumnRun)
            _activeTable.LastColumnRunFormatting = ReadCharacterFormatting(lastColumnRun);
        if (horizontalBandSize > 0)
        {
            if (ResolveConditionalTableRunProperties(table,
                TableStyleOverrideValues.Band1Horizontal) is { } band1Run)
                _activeTable.Band1HorizontalRunFormatting = ReadCharacterFormatting(band1Run);
            if (ResolveConditionalTableRunProperties(table,
                TableStyleOverrideValues.Band2Horizontal) is { } band2Run)
                _activeTable.Band2HorizontalRunFormatting = ReadCharacterFormatting(band2Run);
        }
        if (verticalBandSize > 0)
        {
            if (ResolveConditionalTableRunProperties(table,
                TableStyleOverrideValues.Band1Vertical) is { } band1Run)
                _activeTable.Band1VerticalRunFormatting = ReadCharacterFormatting(band1Run);
            if (ResolveConditionalTableRunProperties(table,
                TableStyleOverrideValues.Band2Vertical) is { } band2Run)
                _activeTable.Band2VerticalRunFormatting = ReadCharacterFormatting(band2Run);
        }
        if (firstRowLook && firstColumnLook &&
            ResolveConditionalTableRunProperties(table,
                TableStyleOverrideValues.NorthWestCell) is { } northWestRun)
            _activeTable.NorthWestRunFormatting = ReadCharacterFormatting(northWestRun);
        if (firstRowLook && lastColumnLook &&
            ResolveConditionalTableRunProperties(table,
                TableStyleOverrideValues.NorthEastCell) is { } northEastRun)
            _activeTable.NorthEastRunFormatting = ReadCharacterFormatting(northEastRun);
        if (lastRowLook && firstColumnLook &&
            ResolveConditionalTableRunProperties(table,
                TableStyleOverrideValues.SouthWestCell) is { } southWestRun)
            _activeTable.SouthWestRunFormatting = ReadCharacterFormatting(southWestRun);
        if (lastRowLook && lastColumnLook &&
            ResolveConditionalTableRunProperties(table,
                TableStyleOverrideValues.SouthEastCell) is { } southEastRun)
            _activeTable.SouthEastRunFormatting = ReadCharacterFormatting(southEastRun);
        foreach (var condition in new[]
        {
            TableStyleOverrideValues.FirstRow, TableStyleOverrideValues.LastRow,
            TableStyleOverrideValues.FirstColumn, TableStyleOverrideValues.LastColumn,
            TableStyleOverrideValues.Band1Horizontal,
            TableStyleOverrideValues.Band2Horizontal,
            TableStyleOverrideValues.Band1Vertical,
            TableStyleOverrideValues.Band2Vertical,
            TableStyleOverrideValues.NorthWestCell,
            TableStyleOverrideValues.NorthEastCell,
            TableStyleOverrideValues.SouthWestCell,
            TableStyleOverrideValues.SouthEastCell
        })
        {
            var enabled = condition == TableStyleOverrideValues.FirstRow && firstRowLook ||
                condition == TableStyleOverrideValues.LastRow && lastRowLook ||
                condition == TableStyleOverrideValues.FirstColumn && firstColumnLook ||
                condition == TableStyleOverrideValues.LastColumn && lastColumnLook ||
                (condition == TableStyleOverrideValues.Band1Horizontal ||
                    condition == TableStyleOverrideValues.Band2Horizontal) &&
                    horizontalBandSize > 0 ||
                (condition == TableStyleOverrideValues.Band1Vertical ||
                    condition == TableStyleOverrideValues.Band2Vertical) &&
                    verticalBandSize > 0 ||
                condition == TableStyleOverrideValues.NorthWestCell &&
                    firstRowLook && firstColumnLook ||
                condition == TableStyleOverrideValues.NorthEastCell &&
                    firstRowLook && lastColumnLook ||
                condition == TableStyleOverrideValues.SouthWestCell &&
                    lastRowLook && firstColumnLook ||
                condition == TableStyleOverrideValues.SouthEastCell &&
                    lastRowLook && lastColumnLook;
            if (enabled && ResolveConditionalTableParagraphProperties(table,
                condition) is { } properties)
                _activeTable.ConditionalParagraphFormatting[condition] =
                    ReadParagraphFormatting(properties);
        }
        return DxpDisposable.Create(() =>
        {
            _activeTable = null;
            if (_tableStack.Count == 0) return;
            (_activeTable, _activeRow) = _tableStack.Pop();
        });
    }

    private T? ResolveTableProperty<T>(Table table) where T : OpenXmlElement
    {
        var direct = table.TableProperties?.GetFirstChild<T>();
        if (direct != null) return direct;
        var styleId = table.TableProperties?.GetFirstChild<TableStyle>()?.Val?.Value;
        var styles = _mainPart?.StyleDefinitionsPart?.Styles;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (styleId != null && styles != null && visited.Add(styleId))
        {
            var style = styles.Elements<Style>().FirstOrDefault(x =>
                x.Type?.Value == StyleValues.Table && x.StyleId?.Value == styleId);
            if (style == null) break;
            var inherited = style.StyleTableProperties?.GetFirstChild<T>();
            if (inherited != null) return inherited;
            styleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    private T? ResolveTableCellStyleProperty<T>(Table table) where T : OpenXmlElement
    {
        var styleId = table.TableProperties?.GetFirstChild<TableStyle>()?.Val?.Value;
        var styles = _mainPart?.StyleDefinitionsPart?.Styles;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (styleId != null && styles != null && visited.Add(styleId))
        {
            var style = styles.Elements<Style>().FirstOrDefault(x =>
                x.Type?.Value == StyleValues.Table && x.StyleId?.Value == styleId);
            if (style == null) break;
            var inherited = style.StyleTableCellProperties?.GetFirstChild<T>();
            if (inherited != null) return inherited;
            styleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    private T? ResolveConditionalTableCellStyleProperty<T>(Table table,
        TableStyleOverrideValues condition) where T : OpenXmlElement
    {
        var styleId = table.TableProperties?.GetFirstChild<TableStyle>()?.Val?.Value;
        var styles = _mainPart?.StyleDefinitionsPart?.Styles;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (styleId != null && styles != null && visited.Add(styleId))
        {
            var style = styles.Elements<Style>().FirstOrDefault(x =>
                x.Type?.Value == StyleValues.Table && x.StyleId?.Value == styleId);
            if (style == null) break;
            var inherited = style.Elements<TableStyleProperties>()
                .FirstOrDefault(x => x.Type?.Value == condition)?
                .TableStyleConditionalFormattingTableCellProperties?
                .GetFirstChild<T>();
            if (inherited != null) return inherited;
            styleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    private RunPropertiesBaseStyle? ResolveConditionalTableRunProperties(
        Table table, TableStyleOverrideValues condition)
    {
        var styleId = table.TableProperties?.GetFirstChild<TableStyle>()?.Val?.Value;
        var styles = _mainPart?.StyleDefinitionsPart?.Styles;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (styleId != null && styles != null && visited.Add(styleId))
        {
            var style = styles.Elements<Style>().FirstOrDefault(x =>
                x.Type?.Value == StyleValues.Table && x.StyleId?.Value == styleId);
            if (style == null) break;
            var properties = style.Elements<TableStyleProperties>()
                .FirstOrDefault(x => x.Type?.Value == condition)?
                .RunPropertiesBaseStyle;
            if (properties != null) return properties;
            styleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }

    private StyleParagraphProperties? ResolveConditionalTableParagraphProperties(
        Table table, TableStyleOverrideValues condition)
    {
        var styleId = table.TableProperties?.GetFirstChild<TableStyle>()?.Val?.Value;
        var styles = _mainPart?.StyleDefinitionsPart?.Styles;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (styleId != null && styles != null && visited.Add(styleId))
        {
            var style = styles.Elements<Style>().FirstOrDefault(x =>
                x.Type?.Value == StyleValues.Table && x.StyleId?.Value == styleId);
            if (style == null) break;
            var properties = style.Elements<TableStyleProperties>()
                .FirstOrDefault(x => x.Type?.Value == condition)?
                .StyleParagraphProperties;
            if (properties != null) return properties;
            styleId = style.BasedOn?.Val?.Value;
        }
        return null;
    }


    private DocCellMargins? ResolveTableDefaultMargins(Table table)
    {
        var styleId = table.TableProperties?.GetFirstChild<TableStyle>()?.Val?.Value;
        var styles = _mainPart?.StyleDefinitionsPart?.Styles;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        var chain = new List<TableCellMarginDefault>();
        while (styleId != null && styles != null && visited.Add(styleId))
        {
            var style = styles.Elements<Style>().FirstOrDefault(x =>
                x.Type?.Value == StyleValues.Table && x.StyleId?.Value == styleId);
            if (style == null) break;
            if (style.StyleTableProperties?.GetFirstChild<TableCellMarginDefault>()
                is { } styleMargins)
                chain.Add(styleMargins);
            styleId = style.BasedOn?.Val?.Value;
        }
        DocCellMargins? result = null;
        for (var i = chain.Count - 1; i >= 0; i--)
            result = MergeCellMargins(result, ReadCellMargins(chain[i]));
        if (result != null)
            result = result with
            {
                Left = result.Left ?? 0,
                Right = result.Right ?? 0
            };
        return MergeCellMargins(result, ReadCellMargins(
            table.TableProperties?.GetFirstChild<TableCellMarginDefault>()));
    }

    private static DocCellMargins? MergeCellMargins(DocCellMargins? inherited,
        DocCellMargins? direct) => direct == null ? inherited : new DocCellMargins(
            direct.Top ?? inherited?.Top,
            direct.Left ?? inherited?.Left,
            direct.Bottom ?? inherited?.Bottom,
            direct.Right ?? inherited?.Right);

    private static short? ReadTableIndent(TableIndentation? indent)
    {
        if (indent == null) return null;
        if (indent.Type?.Value != TableWidthUnitValues.Dxa ||
            indent.Width?.Value is not int twips || twips is < -31560 or > 31680)
            throw new NotSupportedException("The DOCX table indent must be a DOC twip width.");
        return checked((short)twips);
    }

    private static byte? ReadTableJustification(TableJustification? justification)
    {
        if (justification == null) return null;
        var value = justification.Val?.Value;
        if (value == TableRowAlignmentValues.Left) return 0;
        if (value == TableRowAlignmentValues.Center) return 1;
        if (value == TableRowAlignmentValues.Right) return 2;
        throw new NotSupportedException("The DOCX table justification is unsupported.");
    }

    private static ushort? ReadCellSpacing(TableCellSpacing? spacing)
    {
        if (spacing == null) return null;
        if (spacing.Type?.Value != TableWidthUnitValues.Dxa ||
            !ushort.TryParse(spacing.Width?.Value, out var value) || value > 15840)
            throw new NotSupportedException("Only twip table cell spacing up to 11 inches is supported.");
        return value;
    }

    public override IDisposable VisitTableRowBegin(TableRow row,
        DxpITableRowContext rowContext, DxpIDocumentContext context)
    {
        if (_activeTable == null) return DxpDisposable.Empty;
        if (_activeRow != null)
            throw new InvalidDataException("A DOC table row is already open.");
        var table = _activeTable;
        _activeRow = new RowCapture();
        _activeRow.TableBorders = ReadTableBorders(row.TablePropertyExceptions?
            .GetFirstChild<TableBorders>());
        var rowProperties = row.TableRowProperties;
        DocTablePreferredWidth? ReadGapWidth(int count, string elementName,
            bool before)
        {
            if (count <= 0) return null;
            if (count > table.GridWidths.Count)
                throw new NotSupportedException("A table row skips beyond its grid.");
            var element = rowProperties?.ChildElements.FirstOrDefault(x =>
                x.LocalName == elementName);
            var unit = element?.GetAttribute("type",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
            var widthText = element?.GetAttribute("w",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
            var widthUnit = unit switch
            {
                "auto" => (byte)1,
                "pct" => (byte)2,
                "dxa" => (byte)3,
                null when element == null => (byte)3,
                _ => throw new NotSupportedException("The DOCX row-gap width unit is unsupported in DOC.")
            };
            if (element != null && !ushort.TryParse(widthText, out _))
                throw new NotSupportedException("The DOCX row-gap width is invalid.");
            var width = element == null
                ? (before ? table.GridWidths.Take(count) :
                    table.GridWidths.Skip(table.GridWidths.Count - count)).Sum(x => (int)x)
                : ushort.Parse(widthText!, System.Globalization.CultureInfo.InvariantCulture);
            if ((widthUnit == 1 && width != 0) ||
                (widthUnit == 2 && width > 5000) ||
                (widthUnit == 3 && width > 31680))
                throw new NotSupportedException("The DOCX row-gap width exceeds the DOC limit.");
            return new DocTablePreferredWidth(widthUnit, checked((ushort)width));
        }
        var beforeCount = rowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
        var afterCount = rowProperties?.GetFirstChild<GridAfter>()?.Val?.Value ?? 0;
        var beforeWidthOmitted = beforeCount > 0 && rowProperties?
            .ChildElements.All(x => x.LocalName != "wBefore") == true;
        var afterWidthOmitted = afterCount > 0 && rowProperties?
            .ChildElements.All(x => x.LocalName != "wAfter") == true;
        var widthBefore = beforeWidthOmitted
            ? new DocTablePreferredWidth(3, 0)
            : ReadGapWidth(beforeCount, "wBefore", true);
        var widthAfter = afterWidthOmitted
            ? new DocTablePreferredWidth(3, 0)
            : ReadGapWidth(afterCount, "wAfter", false);
        var omittedBeforeGridWidth = beforeWidthOmitted
            ? table.GridWidths.Take(beforeCount).Sum(x => (int)x) : 0;
        var rowCellSpacing = ReadCellSpacing(rowProperties?
            .GetFirstChild<TableCellSpacing>() ?? row.TablePropertyExceptions?
            .GetFirstChild<TableCellSpacing>()) ?? table.CellSpacingTwips;
        var rowDefaultMargins = MergeCellMargins(table.DefaultMargins,
            ReadCellMargins(row.TablePropertyExceptions?
                .GetFirstChild<TableCellMarginDefault>()));
        var header = rowProperties?.GetFirstChild<TableHeader>();
        bool? isHeader = header == null ? null :
            header.Val?.Value != OnOffOnlyValues.Off;
        var cantSplit = rowProperties?.GetFirstChild<CantSplit>();
        bool? cannotSplit = cantSplit == null ? null :
            cantSplit.Val?.Value != OnOffOnlyValues.Off;
        var height = rowProperties?.GetFirstChild<TableRowHeight>();
        short? rowHeight = null;
        if (short.TryParse(height?.Val?.Value.ToString(), out var parsedHeight))
            rowHeight = height?.HeightType?.Value == HeightRuleValues.Exact
                ? checked((short)-parsedHeight) : parsedHeight;
        return DxpDisposable.Create(() =>
        {
            var cellCount = _activeRow!.CellCount;
            if (_activeRow.PendingHorizontalContinuations != 0)
                throw new InvalidDataException("A DOCX horizontal merge has missing continuation cells.");
            if (cellCount is < 1 or > 63)
                throw new NotSupportedException("A DOC table row must have 1â€“63 cells.");
            var edges = new short[cellCount + 1];
            // In RTL tables with outer borders and an inferred zero indent,
            // Word anchors the cell grid at zero. Side-inset compensation
            // moves the edges left and changes which border wins at a join.
            if ((table.IndentTwips != null || table.HasFitText) &&
                !(table.RightToLeft && table.IndentTwips == 0 &&
                  !table.HasFitText && table.Borders != null))
            {
                var leftMargin = rowDefaultMargins?.Left ?? 108;
                var rightMargin = rowDefaultMargins?.Right ?? 108;
                // Keep every row on the same cell origin when fit text
                // requires Word's default cell-inset compensation.
                edges[0] = checked((short)((table.IndentTwips ?? 0) -
                    (leftMargin + rightMargin) / 2));
            }
            if (omittedBeforeGridWidth != 0)
                edges[0] = checked((short)(edges[0] + omittedBeforeGridWidth));
            else if (beforeCount > 0 && widthBefore is { Unit: 1 or 2 })
                edges[0] = checked((short)(edges[0] +
                    table.GridWidths.Take(beforeCount).Sum(x => (int)x)));
            // A styled table without an explicit indent places the left
            // border inside the DOC cell origin. Shift the stored grid by
            // half that border width so Word draws it at the DOCX edge.
            if (!table.AutoFit && table.StyleIndex != null &&
                table.IndentTwips == null &&
                table.Borders?.Left is { WidthEighthPoints: > 0 } placementBorder)
                edges[0] = checked((short)(edges[0] +
                    (placementBorder.WidthEighthPoints * 5 + 2) / 4));
            for (var i = 0; i < cellCount; i++)
                edges[i + 1] = checked((short)(edges[i] +
                    _activeRow.CellWidths[i]));
            var start = table.Story.Length;
            table.Story.Append(_tableStack.Count == 0 ? '\u0007' : '\r');
            if (!_paragraphStyles.TryGetValue(table.Story, out var runs))
                _paragraphStyles[table.Story] = runs = new List<DocPlainTextParagraphStyleRun>();
            runs.Add(new DocPlainTextParagraphStyleRun(start, table.Story.Length, 0,
                DocParagraphFormatting.Empty with
                {
                    InTable = true, TableTerminator = _tableStack.Count == 0 ? true : null,
                    TableDepth = _tableStack.Count + 1,
                    InnerTableRow = _tableStack.Count == 0 ? null : true,
                    TableCellEdges = edges,
                    TableWidthBefore = widthBefore,
                    TableWidthAfter = widthAfter,
                    TableHeader = isHeader, TableCantSplit = cannotSplit,
                    TableRowHeightTwips = rowHeight,
                    TableAutoFit = table.AutoFit,
                    TableRightToLeft = table.RightToLeft,
                    TableStyleIndex = table.StyleIndex,
                    TableLookMask = table.LookMask,
                    ParagraphGroupId = table.GroupId,
                    TableGroupId = table.GroupId,
                    TableCellSpacingTwips = rowCellSpacing,
                    TablePreferredWidth = table.PreferredWidth,
                    TableBackgroundShading = table.BackgroundShading,
                    TableIndentTwips = table.IndentTwips,
                    TableRowOriginTwips = table.IndentTwips == 0 &&
                        table.Borders?.Left is { WidthEighthPoints: > 0 } leftBorder
                        ? checked((short)(((rowDefaultMargins?.Left ?? 108) +
                            (rowDefaultMargins?.Right ?? 108)) / 2 +
                            (leftBorder.WidthEighthPoints * 5 + 2) / 4))
                        : null,
                    TableJustification = table.Justification,
                    TableCellPreferredWidths = _activeRow.PreferredCellWidths.Any(x => x != null)
                        ? _activeRow.PreferredCellWidths.ToArray() : null,
                    TableCellNoWraps = _activeRow.NoWraps.Any(x => x != null)
                        ? _activeRow.NoWraps.ToArray() : null,
                    TableCellFitTexts = _activeRow.FitTexts.Any(x => x != null)
                        ? _activeRow.FitTexts.ToArray() : null,
                    TableDefaultCellMargins = rowDefaultMargins,
                    TableCellMargins = _activeRow.CellMargins.Any(x => x != null)
                        ? _activeRow.CellMargins.ToArray() : null,
                    TableCellShadings = _activeRow.Shadings.Any(x => x != null)
                        ? _activeRow.Shadings.ToArray() : null,
                    TableCellVerticalAlignments = _activeRow.VerticalAlignments.Any(x => x != null)
                        ? _activeRow.VerticalAlignments.ToArray() : null,
                    TableCellTextFlows = _activeRow.TextFlows.Any(x => x != null)
                        ? _activeRow.TextFlows.ToArray() : null,
                    TableCellHideMarks = _activeRow.HideMarks.Any(x => x != null)
                        ? _activeRow.HideMarks.ToArray() : null,
                    TableCellVerticalMerges = _activeRow.VerticalMerges.Any(x => x != null)
                        ? _activeRow.VerticalMerges.ToArray() : null,
                    TableCellHorizontalMerges = _activeRow.HorizontalMerges.Any(x => x != null)
                        ? _activeRow.HorizontalMerges.ToArray() : null,
                    TableCellBorders = _activeRow.Borders.Any(x => x != null)
                        ? _activeRow.Borders.ToArray() : null
                }));
            _activeRow = null;
            table.RowIndex++;
        });
    }

    public override IDisposable VisitTableCellBegin(TableCell cell,
        DxpITableCellContext cellContext, DxpIDocumentContext context)
    {
        if (_activeTable == null) return DxpDisposable.Empty;
        var horizontal = cell.TableCellProperties?.GetFirstChild<HorizontalMerge>();
        var physicalContinuation = false;
        if (horizontal != null &&
            (horizontal.Val == null || horizontal.Val.Value == MergedCellValues.Continue))
        {
            if (_activeRow!.PendingHorizontalContinuations <= 0)
                throw new InvalidDataException("A DOCX horizontal merge has an orphaned continuation cell.");
            physicalContinuation = _activeRow.PreserveHorizontalMergeCells;
            _activeRow.PendingHorizontalContinuations--;
            if (!physicalContinuation)
            {
                _suppressHorizontalContinuation = true;
                return DxpDisposable.Create(() => _suppressHorizontalContinuation = false);
            }
            if (_activeRow.PendingHorizontalContinuations == 0)
                _activeRow.PreserveHorizontalMergeCells = false;
        }
        if (!physicalContinuation && _activeRow!.PendingHorizontalContinuations != 0)
            throw new InvalidDataException("A DOCX horizontal merge is interrupted.");
        var story = _activeTable.Story;
        var start = story.Length;
        var borders = cell.TableCellProperties?.GetFirstChild<TableCellBorders>();
        DocParagraphBorder? ReadCellBorder(BorderType? border) => ReadBorder(border);
        DocParagraphBorder? LeftBorderOf(TableCellBorders? source) =>
            ReadCellBorder(source?.GetFirstChild<StartBorder>()) ??
            ReadCellBorder(source?.GetFirstChild<LeftBorder>());
        DocParagraphBorder? RightBorderOf(TableCellBorders? source) =>
            ReadCellBorder(source?.GetFirstChild<EndBorder>()) ??
            ReadCellBorder(source?.GetFirstChild<RightBorder>());
        var table = _activeTable;
        var tableBorders = table.Borders;
        var rowBorders = _activeRow!.TableBorders;
        var rowIndex = table.RowIndex;
        var cellIndex = _activeRow!.CellCount;
        var span = cell.TableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value ??
            cellContext.ColSpan;
        if (horizontal?.Val?.Value == MergedCellValues.Restart)
        {
            _activeRow.UsesHorizontalMerge = true;
            if (span != 1)
                throw new NotSupportedException("A DOCX cell cannot combine hMerge and gridSpan.");
            foreach (var next in cell.ElementsAfter().OfType<TableCell>())
            {
                var nextMerge = next.TableCellProperties?.GetFirstChild<HorizontalMerge>();
                if (nextMerge == null || nextMerge.Val?.Value == MergedCellValues.Restart)
                    break;
                if (nextMerge.Val != null && nextMerge.Val.Value != MergedCellValues.Continue)
                    throw new NotSupportedException("A DOCX horizontal merge has an unsupported value.");
                span++;
            }
            _activeRow.PendingHorizontalContinuations = span - 1;
            _activeRow.PreserveHorizontalMergeCells = cell.ElementsAfter()
                .OfType<TableCell>().Take(span - 1)
                .Any(next => !string.IsNullOrEmpty(next.InnerText));
            if (_activeRow.PreserveHorizontalMergeCells) span = 1;
        }
        if (span is < 1 or > 63)
            throw new NotSupportedException("A DOC table cell has an invalid grid span.");
        var previousCellContentWidth = _positionalCellContentWidth;
        var cellWidth = table.GridWidths.Count > 0
            ? table.GridWidths.Skip(cellContext.ColumnIndex).Take(span).Sum(x => (int)x)
            : int.TryParse(cell.TableCellProperties?
                .GetFirstChild<TableCellWidth>()?.Width?.Value, out var suppliedCellWidth)
                ? suppliedCellWidth : 1440;
        var marginsForTab = MergeCellMargins(MergeCellMargins(table.DefaultMargins,
            ReadCellMargins((cell.Parent as TableRow)?.TablePropertyExceptions?
                .GetFirstChild<TableCellMarginDefault>())),
            ReadCellMargins(cell.TableCellProperties?.GetFirstChild<TableCellMargin>()));
        _positionalCellContentWidth = cellWidth - (marginsForTab?.Left ?? 0) -
            (marginsForTab?.Right ?? 0);
        var isLastCell = !cell.ElementsAfter().OfType<TableCell>().Any(next =>
        {
            var merge = next.TableCellProperties?.GetFirstChild<HorizontalMerge>();
            return merge == null || merge.Val?.Value == MergedCellValues.Restart;
        });
        // Word counts hMerge continuation cells in the band phase, but counts
        // a gridSpan as one visible cell.
        var bandColumn = _activeRow.UsesHorizontalMerge
            ? cellContext.ColumnIndex : cellIndex;
        var previousConditionalRunFormatting = _activeConditionalTableRunFormatting;
        var previousConditionalParagraphFormatting =
            _activeConditionalTableParagraphFormatting;
        var previousCellHasConditionalStyle = _activeCellHasConditionalStyle;
        _activeCellHasConditionalStyle = cell.TableCellProperties?.ChildElements
            .Any(x => x.LocalName == "cnfStyle") == true;
        var horizontalBandFormatting = table.HorizontalBandSize == 0 ||
            rowIndex < table.HorizontalBandOffset ? null :
            ((rowIndex - table.HorizontalBandOffset) / table.HorizontalBandSize) % 2 == 0
                ? table.Band1HorizontalRunFormatting
                : table.Band2HorizontalRunFormatting;
        var verticalBandFormatting = table.VerticalBandSize == 0 ||
            bandColumn < table.VerticalBandOffset ? null :
            ((bandColumn - table.VerticalBandOffset) /
                table.VerticalBandSize) % 2 == 0
                ? table.Band1VerticalRunFormatting
                : table.Band2VerticalRunFormatting;
        var cornerRunFormatting = rowIndex == 0
            ? cellIndex == 0 ? table.NorthWestRunFormatting
                : isLastCell ? table.NorthEastRunFormatting : null
            : rowIndex == table.RowCount - 1
                ? cellIndex == 0 ? table.SouthWestRunFormatting
                    : isLastCell ? table.SouthEastRunFormatting : null
                : null;
        DocCharacterFormatting? combinedRunFormatting = null;
        void ApplyConditionalRun(DocCharacterFormatting? higherPriority)
        {
            if (higherPriority == null) return;
            combinedRunFormatting = combinedRunFormatting == null
                ? higherPriority
                : MergeCharacterDefaults(higherPriority, combinedRunFormatting);
        }
        ApplyConditionalRun(horizontalBandFormatting);
        ApplyConditionalRun(verticalBandFormatting);
        ApplyConditionalRun(cellIndex == 0
            ? table.FirstColumnRunFormatting
            : isLastCell ? table.LastColumnRunFormatting : null);
        ApplyConditionalRun(rowIndex == 0
            ? table.FirstRowRunFormatting
            : rowIndex == table.RowCount - 1 ? table.LastRowRunFormatting : null);
        ApplyConditionalRun(cornerRunFormatting);
        _activeConditionalTableRunFormatting = combinedRunFormatting;
        DocParagraphFormatting? combinedParagraphFormatting = null;
        void ApplyConditionalParagraph(TableStyleOverrideValues condition)
        {
            if (!table.ConditionalParagraphFormatting.TryGetValue(condition,
                out var higherPriority)) return;
            combinedParagraphFormatting = combinedParagraphFormatting == null
                ? higherPriority
                : MergeParagraphDefaults(higherPriority, combinedParagraphFormatting);
        }
        if (table.HorizontalBandSize > 0 && rowIndex >= table.HorizontalBandOffset)
            ApplyConditionalParagraph(((rowIndex - table.HorizontalBandOffset) /
                table.HorizontalBandSize) % 2 == 0
                ? TableStyleOverrideValues.Band1Horizontal
                : TableStyleOverrideValues.Band2Horizontal);
        if (table.VerticalBandSize > 0 && bandColumn >= table.VerticalBandOffset)
            ApplyConditionalParagraph(((bandColumn - table.VerticalBandOffset) /
                table.VerticalBandSize) % 2 == 0
                ? TableStyleOverrideValues.Band1Vertical
                : TableStyleOverrideValues.Band2Vertical);
        if (cellIndex == 0)
            ApplyConditionalParagraph(TableStyleOverrideValues.FirstColumn);
        else if (isLastCell)
            ApplyConditionalParagraph(TableStyleOverrideValues.LastColumn);
        if (rowIndex == 0)
            ApplyConditionalParagraph(TableStyleOverrideValues.FirstRow);
        else if (rowIndex == table.RowCount - 1)
            ApplyConditionalParagraph(TableStyleOverrideValues.LastRow);
        if (rowIndex == 0 && cellIndex == 0)
            ApplyConditionalParagraph(TableStyleOverrideValues.NorthWestCell);
        else if (rowIndex == 0 && isLastCell)
            ApplyConditionalParagraph(TableStyleOverrideValues.NorthEastCell);
        else if (rowIndex == table.RowCount - 1 && cellIndex == 0)
            ApplyConditionalParagraph(TableStyleOverrideValues.SouthWestCell);
        else if (rowIndex == table.RowCount - 1 && isLastCell)
            ApplyConditionalParagraph(TableStyleOverrideValues.SouthEastCell);
        _activeConditionalTableParagraphFormatting = combinedParagraphFormatting;
        var cornerShading = rowIndex == 0
            ? cellIndex == 0 ? table.NorthWestShading
                : isLastCell ? table.NorthEastShading : null
            : rowIndex == table.RowCount - 1
                ? cellIndex == 0 ? table.SouthWestShading
                    : isLastCell ? table.SouthEastShading : null
                : null;
        var directShading = cell.TableCellProperties?.GetFirstChild<Shading>();
        var shading = (directShading?.Val?.Value == ShadingPatternValues.Nil
            ? null : directShading) ??
            cornerShading ??
            (rowIndex == 0 ? table.FirstRowShading : null) ??
            (rowIndex == table.RowCount - 1 ? table.LastRowShading : null) ??
            (cellIndex == 0 ? table.FirstColumnShading : null) ??
            (isLastCell ? table.LastColumnShading : null) ??
            (table.VerticalBandSize == 0 ||
                bandColumn < table.VerticalBandOffset ? null :
                ((bandColumn - table.VerticalBandOffset) /
                    table.VerticalBandSize) % 2 == 0
                    ? table.Band1VerticalShading
                    : table.Band2VerticalShading) ??
            (table.HorizontalBandSize == 0 ||
                rowIndex < table.HorizontalBandOffset ? null :
                ((rowIndex - table.HorizontalBandOffset) /
                    table.HorizontalBandSize) % 2 == 0
                    ? table.Band1HorizontalShading
                    : table.Band2HorizontalShading) ??
            table.DefaultShading;
        var conditionalBorders = rowIndex == 0 ? table.FirstRowBorders : null;
        var lastRowBorders = rowIndex == table.RowCount - 1
            ? table.LastRowBorders : null;
        var columnBorders = cellIndex == 0
            ? table.FirstColumnBorders : isLastCell ? table.LastColumnBorders : null;
        var horizontalBandBorders = table.HorizontalBandSize == 0 ||
            rowIndex < table.HorizontalBandOffset ? null :
            ((rowIndex - table.HorizontalBandOffset) / table.HorizontalBandSize) % 2 == 0
                ? table.Band1HorizontalBorders : table.Band2HorizontalBorders;
        var verticalBandBorders = table.VerticalBandSize == 0 ||
            bandColumn < table.VerticalBandOffset ? null :
            ((bandColumn - table.VerticalBandOffset) / table.VerticalBandSize) % 2 == 0
                ? table.Band1VerticalBorders : table.Band2VerticalBorders;
        var cornerBorders = rowIndex == 0
            ? cellIndex == 0 ? table.NorthWestBorders
                : isLastCell ? table.NorthEastBorders : null
            : rowIndex == table.RowCount - 1
                ? cellIndex == 0 ? table.SouthWestBorders
                    : isLastCell ? table.SouthEastBorders : null
                : null;
        var top = ReadCellBorder(borders?.GetFirstChild<TopBorder>()) ??
            ReadCellBorder(cornerBorders?.GetFirstChild<TopBorder>()) ??
            ReadCellBorder(conditionalBorders?.GetFirstChild<TopBorder>()) ??
            ReadCellBorder(lastRowBorders?.GetFirstChild<TopBorder>()) ??
            // A column-style top edge belongs to the first table row.
            (rowIndex == 0
                ? ReadCellBorder(columnBorders?.GetFirstChild<TopBorder>())
                : null) ??
            ReadCellBorder(verticalBandBorders?.GetFirstChild<TopBorder>()) ??
            ReadCellBorder(horizontalBandBorders?.GetFirstChild<TopBorder>()) ??
            rowBorders?.Top ?? (rowIndex == 0 ? tableBorders?.Top :
                rowBorders?.InsideHorizontal ?? tableBorders?.InsideHorizontal);
        var bottom = ReadCellBorder(borders?.GetFirstChild<BottomBorder>()) ??
            ReadCellBorder(cornerBorders?.GetFirstChild<BottomBorder>()) ??
            ReadCellBorder(conditionalBorders?.GetFirstChild<BottomBorder>()) ??
            ReadCellBorder(lastRowBorders?.GetFirstChild<BottomBorder>()) ??
            // A column-style bottom edge is a table boundary, not an
            // internal border repeated after every row.
            (rowIndex == table.RowCount - 1
                ? ReadCellBorder(columnBorders?.GetFirstChild<BottomBorder>())
                : null) ??
            ReadCellBorder(verticalBandBorders?.GetFirstChild<BottomBorder>()) ??
            ReadCellBorder(horizontalBandBorders?.GetFirstChild<BottomBorder>()) ??
            rowBorders?.Bottom ?? (rowIndex == table.RowCount - 1 ?
                tableBorders?.Bottom : rowBorders?.InsideHorizontal ??
                tableBorders?.InsideHorizontal);
        // Word uses the logical edge when both forms are present.
        var left = LeftBorderOf(borders) ??
            LeftBorderOf(cornerBorders) ??
            (cellIndex == 0 ? LeftBorderOf(conditionalBorders) : null) ??
            (cellIndex == 0 ? LeftBorderOf(lastRowBorders) : null) ??
            LeftBorderOf(columnBorders) ??
            LeftBorderOf(verticalBandBorders) ??
            LeftBorderOf(horizontalBandBorders) ??
            (cellIndex == 0 ? rowBorders?.Left ?? tableBorders?.Left :
                rowBorders?.InsideVertical ?? tableBorders?.InsideVertical);
        var right = RightBorderOf(borders) ??
            RightBorderOf(cornerBorders) ??
            (isLastCell ? RightBorderOf(conditionalBorders) : null) ??
            (isLastCell ? RightBorderOf(lastRowBorders) : null) ??
            RightBorderOf(columnBorders) ??
            RightBorderOf(verticalBandBorders) ??
            RightBorderOf(horizontalBandBorders) ??
            (isLastCell ? rowBorders?.Right ?? tableBorders?.Right :
                rowBorders?.InsideVertical ?? tableBorders?.InsideVertical);
        var topLeftToBottomRight = ReadBorder(borders?
            .GetFirstChild<TopLeftToBottomRightCellBorder>()) ??
            ReadBorder(cornerBorders?
                .GetFirstChild<TopLeftToBottomRightCellBorder>()) ??
            ReadBorder(conditionalBorders?
                .GetFirstChild<TopLeftToBottomRightCellBorder>()) ??
            ReadBorder(lastRowBorders?
                .GetFirstChild<TopLeftToBottomRightCellBorder>()) ??
            ReadBorder(columnBorders?
                .GetFirstChild<TopLeftToBottomRightCellBorder>()) ??
            ReadBorder(verticalBandBorders?
                .GetFirstChild<TopLeftToBottomRightCellBorder>()) ??
            ReadBorder(horizontalBandBorders?
                .GetFirstChild<TopLeftToBottomRightCellBorder>());
        var topRightToBottomLeft = ReadBorder(borders?
            .GetFirstChild<TopRightToBottomLeftCellBorder>()) ??
            ReadBorder(cornerBorders?
                .GetFirstChild<TopRightToBottomLeftCellBorder>()) ??
            ReadBorder(conditionalBorders?
                .GetFirstChild<TopRightToBottomLeftCellBorder>()) ??
            ReadBorder(lastRowBorders?
                .GetFirstChild<TopRightToBottomLeftCellBorder>()) ??
            ReadBorder(columnBorders?
                .GetFirstChild<TopRightToBottomLeftCellBorder>()) ??
            ReadBorder(verticalBandBorders?
                .GetFirstChild<TopRightToBottomLeftCellBorder>()) ??
            ReadBorder(horizontalBandBorders?
                .GetFirstChild<TopRightToBottomLeftCellBorder>());
        // Word draws the reusable first-column rule through the full styled cell edge.
        // Flattening that rule into tc borders shortens it in the DOC render.
        var styleLeftOnly = table.StyleIndex != null &&
            cellIndex == 0 && (!isLastCell || LeftBorderOf(table.LastColumnBorders) == null) &&
            LeftBorderOf(columnBorders) != null && LeftBorderOf(borders) == null &&
            LeftBorderOf(cornerBorders) == null &&
            LeftBorderOf(conditionalBorders) == null &&
            LeftBorderOf(lastRowBorders) == null &&
            LeftBorderOf(verticalBandBorders) == null &&
            LeftBorderOf(horizontalBandBorders) == null;
        var directLeft = styleLeftOnly ? null : left;
        var cellBorders = top == null && bottom == null && directLeft == null && right == null &&
            topLeftToBottomRight == null && topRightToBottomLeft == null
            ? null : new DocCellBorders(top, directLeft, bottom, right,
                topLeftToBottomRight, topRightToBottomLeft);
        var alignment = cell.TableCellProperties?
            .GetFirstChild<TableCellVerticalAlignment>()?.Val?.Value ??
            (rowIndex == 0 ? table.FirstRowVerticalAlignment : null) ??
            (rowIndex == table.RowCount - 1 ? table.LastRowVerticalAlignment : null) ??
            (table.HorizontalBandSize > 0 && rowIndex >= table.HorizontalBandOffset
                ? ((rowIndex - table.HorizontalBandOffset) /
                    table.HorizontalBandSize) % 2 == 0
                    ? table.Band1HorizontalVerticalAlignment
                    : table.Band2HorizontalVerticalAlignment : null) ??
            (table.VerticalBandSize > 0 ?
                table.Band2VerticalVerticalAlignment ??
                table.Band1VerticalVerticalAlignment : null) ??
            (rowIndex == 0 ? table.NorthEastVerticalAlignment ??
                table.NorthWestVerticalAlignment : null) ??
            (rowIndex == table.RowCount - 1 ? table.SouthEastVerticalAlignment ??
                table.SouthWestVerticalAlignment : null) ??
            table.LastColumnVerticalAlignment ??
            table.FirstColumnVerticalAlignment ??
            table.DefaultVerticalAlignment;
        byte? verticalAlignment = alignment == TableVerticalAlignmentValues.Top ? (byte)0 :
            alignment == TableVerticalAlignmentValues.Center ? (byte)1 :
            alignment == TableVerticalAlignmentValues.Bottom ? (byte)2 : null;
        var direction = cell.TableCellProperties?.GetFirstChild<TextDirection>()?.Val?.Value;
        ushort? textFlow = direction == TextDirectionValues.LefToRightTopToBottom ||
            direction == TextDirectionValues.LeftToRightTopToBottom2010
            ? (ushort)0 :
            direction == TextDirectionValues.TopToBottomRightToLeft ||
            direction == TextDirectionValues.TopToBottomRightToLeft2010 ? (ushort)1 :
            direction == TextDirectionValues.BottomToTopLeftToRight ||
            direction == TextDirectionValues.BottomToTopLeftToRight2010 ? (ushort)3 :
            direction == TextDirectionValues.LefttoRightTopToBottomRotated ||
            direction == TextDirectionValues.LeftToRightTopToBottomRotated2010
                ? (ushort)4 :
            direction == TextDirectionValues.TopToBottomRightToLeftRotated ||
            direction == TextDirectionValues.TopToBottomRightToLeftRotated2010
                ? (ushort)5 :
            null;
        if (direction != null && textFlow == null)
            throw new NotSupportedException("The DOCX table cell text direction is unsupported.");
        var hideMarkProperty = cell.TableCellProperties?.GetFirstChild<HideMark>();
        bool? hideMark = hideMarkProperty == null ? null :
            hideMarkProperty.Val?.Value != OnOffOnlyValues.Off;
        var merge = cell.TableCellProperties?.GetFirstChild<VerticalMerge>();
        byte? verticalMerge = merge == null ? null :
            merge.Val?.Value == MergedCellValues.Restart ? (byte)3 :
            merge.Val?.Value == null || merge.Val.Value == MergedCellValues.Continue
                ? (byte)1 :
            throw new NotSupportedException("The DOCX vertical merge value is unsupported.");
        if (horizontal != null && !physicalContinuation &&
            horizontal.Val?.Value != MergedCellValues.Restart)
            throw new NotSupportedException("The DOCX horizontal merge value is unsupported.");
        DocCellShading? cellShading = null;
        if (shading != null)
        {
            if (shading.Val?.Value == ShadingPatternValues.Clear &&
                ReadShadingFill(shading) == null &&
                ReadShadingForeground(shading) == null)
                cellShading = new DocCellShading(null, null, ushort.MaxValue);
            else if (shading.Val?.Value == ShadingPatternValues.Nil)
                cellShading = null;
            else
            {
                var pattern = DocShadingPatterns.ToDoc(shading.Val?.Value ??
                    ShadingPatternValues.Clear);
                if (pattern == null)
                    throw new NotSupportedException("The DOC table cell shading pattern is unsupported.");
                cellShading = new DocCellShading(ReadShadingFill(shading),
                    ReadShadingForeground(shading), pattern.Value);
            }
        }
        return DxpDisposable.Create(() =>
        {
            _positionalCellContentWidth = previousCellContentWidth;
            _activeConditionalTableRunFormatting = previousConditionalRunFormatting;
            _activeConditionalTableParagraphFormatting =
                previousConditionalParagraphFormatting;
            _activeCellHasConditionalStyle = previousCellHasConditionalStyle;
            if (story.Length <= start || story[story.Length - 1] != '\r')
                throw new InvalidDataException("A DOC table cell has no ending paragraph.");
            if (_tableStack.Count == 0)
                story[story.Length - 1] = '\u0007';
            else
            {
                var paragraphRuns = _paragraphStyles[story];
                var lastRun = paragraphRuns.FindLastIndex(x =>
                    x.Start <= story.Length - 1 && story.Length - 1 < x.End);
                if (lastRun < 0)
                    throw new InvalidDataException("A nested DOC cell has no paragraph formatting.");
                paragraphRuns[lastRun] = paragraphRuns[lastRun] with
                {
                    Formatting = paragraphRuns[lastRun].Formatting with
                    {
                        InnerTableCell = true,
                        TableDepth = _tableStack.Count + 1
                    }
                };
            }
            _activeRow!.CellCount++;
            _activeRow.Shadings.Add(cellShading);
            _activeRow.VerticalAlignments.Add(verticalAlignment);
            _activeRow.TextFlows.Add(textFlow);
            _activeRow.HideMarks.Add(hideMark);
            _activeRow.VerticalMerges.Add(verticalMerge);
            _activeRow.HorizontalMerges.Add(physicalContinuation ? (byte)1 :
                horizontal?.Val?.Value == MergedCellValues.Restart &&
                _activeRow.PreserveHorizontalMergeCells ? (byte)2 : null);
            _activeRow.Borders.Add(cellBorders);
            var cellWidth = cell.TableCellProperties?.GetFirstChild<TableCellWidth>();
            var cellWidthType = cellWidth?.Type?.Value;
            DocTablePreferredWidth? preferredCellWidth = null;
            if (cellWidthType == TableWidthUnitValues.Auto ||
                cellWidthType == null && table.AutoFit)
                preferredCellWidth = new DocTablePreferredWidth(1, 0);
            else if (cellWidthType == TableWidthUnitValues.Pct ||
                cellWidthType == TableWidthUnitValues.Dxa)
            {
                if (!ushort.TryParse(cellWidth?.Width?.Value, out var parsedPreferredWidth))
                    throw new NotSupportedException("A DOCX cell has an invalid preferred width.");
                preferredCellWidth = new DocTablePreferredWidth(
                    cellWidthType == TableWidthUnitValues.Pct ? (byte)2 : (byte)3,
                    parsedPreferredWidth);
            }
            _activeRow.PreferredCellWidths.Add(preferredCellWidth);
            var noWrap = cell.TableCellProperties?.GetFirstChild<NoWrap>();
            var styleNoWrap =
                (rowIndex == 0 ? table.FirstRowNoWrap : null) ??
                (rowIndex == 0
                    ? table.NorthEastNoWrap ?? table.NorthWestNoWrap : null) ??
                (rowIndex == table.RowCount - 1
                    ? table.SouthEastNoWrap ?? table.SouthWestNoWrap : null) ??
                (rowIndex == table.RowCount - 1 ? table.LastRowNoWrap : null) ??
                (table.HorizontalBandSize > 0 &&
                    rowIndex >= table.HorizontalBandOffset
                    ? ((rowIndex - table.HorizontalBandOffset) /
                        table.HorizontalBandSize) % 2 == 0
                        ? table.Band1HorizontalNoWrap
                        : table.Band2HorizontalNoWrap : null) ??
                (table.VerticalBandSize > 0
                    ? table.Band2VerticalNoWrap ?? table.Band1VerticalNoWrap
                    : null) ?? table.LastColumnNoWrap ??
                table.FirstColumnNoWrap ?? table.DefaultNoWrap;
            _activeRow.NoWraps.Add(noWrap != null
                ? noWrap.Val == null || noWrap.Val.Value == OnOffOnlyValues.On
                : styleNoWrap);
            var fitText = cell.TableCellProperties?.GetFirstChild<TableCellFitText>();
            _activeRow.FitTexts.Add(fitText == null ? null :
                fitText.Val == null || fitText.Val.Value == OnOffOnlyValues.On);
            _activeRow.CellMargins.Add(ReadCellMargins(
                cell.TableCellProperties?.GetFirstChild<TableCellMargin>()));
            short width;
            if (table.GridWidths.Count > 0)
            {
                if (cellContext.ColumnIndex < 0 ||
                    cellContext.ColumnIndex + span > table.GridWidths.Count)
                    throw new InvalidDataException("A DOCX table cell extends beyond its grid.");
                width = checked((short)table.GridWidths.Skip(cellContext.ColumnIndex)
                    .Take(span).Sum(x => (int)x));
            }
            else width = short.TryParse(cell.TableCellProperties?
                .GetFirstChild<TableCellWidth>()?.Width?.Value, out var suppliedWidth) &&
                suppliedWidth > 0 ? suppliedWidth : (short)1440;
            _activeRow.CellWidths.Add(width);
        });
    }

    private static DocCharacterFormatting MergeCharacterDefaults(
        DocCharacterFormatting own, DocCharacterFormatting defaults) => own with
    {
        Bold = own.Bold ?? defaults.Bold,
        Italic = own.Italic ?? defaults.Italic,
        ComplexScriptBold = own.ComplexScriptBold ?? defaults.ComplexScriptBold,
        ComplexScriptItalic = own.ComplexScriptItalic ?? defaults.ComplexScriptItalic,
        ComplexScriptSizeHalfPoints = own.ComplexScriptSizeHalfPoints ??
            defaults.ComplexScriptSizeHalfPoints,
        RightToLeftText = own.RightToLeftText ?? defaults.RightToLeftText,
        ForceComplexScript = own.ForceComplexScript ?? defaults.ForceComplexScript,
        SizeHalfPoints = own.SizeHalfPoints ?? defaults.SizeHalfPoints,
        Strike = own.Strike ?? defaults.Strike,
        DoubleStrike = own.DoubleStrike ?? defaults.DoubleStrike,
        CharacterScalePercent = own.CharacterScalePercent ??
            defaults.CharacterScalePercent,
        SnapToGrid = own.SnapToGrid ?? defaults.SnapToGrid,
        BaselineOffsetHalfPoints = own.BaselineOffsetHalfPoints ??
            defaults.BaselineOffsetHalfPoints,
        Shadow = own.Shadow ?? defaults.Shadow,
        Outline = own.Outline ?? defaults.Outline,
        Emboss = own.Emboss ?? defaults.Emboss,
        Imprint = own.Imprint ?? defaults.Imprint,
        Caps = own.Caps ?? defaults.Caps,
        SmallCaps = own.SmallCaps ?? defaults.SmallCaps,
        Hidden = own.Hidden ?? defaults.Hidden,
        UnderlineCode = own.UnderlineCode ?? defaults.UnderlineCode,
        UnderlineColorRef = own.UnderlineColorRef ?? defaults.UnderlineColorRef,
        Shading = own.Shading ?? defaults.Shading,
        Border = own.Border ?? defaults.Border,
        ColorRef = own.ColorRef ?? defaults.ColorRef,
        HighlightCode = own.HighlightCode ?? defaults.HighlightCode,
        ScriptCode = own.ScriptCode ?? defaults.ScriptCode,
        EmphasisMarkCode = own.EmphasisMarkCode ?? defaults.EmphasisMarkCode,
        FitText = own.FitText ?? defaults.FitText,
        AsciiFontName = own.AsciiFontName ?? defaults.AsciiFontName,
        EastAsiaFontName = own.EastAsiaFontName ?? defaults.EastAsiaFontName,
        HighAnsiFontName = own.HighAnsiFontName ?? defaults.HighAnsiFontName,
        ComplexScriptFontName = own.ComplexScriptFontName ?? defaults.ComplexScriptFontName,
        CharacterSpacingTwips = own.CharacterSpacingTwips ?? defaults.CharacterSpacingTwips,
        KerningThresholdHalfPoints = own.KerningThresholdHalfPoints ??
            defaults.KerningThresholdHalfPoints,
        LanguageId = own.LanguageId ?? defaults.LanguageId,
        EastAsiaLanguageId = own.EastAsiaLanguageId ?? defaults.EastAsiaLanguageId,
        ComplexScriptLanguageId = own.ComplexScriptLanguageId ?? defaults.ComplexScriptLanguageId
    };

    private static DocCharacterFormatting ExcludeCharacterStyleProperties(
        DocCharacterFormatting table, DocCharacterFormatting characterStyle) => table with
    {
        Bold = characterStyle.Bold == null ? table.Bold : null,
        Italic = characterStyle.Italic == null ? table.Italic : null,
        ComplexScriptBold = characterStyle.ComplexScriptBold == null
            ? table.ComplexScriptBold : null,
        ComplexScriptItalic = characterStyle.ComplexScriptItalic == null
            ? table.ComplexScriptItalic : null,
        ComplexScriptSizeHalfPoints = characterStyle.ComplexScriptSizeHalfPoints == null
            ? table.ComplexScriptSizeHalfPoints : null,
        RightToLeftText = characterStyle.RightToLeftText == null ? table.RightToLeftText : null,
        ForceComplexScript = characterStyle.ForceComplexScript == null
            ? table.ForceComplexScript : null,
        SizeHalfPoints = characterStyle.SizeHalfPoints == null ? table.SizeHalfPoints : null,
        Strike = characterStyle.Strike == null ? table.Strike : null,
        DoubleStrike = characterStyle.DoubleStrike == null ? table.DoubleStrike : null,
        CharacterScalePercent = characterStyle.CharacterScalePercent == null
            ? table.CharacterScalePercent : null,
        SnapToGrid = characterStyle.SnapToGrid == null ? table.SnapToGrid : null,
        BaselineOffsetHalfPoints = characterStyle.BaselineOffsetHalfPoints == null
            ? table.BaselineOffsetHalfPoints : null,
        Shadow = characterStyle.Shadow == null ? table.Shadow : null,
        Outline = characterStyle.Outline == null ? table.Outline : null,
        Emboss = characterStyle.Emboss == null ? table.Emboss : null,
        Imprint = characterStyle.Imprint == null ? table.Imprint : null,
        Caps = characterStyle.Caps == null ? table.Caps : null,
        SmallCaps = characterStyle.SmallCaps == null ? table.SmallCaps : null,
        Hidden = characterStyle.Hidden == null ? table.Hidden : null,
        UnderlineCode = characterStyle.UnderlineCode == null ? table.UnderlineCode : null,
        UnderlineColorRef = characterStyle.UnderlineColorRef == null
            ? table.UnderlineColorRef : null,
        Shading = characterStyle.Shading == null ? table.Shading : null,
        Border = characterStyle.Border == null ? table.Border : null,
        ColorRef = characterStyle.ColorRef == null ? table.ColorRef : null,
        HighlightCode = characterStyle.HighlightCode == null ? table.HighlightCode : null,
        ScriptCode = characterStyle.ScriptCode == null ? table.ScriptCode : null,
        EmphasisMarkCode = characterStyle.EmphasisMarkCode == null
            ? table.EmphasisMarkCode : null,
        FitText = characterStyle.FitText == null ? table.FitText : null,
        AsciiFontIndex = characterStyle.AsciiFontIndex == null ? table.AsciiFontIndex : null,
        EastAsiaFontIndex = characterStyle.EastAsiaFontIndex == null
            ? table.EastAsiaFontIndex : null,
        HighAnsiFontIndex = characterStyle.HighAnsiFontIndex == null
            ? table.HighAnsiFontIndex : null,
        ComplexScriptFontIndex = characterStyle.ComplexScriptFontIndex == null
            ? table.ComplexScriptFontIndex : null,
        AsciiFontName = characterStyle.AsciiFontName == null ? table.AsciiFontName : null,
        EastAsiaFontName = characterStyle.EastAsiaFontName == null
            ? table.EastAsiaFontName : null,
        HighAnsiFontName = characterStyle.HighAnsiFontName == null
            ? table.HighAnsiFontName : null,
        ComplexScriptFontName = characterStyle.ComplexScriptFontName == null
            ? table.ComplexScriptFontName : null,
        CharacterSpacingTwips = characterStyle.CharacterSpacingTwips == null
            ? table.CharacterSpacingTwips : null,
        KerningThresholdHalfPoints = characterStyle.KerningThresholdHalfPoints == null
            ? table.KerningThresholdHalfPoints : null,
        LanguageId = characterStyle.LanguageId == null ? table.LanguageId : null,
        EastAsiaLanguageId = characterStyle.EastAsiaLanguageId == null
            ? table.EastAsiaLanguageId : null,
        ComplexScriptLanguageId = characterStyle.ComplexScriptLanguageId == null
            ? table.ComplexScriptLanguageId : null
    };

    private static DocParagraphFormatting MergeParagraphDefaults(
        DocParagraphFormatting own, DocParagraphFormatting defaults) => own with
    {
        Justification = own.Justification ?? defaults.Justification,
        BeforeTwips = own.BeforeTwips ?? defaults.BeforeTwips,
        AfterTwips = own.AfterTwips ?? defaults.AfterTwips,
        LeftTwips = own.LeftTwips ?? defaults.LeftTwips,
        RightTwips = own.RightTwips ?? defaults.RightTwips,
        FirstLineTwips = own.FirstLineTwips ?? defaults.FirstLineTwips,
        KeepLines = own.KeepLines ?? defaults.KeepLines,
        KeepWithNext = own.KeepWithNext ?? defaults.KeepWithNext,
        PageBreakBefore = own.PageBreakBefore ?? defaults.PageBreakBefore,
        WidowControl = own.WidowControl ?? defaults.WidowControl,
        ContextualSpacing = own.ContextualSpacing ?? defaults.ContextualSpacing,
        MirrorIndents = own.MirrorIndents ?? defaults.MirrorIndents,
        SuppressAutoHyphens = own.SuppressAutoHyphens ??
            defaults.SuppressAutoHyphens,
        TextAlignmentCode = own.TextAlignmentCode ?? defaults.TextAlignmentCode,
        SuppressLineNumbers = own.SuppressLineNumbers ?? defaults.SuppressLineNumbers,
        BeforeAutoSpacing = own.BeforeAutoSpacing ?? defaults.BeforeAutoSpacing,
        AfterAutoSpacing = own.AfterAutoSpacing ?? defaults.AfterAutoSpacing,
        BeforeLines = own.BeforeLines ?? defaults.BeforeLines,
        AfterLines = own.AfterLines ?? defaults.AfterLines,
        LeftChars = own.LeftChars ?? defaults.LeftChars,
        RightChars = own.RightChars ?? defaults.RightChars,
        FirstLineChars = own.FirstLineChars ?? defaults.FirstLineChars,
        ParagraphRightToLeft = own.ParagraphRightToLeft ?? defaults.ParagraphRightToLeft,
        OutlineLevel = own.OutlineLevel ?? defaults.OutlineLevel,
        Kinsoku = own.Kinsoku ?? defaults.Kinsoku,
        WordWrap = own.WordWrap ?? defaults.WordWrap,
        SnapToGrid = own.SnapToGrid ?? defaults.SnapToGrid,
        AutoSpaceDE = own.AutoSpaceDE ?? defaults.AutoSpaceDE,
        AutoSpaceDN = own.AutoSpaceDN ?? defaults.AutoSpaceDN,
        AdjustRightIndent = own.AdjustRightIndent ?? defaults.AdjustRightIndent,
        LineValue = own.LineValue ?? defaults.LineValue,
        LineIsMultiple = own.LineIsMultiple ?? defaults.LineIsMultiple,
        ClearedTabPositions = own.ClearedTabPositions is { Count: > 0 }
            ? own.ClearedTabPositions : defaults.ClearedTabPositions,
        TabStops = own.TabStops is { Count: > 0 }
            ? own.TabStops : defaults.TabStops,
        FillRgb = own.FillRgb ?? defaults.FillRgb,
        ShadingForegroundRgb = own.ShadingForegroundRgb ?? defaults.ShadingForegroundRgb,
        ShadingPattern = own.ShadingPattern ?? defaults.ShadingPattern,
        TopBorder = own.TopBorder ?? defaults.TopBorder,
        LeftBorder = own.LeftBorder ?? defaults.LeftBorder,
        BottomBorder = own.BottomBorder ?? defaults.BottomBorder,
        RightBorder = own.RightBorder ?? defaults.RightBorder,
        BetweenBorder = own.BetweenBorder ?? defaults.BetweenBorder
    };

    private static DocParagraphFormatting ExcludeParagraphStyleProperties(
        DocParagraphFormatting table, DocParagraphFormatting style) => table with
    {
        Justification = style.Justification == null ? table.Justification : null,
        BeforeTwips = style.BeforeTwips == null ? table.BeforeTwips : null,
        AfterTwips = style.AfterTwips == null ? table.AfterTwips : null,
        LeftTwips = style.LeftTwips == null ? table.LeftTwips : null,
        RightTwips = style.RightTwips == null ? table.RightTwips : null,
        FirstLineTwips = style.FirstLineTwips == null ? table.FirstLineTwips : null,
        KeepLines = style.KeepLines == null ? table.KeepLines : null,
        KeepWithNext = style.KeepWithNext == null ? table.KeepWithNext : null,
        PageBreakBefore = style.PageBreakBefore == null ? table.PageBreakBefore : null,
        WidowControl = style.WidowControl == null ? table.WidowControl : null,
        ContextualSpacing = style.ContextualSpacing == null
            ? table.ContextualSpacing : null,
        MirrorIndents = style.MirrorIndents == null ? table.MirrorIndents : null,
        SuppressAutoHyphens = style.SuppressAutoHyphens == null
            ? table.SuppressAutoHyphens : null,
        TextAlignmentCode = style.TextAlignmentCode == null
            ? table.TextAlignmentCode : null,
        SuppressLineNumbers = style.SuppressLineNumbers == null
            ? table.SuppressLineNumbers : null,
        BeforeAutoSpacing = style.BeforeAutoSpacing == null
            ? table.BeforeAutoSpacing : null,
        AfterAutoSpacing = style.AfterAutoSpacing == null
            ? table.AfterAutoSpacing : null,
        BeforeLines = style.BeforeLines == null ? table.BeforeLines : null,
        AfterLines = style.AfterLines == null ? table.AfterLines : null,
        LeftChars = style.LeftChars == null ? table.LeftChars : null,
        RightChars = style.RightChars == null ? table.RightChars : null,
        FirstLineChars = style.FirstLineChars == null
            ? table.FirstLineChars : null,
        ParagraphRightToLeft = style.ParagraphRightToLeft == null
            ? table.ParagraphRightToLeft : null,
        OutlineLevel = style.OutlineLevel == null ? table.OutlineLevel : null,
        Kinsoku = style.Kinsoku == null ? table.Kinsoku : null,
        WordWrap = style.WordWrap == null ? table.WordWrap : null,
        SnapToGrid = style.SnapToGrid == null ? table.SnapToGrid : null,
        AutoSpaceDE = style.AutoSpaceDE == null ? table.AutoSpaceDE : null,
        AutoSpaceDN = style.AutoSpaceDN == null ? table.AutoSpaceDN : null,
        AdjustRightIndent = style.AdjustRightIndent == null
            ? table.AdjustRightIndent : null,
        LineValue = style.LineValue == null ? table.LineValue : null,
        LineIsMultiple = style.LineIsMultiple == null ? table.LineIsMultiple : null,
        ClearedTabPositions = style.ClearedTabPositions == null
            ? table.ClearedTabPositions : null,
        TabStops = style.TabStops == null ? table.TabStops : null,
        FillRgb = style.FillRgb == null ? table.FillRgb : null,
        ShadingForegroundRgb = style.ShadingForegroundRgb == null
            ? table.ShadingForegroundRgb : null,
        ShadingPattern = style.ShadingPattern == null ? table.ShadingPattern : null,
        TopBorder = style.TopBorder == null ? table.TopBorder : null,
        LeftBorder = style.LeftBorder == null ? table.LeftBorder : null,
        BottomBorder = style.BottomBorder == null ? table.BottomBorder : null,
        RightBorder = style.RightBorder == null ? table.RightBorder : null,
        BetweenBorder = style.BetweenBorder == null ? table.BetweenBorder : null
    };

    private DocParagraphFormatting ReadParagraphFormatting(DocumentFormat.OpenXml.OpenXmlCompositeElement? properties)
    {
        var alignment = properties?.GetFirstChild<Justification>()?.Val?.Value;
        // DOC sprmPJc values 0 and 2 follow the paragraph's leading and
        // trailing edges even when a named style supplies RTL direction.
        byte? justification = alignment == JustificationValues.Left ||
            alignment == JustificationValues.Start ? (byte)0 :
            alignment == JustificationValues.Center ? (byte)1 :
            alignment == JustificationValues.Right ||
            alignment == JustificationValues.End ? (byte)2 :
            alignment == JustificationValues.Both ? (byte)3 :
            alignment == JustificationValues.Distribute ? (byte)4 : null;
        var spacing = properties?.GetFirstChild<SpacingBetweenLines>();
        bool? beforeAutoSpacing = spacing?.BeforeAutoSpacing?.Value;
        bool? afterAutoSpacing = spacing?.AfterAutoSpacing?.Value;
        short? beforeLines = spacing?.BeforeLines?.Value is int beforeLineValue
            ? checked((short)beforeLineValue) : null;
        short? afterLines = spacing?.AfterLines?.Value is int afterLineValue
            ? checked((short)afterLineValue) : null;
        ushort? before = ushort.TryParse(spacing?.Before?.Value,
            out var beforeValue) ? beforeValue : null;
        ushort? after = ushort.TryParse(spacing?.After?.Value,
            out var afterValue) ? afterValue : null;
        var indent = properties?.GetFirstChild<Indentation>();
        short? left = short.TryParse(indent?.Left?.Value, out var leftValue)
            ? leftValue : null;
        short? right = short.TryParse(indent?.Right?.Value, out var rightValue)
            ? rightValue : null;
        short? firstLine = short.TryParse(indent?.FirstLine?.Value, out var firstValue)
            ? firstValue : null;
        if (short.TryParse(indent?.Hanging?.Value, out var hangingValue))
            firstLine = checked((short)-hangingValue);
        short? leftChars = (indent?.StartCharacters?.Value ??
            indent?.LeftChars?.Value) is int leftCharValue
            ? checked((short)leftCharValue) : null;
        short? rightChars = (indent?.EndCharacters?.Value ??
            indent?.RightChars?.Value) is int rightCharValue
            ? checked((short)rightCharValue) : null;
        short? firstLineChars = indent?.FirstLineChars?.Value is int firstCharValue
            ? checked((short)firstCharValue) : null;
        if (indent?.HangingChars?.Value is int hangingCharValue)
            firstLineChars = checked((short)-hangingCharValue);
        var keepLinesElement = properties?.GetFirstChild<KeepLines>();
        bool? keepLines = keepLinesElement == null ? null :
            keepLinesElement.Val?.Value ?? true;
        var keepNextElement = properties?.GetFirstChild<KeepNext>();
        bool? keepWithNext = keepNextElement == null ? null :
            keepNextElement.Val?.Value ?? true;
        var pageBreakElement = properties?.GetFirstChild<PageBreakBefore>();
        bool? pageBreakBefore = pageBreakElement == null ? null :
            pageBreakElement.Val?.Value ?? true;
        var widowElement = properties?.GetFirstChild<WidowControl>();
        bool? widowControl = widowElement == null ? null :
            widowElement.Val?.Value ?? true;
        var contextualElement = properties?.GetFirstChild<ContextualSpacing>();
        bool? contextualSpacing = contextualElement == null ? null :
            contextualElement.Val?.Value ?? true;
        var mirrorElement = properties?.GetFirstChild<MirrorIndents>();
        bool? mirrorIndents = mirrorElement == null ? null :
            mirrorElement.Val?.Value ?? true;
        var suppressHyphensElement = properties?.GetFirstChild<SuppressAutoHyphens>();
        bool? suppressAutoHyphens = suppressHyphensElement == null ? null :
            suppressHyphensElement.Val?.Value ?? true;
        var suppressLineNumbersElement = properties?.GetFirstChild<SuppressLineNumbers>();
        bool? suppressLineNumbers = suppressLineNumbersElement == null ? null :
            suppressLineNumbersElement.Val?.Value ?? true;
        short? textAlignmentCode = properties?.GetFirstChild<TextAlignment>()?
            .Val?.Value switch
        {
            var value when value == VerticalTextAlignmentValues.Top => 0,
            var value when value == VerticalTextAlignmentValues.Center => 1,
            var value when value == VerticalTextAlignmentValues.Baseline => 2,
            var value when value == VerticalTextAlignmentValues.Bottom => 3,
            var value when value == VerticalTextAlignmentValues.Auto => 4,
            _ => null
        };
        var kinsokuElement = properties?.GetFirstChild<Kinsoku>();
        bool? kinsoku = kinsokuElement == null ? null :
            kinsokuElement.Val?.Value ?? true;
        var wordWrapElement = properties?.GetFirstChild<WordWrap>();
        bool? wordWrap = wordWrapElement == null ? null :
            wordWrapElement.Val?.Value ?? true;
        var snapToGridElement = properties?.GetFirstChild<SnapToGrid>();
        bool? snapToGrid = snapToGridElement == null ? null :
            snapToGridElement.Val?.Value ?? true;
        var autoSpaceDEElement = properties?.GetFirstChild<AutoSpaceDE>();
        bool? autoSpaceDE = autoSpaceDEElement == null ? null :
            autoSpaceDEElement.Val?.Value ?? true;
        var autoSpaceDNElement = properties?.GetFirstChild<AutoSpaceDN>();
        bool? autoSpaceDN = autoSpaceDNElement == null ? null :
            autoSpaceDNElement.Val?.Value ?? true;
        var adjustRightElement = properties?.GetFirstChild<AdjustRightIndent>();
        bool? adjustRightIndent = adjustRightElement == null ? null :
            adjustRightElement.Val?.Value ?? true;
        var outlineValue = properties?.GetFirstChild<OutlineLevel>()?.Val?.Value;
        if (outlineValue is < 0 or > 9)
            throw new NotSupportedException("DOC outline levels must be between zero and nine.");
        byte? outlineLevel = outlineValue is int parsedOutlineLevel
            ? checked((byte)parsedOutlineLevel) : null;
        var bidiElement = properties?.GetFirstChild<BiDi>();
        bool? paragraphRightToLeft = bidiElement == null ? null :
            bidiElement.Val?.Value ?? true;
        short? lineValue = null;
        bool? lineIsMultiple = null;
        if (short.TryParse(spacing?.Line?.Value, out var parsedLine))
        {
            var rule = spacing?.LineRule?.Value;
            if (rule == LineSpacingRuleValues.Exact)
            {
                lineValue = checked((short)-parsedLine);
                lineIsMultiple = false;
            }
            else if (rule == LineSpacingRuleValues.AtLeast)
            {
                lineValue = parsedLine;
                lineIsMultiple = false;
            }
            else if (rule == null || rule == LineSpacingRuleValues.Auto)
            {
                lineValue = parsedLine;
                lineIsMultiple = true;
            }
        }
        var clearedTabs = new List<short>();
        var tabStops = new List<DocTabStop>();
        foreach (var tab in properties?.GetFirstChild<Tabs>()?.Elements<TabStop>() ?? [])
        {
            var rawPosition = tab.Position?.Value;
            if (rawPosition == null || rawPosition < short.MinValue ||
                rawPosition > short.MaxValue) continue;
            var position = checked((short)rawPosition.Value);
            var tabAlignment = tab.Val?.Value;
            if (tabAlignment == TabStopValues.Clear)
            {
                clearedTabs.Add(position);
                continue;
            }
            byte? docAlignment = tabAlignment == TabStopValues.Left ? (byte)0 :
                tabAlignment == TabStopValues.Center ? (byte)1 :
                tabAlignment == TabStopValues.Right ? (byte)2 :
                tabAlignment == TabStopValues.Decimal ? (byte)3 :
                tabAlignment == TabStopValues.Bar ? (byte)4 :
                tabAlignment == TabStopValues.Number ? (byte)6 : null;
            if (docAlignment == null) continue;
            var leader = tab.Leader?.Value;
            byte docLeader = leader == TabStopLeaderCharValues.Dot ? (byte)1 :
                leader == TabStopLeaderCharValues.Hyphen ? (byte)2 :
                leader == TabStopLeaderCharValues.Underscore ? (byte)3 :
                leader == TabStopLeaderCharValues.Heavy ? (byte)4 :
                leader == TabStopLeaderCharValues.MiddleDot ? (byte)5 : (byte)0;
            tabStops.Add(new DocTabStop(position, docAlignment.Value, docLeader));
        }
        var shading = properties?.GetFirstChild<Shading>();
        var pattern = shading?.Val?.Value == ShadingPatternValues.Nil
            ? ushort.MaxValue : shading == null ? null : DocShadingPatterns.ToDoc(
                shading.Val?.Value ?? ShadingPatternValues.Clear);
        var fillRgb = pattern == null ? null : ReadShadingFill(shading);
        var foregroundRgb = pattern == null ? null : ReadShadingForeground(shading);
        var borders = properties?.GetFirstChild<ParagraphBorders>();
        short? listOverride = null;
        byte? listLevel = null;
        var numbering = properties?.GetFirstChild<NumberingProperties>();
        if (numbering?.NumberingId?.Val?.Value is int numberId && numberId != 0)
        {
            if (!_listNumberIds.TryGetValue(numberId, out var index))
                throw new InvalidDataException($"DOCX list {numberId} has no numbering instance.");
            listOverride = index;
            var level = numbering.NumberingLevelReference?.Val?.Value ?? 0;
            if (level is < 0 or > 8)
                throw new NotSupportedException("DOC list levels must be between zero and eight.");
            listLevel = checked((byte)level);
        }
        return new DocParagraphFormatting(justification, before, after, left, right,
            firstLine, keepLines, keepWithNext, pageBreakBefore, lineValue,
            lineIsMultiple, clearedTabs, tabStops, fillRgb, foregroundRgb, pattern,
            ReadBorder(borders?.GetFirstChild<TopBorder>()),
            ReadBorder(borders?.GetFirstChild<LeftBorder>()),
            ReadBorder(borders?.GetFirstChild<BottomBorder>()),
            ReadBorder(borders?.GetFirstChild<RightBorder>()),
            ReadBorder(borders?.GetFirstChild<BetweenBorder>()),
            ListOverrideIndex: listOverride, ListLevel: listLevel,
            WidowControl: widowControl, ContextualSpacing: contextualSpacing,
            MirrorIndents: mirrorIndents,
            SuppressAutoHyphens: suppressAutoHyphens,
            TextAlignmentCode: textAlignmentCode,
            SuppressLineNumbers: suppressLineNumbers,
            ParagraphRightToLeft: paragraphRightToLeft,
            OutlineLevel: outlineLevel, Kinsoku: kinsoku, WordWrap: wordWrap,
            SnapToGrid: snapToGrid, AutoSpaceDE: autoSpaceDE,
            AutoSpaceDN: autoSpaceDN, AdjustRightIndent: adjustRightIndent,
            BeforeAutoSpacing: beforeAutoSpacing,
            AfterAutoSpacing: afterAutoSpacing,
            BeforeLines: beforeLines, AfterLines: afterLines,
            LeftChars: leftChars, RightChars: rightChars,
            FirstLineChars: firstLineChars);
    }

    private DocTableBorders? ReadTableBorders(TableBorders? borders)
    {
        if (borders == null) return null;
        var physicalLeft = ReadBorder(borders.LeftBorder);
        var logicalStart = ReadBorder(borders.StartBorder);
        var physicalRight = ReadBorder(borders.RightBorder);
        var logicalEnd = ReadBorder(borders.EndBorder);
        return new DocTableBorders(ReadBorder(borders.TopBorder),
            logicalStart ?? physicalLeft, ReadBorder(borders.BottomBorder),
            logicalEnd ?? physicalRight,
            ReadBorder(borders.InsideHorizontalBorder),
            ReadBorder(borders.InsideVerticalBorder));
    }

    private DocTableBorders? ResolveEffectiveTableBorders(Table table)
    {
        static DocTableBorders? Merge(DocTableBorders? preferred,
            DocTableBorders? inherited) => preferred == null ? inherited :
            inherited == null ? preferred : new DocTableBorders(
                preferred.Top ?? inherited.Top,
                preferred.Left ?? inherited.Left,
                preferred.Bottom ?? inherited.Bottom,
                preferred.Right ?? inherited.Right,
                preferred.InsideHorizontal ?? inherited.InsideHorizontal,
                preferred.InsideVertical ?? inherited.InsideVertical);

        var resolved = ReadTableBorders(table.TableProperties?
            .GetFirstChild<TableBorders>());
        var styleId = table.TableProperties?.GetFirstChild<TableStyle>()?.Val?.Value;
        var styles = _mainPart?.StyleDefinitionsPart?.Styles;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (styleId != null && styles != null && visited.Add(styleId))
        {
            var style = styles.Elements<Style>().FirstOrDefault(x =>
                x.Type?.Value == StyleValues.Table && x.StyleId?.Value == styleId);
            if (style == null) break;
            resolved = Merge(resolved, ReadTableBorders(style.StyleTableProperties?
                .GetFirstChild<TableBorders>()));
            styleId = style.BasedOn?.Val?.Value;
        }
        return resolved;
    }

    private DocParagraphBorder? ReadBorder(BorderType? border)
    {
        var value = DocParagraphBorder.FromOpenXml(border);
        if (value == null) return null;
        var color = ResolveColor(border?.Color?.Value, border?.ThemeColor?.Value,
            border?.ThemeTint?.Value, border?.ThemeShade?.Value);
        return value with { ColorRgb = color == 0xFF000000 ? null : color };
    }


    private static DocCellMargins? ReadCellMargins(OpenXmlElement? element)
    {
        if (element == null) return null;
        static ushort? ReadSide(OpenXmlElement? side)
        {
            if (side == null) return null;
            const string ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var type = side.GetAttribute("type", ns).Value;
            if (type.Length != 0 && type != "dxa")
                throw new NotSupportedException("Only twip table cell margins are supported.");
            if (!ushort.TryParse(side.GetAttribute("w", ns).Value, out var value) ||
                value > 31680)
                throw new InvalidDataException("A DOCX table cell margin is invalid.");
            return value;
        }
        var top = ReadSide(element.GetFirstChild<TopMargin>());
        var physicalLeft = ReadSide((OpenXmlElement?)element.GetFirstChild<LeftMargin>() ??
            element.GetFirstChild<TableCellLeftMargin>());
        var logicalStart = ReadSide(element.GetFirstChild<StartMargin>());
        var bottom = ReadSide(element.GetFirstChild<BottomMargin>());
        var physicalRight = ReadSide((OpenXmlElement?)element.GetFirstChild<RightMargin>() ??
            element.GetFirstChild<TableCellRightMargin>());
        var logicalEnd = ReadSide(element.GetFirstChild<EndMargin>());
        // Word ignores left/right siblings when start/end are present.
        var left = logicalStart ?? physicalLeft;
        var right = logicalEnd ?? physicalRight;
        return top == null && left == null && bottom == null && right == null
            ? null : new DocCellMargins(top, left, bottom, right);
    }

    private DocCellShading? ReadTableBackgroundShading(Shading? shading)
    {
        if (shading == null || shading.Val?.Value == ShadingPatternValues.Nil)
            return null;
        var pattern = DocShadingPatterns.ToDoc(shading.Val?.Value ??
            ShadingPatternValues.Clear);
        if (pattern == null)
            throw new NotSupportedException("The DOC table-background shading pattern is unsupported.");
        return new DocCellShading(ReadShadingFill(shading),
            ReadShadingForeground(shading), pattern.Value);
    }

    private uint? ReadShadingFill(Shading? shading) => ShadingColor(
        shading?.Fill?.Value, shading?.ThemeFill?.Value,
        shading?.ThemeFillTint?.Value, shading?.ThemeFillShade?.Value);

    private uint? ReadShadingForeground(Shading? shading) => ShadingColor(
        shading?.Color?.Value, shading?.ThemeColor?.Value,
        shading?.ThemeTint?.Value, shading?.ThemeShade?.Value);

    private uint? ShadingColor(string? fallback, ThemeColorValues? theme,
        string? tint, string? shade)
    {
        var value = ResolveColor(fallback, theme, tint, shade);
        return value == 0xFF000000 ? null : value;
    }

    public override IDisposable VisitDeletedRunBegin(DeletedRun run,
        DxpIDocumentContext context)
    {
        var previous = _deletedRevisionRun;
        var previousAuthor = _deletedRevisionAuthor;
        var previousAt = _deletedRevisionAt;
        _deletedRevisionRun = true;
        _deletedRevisionAuthor = run.Author?.Value;
        _deletedRevisionAt = run.Date?.Value;
        return DxpDisposable.Create(() =>
        {
            _deletedRevisionRun = previous;
            _deletedRevisionAuthor = previousAuthor;
            _deletedRevisionAt = previousAt;
        });
    }

    public override IDisposable VisitInsertedRunBegin(InsertedRun run,
        DxpIDocumentContext context)
    {
        var previous = _insertedRevisionRun;
        var previousAuthor = _insertedRevisionAuthor;
        var previousAt = _insertedRevisionAt;
        _insertedRevisionRun = true;
        _insertedRevisionAuthor = run.Author?.Value;
        _insertedRevisionAt = run.Date?.Value;
        return DxpDisposable.Create(() =>
        {
            _insertedRevisionRun = previous;
            _insertedRevisionAuthor = previousAuthor;
            _insertedRevisionAt = previousAt;
        });
    }

    public override void VisitDeletedText(DeletedText text, DxpIDocumentContext context) =>
        _paragraphText?.Append(text.Text);

    public override IDisposable VisitRunBegin(Run run, DxpIDocumentContext context)
    {
        var target = _paragraphText;
        if (target == null) return DxpDisposable.Empty;
        var properties = run.RunProperties;
        var runStyleId = properties?.RunStyle?.Val?.Value;
        int? runStyleIndex = runStyleId != null &&
            _styleById.TryGetValue(runStyleId, out var foundStyle) &&
            _styles.Any(x => x.Index == foundStyle && x.Type == 2)
            ? foundStyle : null;
        var formatting = ReadCharacterFormatting(properties, runStyleIndex) with
        {
            DeletedRevision = _deletedRevisionRun ? true : null,
            InsertedRevision = _insertedRevisionRun ? true : null,
            DeletedRevisionAuthor = _deletedRevisionRun ? _deletedRevisionAuthor : null,
            InsertedRevisionAuthor = _insertedRevisionRun ? _insertedRevisionAuthor : null,
            DeletedRevisionAt = _deletedRevisionRun ? _deletedRevisionAt : null,
            InsertedRevisionAt = _insertedRevisionRun ? _insertedRevisionAt : null
        };
        if (runStyleIndex is int characterStyleIndex)
        {
            bool? StyleValue(int index, Func<DocCharacterFormatting, bool?> property)
            {
                var visited = new HashSet<int>();
                while (index != 0 && visited.Add(index))
                {
                    var style = _styles.FirstOrDefault(x => x.Index == index);
                    if (style == null) break;
                    if (property(style.CharacterFormatting) is bool value) return value;
                    index = style.BasedOnIndex ?? 0;
                }
                return null;
            }
            // Word treats a character-style toggle against the paragraph style,
            // while the generated DOC can leave it active when the character
            // style has its own based-on chain. A direct off operand preserves
            // the displayed result without discarding the style relationship.
            if (formatting.Bold == null &&
                StyleValue(characterStyleIndex, x => x.Bold) == true &&
                StyleValue(_currentParagraphStyleIndex, x => x.Bold) == true)
                formatting = formatting with { Bold = false };
            if (formatting.Italic == null &&
                StyleValue(characterStyleIndex, x => x.Italic) == true &&
                StyleValue(_currentParagraphStyleIndex, x => x.Italic) == true)
                formatting = formatting with { Italic = false };
            if (formatting.Strike == null &&
                StyleValue(characterStyleIndex, x => x.Strike) == true &&
                StyleValue(_currentParagraphStyleIndex, x => x.Strike) == true)
                formatting = formatting with { Strike = false };
            if (formatting.DoubleStrike == null &&
                StyleValue(characterStyleIndex, x => x.DoubleStrike) == true &&
                StyleValue(_currentParagraphStyleIndex, x => x.DoubleStrike) == true)
                formatting = formatting with { DoubleStrike = false };
            if (formatting.Shadow == null &&
                StyleValue(characterStyleIndex, x => x.Shadow) == true &&
                StyleValue(_currentParagraphStyleIndex, x => x.Shadow) == true)
                formatting = formatting with { Shadow = false };
            if (formatting.Outline == null &&
                StyleValue(characterStyleIndex, x => x.Outline) == true &&
                StyleValue(_currentParagraphStyleIndex, x => x.Outline) == true)
                formatting = formatting with { Outline = false };
            if (formatting.Emboss == null &&
                StyleValue(characterStyleIndex, x => x.Emboss) == true &&
                StyleValue(_currentParagraphStyleIndex, x => x.Emboss) == true)
                formatting = formatting with { Emboss = false };
            if (formatting.Imprint == null &&
                StyleValue(characterStyleIndex, x => x.Imprint) == true &&
                StyleValue(_currentParagraphStyleIndex, x => x.Imprint) == true)
                formatting = formatting with { Imprint = false };
        }
        if (_activeConditionalTableRunFormatting is { } tableRunFormatting)
        {
            var paragraphStyleFormatting = DocCharacterFormatting.Empty;
            var paragraphVisited = new HashSet<int>();
            var paragraphIndex = _currentParagraphStyleIndex;
            while (paragraphIndex != 0 && paragraphVisited.Add(paragraphIndex))
            {
                var style = _styles.FirstOrDefault(x => x.Index == paragraphIndex);
                if (style == null) break;
                paragraphStyleFormatting = MergeCharacterDefaults(paragraphStyleFormatting,
                    style.CharacterFormatting);
                paragraphIndex = style.BasedOnIndex ?? 0;
            }
            tableRunFormatting = ExcludeCharacterStyleProperties(tableRunFormatting,
                paragraphStyleFormatting);
            if (runStyleIndex is int index)
            {
                var styleFormatting = DocCharacterFormatting.Empty;
                var visited = new HashSet<int>();
                while (index != 0 && visited.Add(index))
                {
                    var style = _styles.FirstOrDefault(x => x.Index == index);
                    if (style == null) break;
                    styleFormatting = MergeCharacterDefaults(styleFormatting,
                        style.CharacterFormatting);
                    index = style.BasedOnIndex ?? 0;
                }
                tableRunFormatting = ExcludeCharacterStyleProperties(tableRunFormatting,
                    styleFormatting);
            }
            formatting = MergeCharacterDefaults(formatting, tableRunFormatting);
        }
        var start = target.Length;
        return DxpDisposable.Create(() =>
        {
            if (target.Length <= start) return;
            if (!_runs.TryGetValue(target, out var runs))
                _runs[target] = runs = new List<DocPlainTextFormatRun>();
            var end = target.Length;
            var cursor = start;
            var symbols = _symbols.TryGetValue(target, out var found) ? found : null;
            foreach (var symbol in (symbols?.Where(x => x.Key >= start && x.Key < end)
                .OrderBy(x => x.Key) as IEnumerable<KeyValuePair<int, (string Font, ushort Character)>>)
                ?? Array.Empty<KeyValuePair<int, (string Font, ushort Character)>>())
            {
                if (symbol.Key > cursor && !formatting.IsEmpty)
                    runs.Add(new DocPlainTextFormatRun(cursor, symbol.Key, formatting));
                runs.Add(new DocPlainTextFormatRun(symbol.Key, symbol.Key + 1,
                    formatting with { SymbolFontName = symbol.Value.Font,
                        SymbolCharacter = symbol.Value.Character }));
                cursor = symbol.Key + 1;
            }
            if (cursor < end && !formatting.IsEmpty)
                runs.Add(new DocPlainTextFormatRun(cursor, end, formatting));
        });
    }

    private DocCharacterFormatting ReadCharacterFormatting(
        DocumentFormat.OpenXml.OpenXmlCompositeElement? properties, int? styleIndex = null)
    {
        ushort? ReadCharacterScale()
        {
            var value = properties?.GetFirstChild<CharacterScale>()?.Val?.Value;
            if (value == null) return null;
            if (value is < 1 or > 600)
                throw new NotSupportedException("The DOC character scale must be 1â€“600 percent.");
            return checked((ushort)value.Value);
        }
        short? ReadBaselineOffset()
        {
            var raw = properties?.GetFirstChild<Position>()?.Val?.Value;
            if (raw == null) return null;
            if (!short.TryParse(raw, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var value) ||
                value is < -3168 or > 3168)
                throw new NotSupportedException("The DOC baseline offset must be within Â±3168 half-points.");
            return value;
        }
        bool? ReadOnOff<T>() where T : DocumentFormat.OpenXml.Wordprocessing.OnOffType
        {
            var element = properties?.GetFirstChild<T>();
            return element == null ? null : element.Val?.Value ?? true;
        }
        bool? ReadRightToLeftText()
        {
            var typed = ReadOnOff<RightToLeftText>();
            if (typed != null) return typed;
            // The SDK loads w:rtl in a character style's rPr as an unknown
            // element, although Word accepts and applies it there.
            var raw = properties?.ChildElements.FirstOrDefault(x =>
                x.LocalName == "rtl" && x.NamespaceUri ==
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            if (raw == null) return null;
            var value = raw.GetAttributes().FirstOrDefault(x => x.LocalName == "val" &&
                x.NamespaceUri ==
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
            return value is not ("0" or "false" or "off");
        }
        ushort? size = ushort.TryParse(properties?.GetFirstChild<FontSize>()?.Val?.Value,
            out var parsedSize) ? parsedSize : null;
        var underline = properties?.GetFirstChild<Underline>()?.Val?.Value;
        byte? underlineCode = underline == UnderlineValues.None ? (byte)0 :
            underline == UnderlineValues.Single ? (byte)1 :
            underline == UnderlineValues.Words ? (byte)2 :
            underline == UnderlineValues.Double ? (byte)3 :
            underline == UnderlineValues.Dotted ? (byte)4 :
            underline == UnderlineValues.Thick ? (byte)6 :
            underline == UnderlineValues.Dash ? (byte)7 :
            underline == UnderlineValues.DotDash ? (byte)9 :
            underline == UnderlineValues.DotDotDash ? (byte)10 :
            underline == UnderlineValues.Wave ? (byte)11 :
            underline == UnderlineValues.DottedHeavy ? (byte)20 :
            underline == UnderlineValues.DashedHeavy ? (byte)23 :
            underline == UnderlineValues.DashDotHeavy ? (byte)25 :
            underline == UnderlineValues.DashDotDotHeavy ? (byte)26 :
            underline == UnderlineValues.WavyHeavy ? (byte)27 :
            underline == UnderlineValues.DashLong ? (byte)39 :
            underline == UnderlineValues.WavyDouble ? (byte)43 :
            underline == UnderlineValues.DashLongHeavy ? (byte)55 : null;
        var textColor = properties?.GetFirstChild<Color>();
        var colorRef = ResolveColor(textColor?.Val?.Value,
            textColor?.ThemeColor?.Value, textColor?.ThemeTint?.Value,
            textColor?.ThemeShade?.Value);
        var underlineColor = properties?.GetFirstChild<Underline>();
        var underlineColorRef = ResolveColor(underlineColor?.Color?.Value,
            underlineColor?.ThemeColor?.Value, underlineColor?.ThemeTint?.Value,
            underlineColor?.ThemeShade?.Value);
        var highlight = properties?.GetFirstChild<Highlight>()?.Val?.Value;
        byte? highlightCode = highlight == HighlightColorValues.None ? (byte)0 :
            highlight == HighlightColorValues.Black ? (byte)1 :
            highlight == HighlightColorValues.Blue ? (byte)2 :
            highlight == HighlightColorValues.Cyan ? (byte)3 :
            highlight == HighlightColorValues.Green ? (byte)4 :
            highlight == HighlightColorValues.Magenta ? (byte)5 :
            highlight == HighlightColorValues.Red ? (byte)6 :
            highlight == HighlightColorValues.Yellow ? (byte)7 :
            highlight == HighlightColorValues.White ? (byte)8 :
            highlight == HighlightColorValues.DarkBlue ? (byte)9 :
            highlight == HighlightColorValues.DarkCyan ? (byte)10 :
            highlight == HighlightColorValues.DarkGreen ? (byte)11 :
            highlight == HighlightColorValues.DarkMagenta ? (byte)12 :
            highlight == HighlightColorValues.DarkRed ? (byte)13 :
            highlight == HighlightColorValues.DarkYellow ? (byte)14 :
            highlight == HighlightColorValues.DarkGray ? (byte)15 :
            highlight == HighlightColorValues.LightGray ? (byte)16 : null;
        var runShading = properties?.GetFirstChild<Shading>();
        DocCellShading? shading = null;
        var runPattern = runShading?.Val?.Value == ShadingPatternValues.Nil
            ? ushort.MaxValue : runShading != null
                ? DocShadingPatterns.ToDoc(runShading.Val?.Value ??
                    ShadingPatternValues.Clear) : null;
        if (runShading != null && runPattern is ushort pattern)
            shading = new DocCellShading(ReadShadingFill(runShading),
                ReadShadingForeground(runShading), pattern);
        var runBorder = DocParagraphBorder.FromOpenXml(
            properties?.GetFirstChild<Border>());
        var script = properties?.GetFirstChild<VerticalTextAlignment>()?.Val?.Value;
        byte? scriptCode = script == VerticalPositionValues.Baseline ? (byte)0 :
            script == VerticalPositionValues.Superscript ? (byte)1 :
            script == VerticalPositionValues.Subscript ? (byte)2 : null;
        var fonts = properties?.GetFirstChild<RunFonts>();
        var spacing = properties?.GetFirstChild<Spacing>();
        short? characterSpacing = short.TryParse(spacing?.Val?.Value.ToString(),
            out var spacingValue) ? spacingValue : null;
        var language = properties?.GetFirstChild<Languages>();
        static ushort? LanguageId(string? tag)
        {
            if (string.IsNullOrWhiteSpace(tag)) return null;
            try
            {
                var id = System.Globalization.CultureInfo.GetCultureInfo(tag).LCID;
                if (id is > 0 and <= ushort.MaxValue && id != 0x1000)
                    return (ushort)id;
            }
            catch (System.Globalization.CultureNotFoundException) { }
            throw new NotSupportedException($"The DOC language '{tag}' has no standard LCID.");
        }
        var fitTextElement = properties?.GetFirstChild<FitText>();
        DocFitText? fitText = null;
        if (fitTextElement?.Val?.Value is uint width)
        {
            if (width > int.MaxValue)
                throw new NotSupportedException("DOC fit-text widths must fit a signed 32-bit twip value.");
            fitText = new DocFitText((int)width, fitTextElement.Id?.Value ?? 0);
        }
        byte? emphasisMarkCode = properties?.GetFirstChild<Emphasis>()?.Val?.Value switch
        {
            var value when value == EmphasisMarkValues.None => 0,
            var value when value == EmphasisMarkValues.Dot => 1,
            var value when value == EmphasisMarkValues.Comma => 2,
            var value when value == EmphasisMarkValues.Circle => 3,
            var value when value == EmphasisMarkValues.UnderDot => 4,
            null => null,
            _ => throw new NotSupportedException("Unsupported emphasis mark.")
        };
        return new DocCharacterFormatting(ReadOnOff<Bold>(), ReadOnOff<Italic>(), size,
            styleIndex, ReadOnOff<Strike>(), ReadOnOff<Caps>(), ReadOnOff<SmallCaps>(),
            ReadOnOff<Vanish>(), underlineCode, colorRef, highlightCode, scriptCode,
            AsciiFontName: ResolveThemeLatinFont(fonts?.AsciiTheme?.Value) ??
                fonts?.Ascii?.Value,
            EastAsiaFontName: ResolveThemeEastAsiaFont(fonts?.EastAsiaTheme?.Value) ??
                fonts?.EastAsia?.Value,
            HighAnsiFontName: ResolveThemeLatinFont(fonts?.HighAnsiTheme?.Value) ??
                fonts?.HighAnsi?.Value,
            ComplexScriptFontName: ResolveThemeComplexFont(fonts?.ComplexScriptTheme?.Value)
                ?? fonts?.ComplexScript?.Value,
            ComplexScriptBold: ReadOnOff<BoldComplexScript>(),
            ComplexScriptItalic: ReadOnOff<ItalicComplexScript>(),
            RightToLeftText: ReadRightToLeftText(),
            ForceComplexScript: ReadOnOff<ComplexScript>(),
            ComplexScriptSizeHalfPoints: ushort.TryParse(
                properties?.GetFirstChild<FontSizeComplexScript>()?.Val?.Value,
                out var complexSize) ? complexSize : null,
            CharacterSpacingTwips: characterSpacing,
            KerningThresholdHalfPoints: ushort.TryParse(
                properties?.GetFirstChild<Kern>()?.Val?.Value.ToString(),
                out var kerningThreshold) ? kerningThreshold : null,
            LanguageId: LanguageId(language?.Val?.Value),
            EastAsiaLanguageId: LanguageId(language?.EastAsia?.Value),
            ComplexScriptLanguageId: LanguageId(language?.Bidi?.Value),
            UnderlineColorRef: underlineColorRef, Shading: shading,
            Border: runBorder,
            DoubleStrike: ReadOnOff<DoubleStrike>(),
            CharacterScalePercent: ReadCharacterScale(),
            BaselineOffsetHalfPoints: ReadBaselineOffset(),
            Shadow: ReadOnOff<Shadow>(), Outline: ReadOnOff<Outline>(),
            Emboss: ReadOnOff<Emboss>(), Imprint: ReadOnOff<Imprint>(),
            SnapToGrid: ReadOnOff<SnapToGrid>(),
            EmphasisMarkCode: emphasisMarkCode, FitText: fitText);
    }

    private string? ResolveThemeLatinFont(ThemeFontValues? theme)
    {
        if (theme == ThemeFontValues.MajorAscii ||
            theme == ThemeFontValues.MajorHighAnsi)
            return ResolveScriptFont(_themeMajorScriptFonts, _themeLatinLanguage)
                ?? _themeMajorLatinFont;
        if (theme == ThemeFontValues.MinorAscii ||
            theme == ThemeFontValues.MinorHighAnsi)
            return ResolveScriptFont(_themeMinorScriptFonts, _themeLatinLanguage)
                ?? _themeMinorLatinFont;
        return null;
    }

    private string? ResolveThemeEastAsiaFont(ThemeFontValues? theme)
    {
        if (theme == ThemeFontValues.MajorEastAsia)
            return ResolveScriptFont(_themeMajorScriptFonts, _themeEastAsiaLanguage)
                ?? _themeMajorEastAsiaFont;
        if (theme == ThemeFontValues.MinorEastAsia)
            return ResolveScriptFont(_themeMinorScriptFonts, _themeEastAsiaLanguage)
                ?? _themeMinorEastAsiaFont;
        return null;
    }

    private string? ResolveThemeComplexFont(ThemeFontValues? theme)
    {
        if (theme == ThemeFontValues.MajorBidi)
            return ResolveScriptFont(_themeMajorScriptFonts, _themeBidiLanguage)
                ?? _themeMajorComplexFont;
        if (theme == ThemeFontValues.MinorBidi)
            return ResolveScriptFont(_themeMinorScriptFonts, _themeBidiLanguage)
                ?? _themeMinorComplexFont;
        return null;
    }

    private static string? ResolveScriptFont(
        IReadOnlyDictionary<string, string> fonts, string? language)
    {
        if (string.IsNullOrWhiteSpace(language)) return null;
        var tag = language.ToLowerInvariant();
        var script = tag.StartsWith("ja", StringComparison.Ordinal) ? "Jpan" :
            tag.StartsWith("ko", StringComparison.Ordinal) ? "Hang" :
            tag.StartsWith("zh-hant", StringComparison.Ordinal) ||
                tag.StartsWith("zh-tw", StringComparison.Ordinal) ||
                tag.StartsWith("zh-hk", StringComparison.Ordinal) ||
                tag.StartsWith("zh-mo", StringComparison.Ordinal) ? "Hant" :
            tag.StartsWith("zh", StringComparison.Ordinal) ? "Hans" : null;
        if (script == null)
            script = tag.StartsWith("ar", StringComparison.Ordinal) ? "Arab" :
                tag.StartsWith("he", StringComparison.Ordinal) ? "Hebr" :
                tag.StartsWith("th", StringComparison.Ordinal) ? "Thai" :
                tag.StartsWith("hi", StringComparison.Ordinal) ? "Deva" : null;
        return script != null && fonts.TryGetValue(script, out var name) ? name : null;
    }

    private static (string? MajorLatin, string? MinorLatin,
        string? MajorEastAsia, string? MinorEastAsia,
        string? MajorComplex, string? MinorComplex,
        IReadOnlyDictionary<string, string> MajorScripts,
        IReadOnlyDictionary<string, string> MinorScripts) ReadThemeFonts(ThemePart? themePart)
    {
        if (themePart == null) return (null, null, null, null, null, null,
            new Dictionary<string, string>(), new Dictionary<string, string>());
        using var stream = themePart.GetStream(FileMode.Open, FileAccess.Read);
        var document = XDocument.Load(stream);
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var scheme = document.Descendants(drawing + "fontScheme").FirstOrDefault();
        static string? Typeface(XElement? font, XNamespace drawing, string element)
        {
            var name = font?.Element(drawing + element)?.Attribute("typeface")?.Value;
            return string.IsNullOrWhiteSpace(name) ? null : name;
        }
        var major = scheme?.Element(drawing + "majorFont");
        var minor = scheme?.Element(drawing + "minorFont");
        static IReadOnlyDictionary<string, string> Scripts(XElement? font, XNamespace drawing) =>
            font?.Elements(drawing + "font")
                .Where(x => !string.IsNullOrWhiteSpace((string?)x.Attribute("script")) &&
                    !string.IsNullOrWhiteSpace((string?)x.Attribute("typeface")))
                .GroupBy(x => (string)x.Attribute("script")!, StringComparer.Ordinal)
                .ToDictionary(x => x.Key,
                    x => (string)x.First().Attribute("typeface")!, StringComparer.Ordinal)
            ?? new Dictionary<string, string>();
        return (Typeface(major, drawing, "latin"), Typeface(minor, drawing, "latin"),
            Typeface(major, drawing, "ea"), Typeface(minor, drawing, "ea"),
            Typeface(major, drawing, "cs"), Typeface(minor, drawing, "cs"),
            Scripts(major, drawing), Scripts(minor, drawing));
    }

    private static IReadOnlyDictionary<string, string> ReadThemeColors(ThemePart? themePart)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        if (themePart == null) return result;
        using var stream = themePart.GetStream(FileMode.Open, FileAccess.Read);
        var document = XDocument.Load(stream);
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var scheme = document.Descendants(drawing + "clrScheme").FirstOrDefault();
        if (scheme == null) return result;
        foreach (var entry in scheme.Elements())
        {
            var color = entry.Elements().FirstOrDefault();
            var value = color?.Attribute("val")?.Value ??
                color?.Attribute("lastClr")?.Value;
            if (value?.Length == 6 && uint.TryParse(value,
                System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out _))
                result[entry.Name.LocalName] = value;
        }
        return result;
    }

    private uint? ResolveColor(string? fallback, ThemeColorValues? theme,
        string? tint, string? shade)
    {
        var mappedKey = theme switch
        {
            var value when value == ThemeColorValues.Dark1 => "dark1",
            var value when value == ThemeColorValues.Light1 => "light1",
            var value when value == ThemeColorValues.Dark2 => "dark2",
            var value when value == ThemeColorValues.Light2 => "light2",
            var value when value == ThemeColorValues.Accent1 => "accent1",
            var value when value == ThemeColorValues.Accent2 => "accent2",
            var value when value == ThemeColorValues.Accent3 => "accent3",
            var value when value == ThemeColorValues.Accent4 => "accent4",
            var value when value == ThemeColorValues.Accent5 => "accent5",
            var value when value == ThemeColorValues.Accent6 => "accent6",
            var value when value == ThemeColorValues.Hyperlink => "hyperlink",
            var value when value == ThemeColorValues.FollowedHyperlink => "followedHyperlink",
            var value when value == ThemeColorValues.Background1 => "bg1",
            var value when value == ThemeColorValues.Text1 => "t1",
            var value when value == ThemeColorValues.Background2 => "bg2",
            var value when value == ThemeColorValues.Text2 => "t2",
            _ => null
        };
        if (mappedKey != null)
        {
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            while (true)
            {
                if (!visited.Add(mappedKey))
                    throw new InvalidDataException("The DOCX theme color mapping contains a cycle.");
                if (!_themeColorMapping.TryGetValue(mappedKey, out var next) ||
                    string.Equals(next, mappedKey, StringComparison.OrdinalIgnoreCase))
                    break;
                mappedKey = next;
            }
        }
        var schemeKey = mappedKey switch
        {
            "dark1" or "text1" or "t1" => "dk1",
            "light1" or "background1" or "bg1" => "lt1",
            "dark2" or "text2" or "t2" => "dk2",
            "light2" or "background2" or "bg2" => "lt2",
            "hyperlink" => "hlink",
            "followedHyperlink" => "folHlink",
            _ => mappedKey
        };
        if (schemeKey != null && _themeColors.TryGetValue(schemeKey, out var themed))
        {
            var modifier = tint ?? shade;
            fallback = modifier == null ? themed :
                ApplyThemeLuminance(themed, modifier, tint != null) ?? fallback;
        }
        if (string.Equals(fallback, "auto", StringComparison.OrdinalIgnoreCase))
            return 0xFF000000;
        if (fallback?.Length != 6 || !uint.TryParse(fallback,
            System.Globalization.NumberStyles.HexNumber,
            System.Globalization.CultureInfo.InvariantCulture, out var rgb))
            return null;
        return ((rgb >> 16) & 0xFF) | (rgb & 0xFF00) | ((rgb & 0xFF) << 16);
    }

    private static string? ApplyThemeLuminance(string color, string modifier,
        bool tint)
    {
        if (modifier.Length != 2 || !byte.TryParse(modifier,
            System.Globalization.NumberStyles.HexNumber,
            System.Globalization.CultureInfo.InvariantCulture, out var amount) ||
            !uint.TryParse(color,
                System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out var rgb))
            return null;
        var red = ((rgb >> 16) & 0xFF) / 255d;
        var green = ((rgb >> 8) & 0xFF) / 255d;
        var blue = (rgb & 0xFF) / 255d;
        var maximum = Math.Max(red, Math.Max(green, blue));
        var minimum = Math.Min(red, Math.Min(green, blue));
        var delta = maximum - minimum;
        var luminance = (maximum + minimum) / 2;
        var saturation = delta == 0 ? 0 :
            delta / (1 - Math.Abs(2 * luminance - 1));
        var hue = delta == 0 ? 0 : maximum == red
            ? ((green - blue) / delta + 6) % 6
            : maximum == green ? (blue - red) / delta + 2
            : (red - green) / delta + 4;
        var factor = amount / 255d;
        luminance = tint ? luminance * factor + 1 - factor : luminance * factor;
        var chroma = (1 - Math.Abs(2 * luminance - 1)) * saturation;
        var x = chroma * (1 - Math.Abs(hue % 2 - 1));
        var middle = luminance - chroma / 2;
        var components = hue switch
        {
            < 1 => (chroma, x, 0d),
            < 2 => (x, chroma, 0d),
            < 3 => (0d, chroma, x),
            < 4 => (0d, x, chroma),
            < 5 => (x, 0d, chroma),
            _ => (chroma, 0d, x)
        };
        static int Channel(double value) => Math.Max(0, Math.Min(255,
            (int)Math.Round(value * 255, MidpointRounding.AwayFromZero)));
        return $"{Channel(components.Item1 + middle):X2}" +
            $"{Channel(components.Item2 + middle):X2}" +
            $"{Channel(components.Item3 + middle):X2}";
    }

    public override void VisitText(Text text, DxpIDocumentContext context)
    {
        _paragraphText?.Append(text.Text);
    }

    public override void VisitSymbol(SymbolChar sym, DxpIDocumentContext context)
    {
        var target = _paragraphText ?? throw new InvalidDataException(
            "A DOC symbol needs a paragraph.");
        var font = sym.Font?.Value;
        var code = sym.Char?.Value;
        if (string.IsNullOrWhiteSpace(font) || code == null ||
            !ushort.TryParse(code, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out var character))
            throw new NotSupportedException("A DOC symbol needs a font and character code.");
        if (!_symbols.TryGetValue(target, out var symbols))
            _symbols[target] = symbols = new Dictionary<int, (string, ushort)>();
        symbols.Add(target.Length, (font, character));
        target.Append('\u0028');
    }

    public override IDisposable VisitDrawingBegin(Drawing drawing, DxpDrawingInfo? info,
        DxpIDocumentContext context)
    {
        var target = _paragraphText;
        if (target == null || !drawing.Descendants<
            DocumentFormat.OpenXml.Drawing.Blip>().Any())
            return DxpDisposable.Empty;
        var anchor = drawing.Descendants<
            DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().FirstOrDefault();
        if (anchor == null && !drawing.Descendants<
            DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>().Any())
            throw new NotSupportedException("The DOCX drawing has no inline or anchored picture.");
        if (info?.EmbedRelId == null || context.CurrentPart == null ||
            context.CurrentPart.GetPartById(info.EmbedRelId) is not ImagePart image)
            throw new NotSupportedException("A DOC inline image needs an embedded image part.");
        var presentation = info.Presentation;
        if (presentation?.FrameWidthPoints is not double width ||
            presentation.FrameHeightPoints is not double height ||
            !DocBinaryCompat.IsFinite(width) || !DocBinaryCompat.IsFinite(height))
            throw new NotSupportedException("A DOC inline image needs a finite display size.");
        using var imageStream = image.GetStream(FileMode.Open, FileAccess.Read);
        using var data = new MemoryStream();
        imageStream.CopyTo(data);
        var picture = new DocInlinePicture(data.ToArray(), image.ContentType,
            checked((long)Math.Round(width * 12700)),
            checked((long)Math.Round(height * 12700)),
            presentation.Crop is { IsEmpty: false } crop
                ? new DocInlinePictureCrop(crop.Left, crop.Top, crop.Right, crop.Bottom)
                : null, presentation.FlipHorizontal, presentation.FlipVertical,
            presentation.RotationDegrees);
        if (anchor != null)
        {
            static byte Alignment(string? value, bool vertical) => (value, vertical) switch
            {
                (null, _) => 0,
                ("left", false) or ("top", true) => 1,
                ("center", _) => 2,
                ("right", false) or ("bottom", true) => 3,
                ("inside", _) => 4,
                ("outside", _) => 5,
                _ => throw new NotSupportedException(vertical
                    ? "DOC floating image vertical alignment is unsupported."
                    : "DOC floating image horizontal alignment is unsupported.")
            };
            var horizontalAlignment = Alignment(anchor.HorizontalPosition?
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment>()?
                .Text, false);
            var verticalAlignment = Alignment(anchor.VerticalPosition?
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalAlignment>()?
                .Text, true);
            var horizontalOffset = anchor.HorizontalPosition?.PositionOffset?.Text;
            var verticalOffset = anchor.VerticalPosition?.PositionOffset?.Text;
            if ((horizontalAlignment == 0 && !long.TryParse(horizontalOffset, out _)) ||
                (verticalAlignment == 0 && !long.TryParse(verticalOffset, out _)))
                throw new NotSupportedException(
                    "DOC floating image writing requires an alignment or numeric position offset.");
            var x = horizontalAlignment == 0 ? long.Parse(horizontalOffset!) : 0L;
            var y = verticalAlignment == 0 ? long.Parse(verticalOffset!) : 0L;
            var horizontalRelative = anchor.HorizontalPosition.RelativeFrom?.Value;
            byte horizontalOrigin = horizontalRelative ==
                DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Margin
                ? (byte)0 : horizontalRelative ==
                DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Page
                ? (byte)1 : horizontalRelative ==
                DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Column
                ? (byte)2 : throw new NotSupportedException(
                    "DOC floating image horizontal origin is unsupported.");
            var verticalRelative = anchor.VerticalPosition.RelativeFrom?.Value;
            byte verticalOrigin = verticalRelative ==
                DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Margin
                ? (byte)0 : verticalRelative ==
                DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Page
                ? (byte)1 : verticalRelative ==
                DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Paragraph
                ? (byte)2 : throw new NotSupportedException(
                    "DOC floating image vertical origin is unsupported.");
            byte wrapCode;
            if (anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapSquare>() != null)
                wrapCode = 2;
            else if (anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapTopBottom>() != null)
                wrapCode = 1;
            else if (anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapNone>() != null)
                wrapCode = 3;
            else if (anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapTight>() != null)
                wrapCode = 4;
            else if (anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapThrough>() != null)
                wrapCode = 5;
            else
                throw new NotSupportedException("DOC floating image wrap mode is unsupported.");
            if (wrapCode is 4 or 5)
            {
                var wrap = wrapCode == 4
                    ? (DocumentFormat.OpenXml.OpenXmlElement?)anchor.GetFirstChild<
                        DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapTight>()
                    : anchor.GetFirstChild<
                        DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapThrough>();
                var polygon = wrap?.GetFirstChild<
                    DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapPolygon>();
                var start = polygon?.GetFirstChild<
                    DocumentFormat.OpenXml.Drawing.Wordprocessing.StartPoint>();
                var lines = polygon?.Elements<
                    DocumentFormat.OpenXml.Drawing.Wordprocessing.LineTo>().ToArray() ?? [];
                if (polygon?.Edited?.Value == true || start == null || lines.Length != 4 ||
                    start.X?.Value != lines[3].X?.Value ||
                    start.Y?.Value != lines[3].Y?.Value ||
                    start.X?.Value != lines[0].X?.Value ||
                    lines[0].Y?.Value != lines[1].Y?.Value ||
                    lines[1].X?.Value != lines[2].X?.Value ||
                    lines[2].Y?.Value != start.Y?.Value)
                    throw new NotSupportedException(
                        "DOC floating picture writing supports rectangular unedited wrap polygons.");
            }
            var side = wrapCode switch
            {
                2 => anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapSquare>()?
                    .WrapText?.Value,
                4 => anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapTight>()?
                    .WrapText?.Value,
                5 => anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapThrough>()?
                    .WrapText?.Value,
                _ => null
            };
            byte wrapSide = side == DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapTextValues.Left
                ? (byte)1 : side == DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapTextValues.Right
                ? (byte)2 : side == DocumentFormat.OpenXml.Drawing.Wordprocessing.WrapTextValues.Largest
                ? (byte)3 : (byte)0;
            _floatingPictures.Add(new FloatingPictureCapture(target, target.Length, picture,
                checked((int)Math.Round(x / 635m)), checked((int)Math.Round(y / 635m)),
                wrapCode, anchor.BehindDoc?.Value == true, wrapSide,
                horizontalOrigin, verticalOrigin,
                horizontalAlignment, verticalAlignment,
                checked((int)(anchor.DistanceFromTop?.Value ?? 0U)),
                checked((int)(anchor.DistanceFromBottom?.Value ?? 0U)),
                checked((int)(anchor.DistanceFromLeft?.Value ?? 114300U)),
                checked((int)(anchor.DistanceFromRight?.Value ?? 114300U))));
            target.Append('\u0008');
        }
        else
        {
            _pictures.Add(new PictureCapture(target, target.Length, picture));
            target.Append('\u0001');
        }
        return DxpDisposable.Empty;
    }

    public override void VisitPositionalTab(PositionalTab tab, DxpIDocumentContext context)
    {
        if (_paragraphText == null || _positionalTabs == null)
            throw new InvalidDataException("A positional tab needs a paragraph.");
        var section = (_section?.Formatting ?? new DocSectionFormatting())
            .WithWriterDefaults();
        if (tab.RelativeTo?.Value != AbsolutePositionTabPositioningBaseValues.Margin &&
            tab.RelativeTo?.Value != AbsolutePositionTabPositioningBaseValues.Indent)
            throw new NotSupportedException("The positional tab positioning base is unsupported in DOC.");
        if (tab.Alignment?.Value == AbsolutePositionTabAlignmentValues.Left)
        {
            if (_paragraphText.Length > _positionalLineStart)
            {
                _paragraphText.Append('\v');
                _positionalLineStart = _paragraphText.Length;
            }
            return;
        }
        if (tab.Alignment?.Value != AbsolutePositionTabAlignmentValues.Center &&
            tab.Alignment?.Value != AbsolutePositionTabAlignmentValues.Right)
            throw new NotSupportedException(
                "Only center/right positional tabs relative to margins or indents are supported in DOC.");
        var width = _positionalCellContentWidth ??
            section.Width!.Value - section.Left!.Value - section.Right!.Value;
        var indent = tab.RelativeTo?.Value == AbsolutePositionTabPositioningBaseValues.Indent;
        var effective = _positionalTabParagraphFormatting ?? DocParagraphFormatting.Empty;
        var position = tab.Alignment?.Value == AbsolutePositionTabAlignmentValues.Center
            ? width / 2 + (indent ? ((effective.LeftTwips ?? 0) +
                (effective.FirstLineTwips ?? 0)) / 2 : 0)
            : width - (indent ? effective.RightTwips ?? 0 : 0);
        if (position <= 0 || position > short.MaxValue)
            throw new NotSupportedException("The positional tab is outside the DOC page.");
        var leader = tab.Leader?.Value;
        byte leaderCode = leader == AbsolutePositionTabLeaderCharValues.None ? (byte)0 :
            leader == AbsolutePositionTabLeaderCharValues.Dot ? (byte)1 :
            leader == AbsolutePositionTabLeaderCharValues.Hyphen ? (byte)2 :
            leader == AbsolutePositionTabLeaderCharValues.Underscore ? (byte)3 :
            leader == AbsolutePositionTabLeaderCharValues.MiddleDot ? (byte)5 :
            throw new NotSupportedException("The positional tab leader is unsupported in DOC.");
        _positionalTabs.Add(new DocTabStop((short)position,
            tab.Alignment?.Value == AbsolutePositionTabAlignmentValues.Center ? (byte)1 : (byte)2,
            leaderCode));
        _paragraphText.Append('\t');
    }

    private void AppendLegacyDateBlock(string picture)
    {
        if (_paragraphText == null) return;
        var result = DateTime.Today.ToString(picture,
            System.Globalization.CultureInfo.CurrentCulture);
        _paragraphText.Append('\u0013').Append(" DATE \\@ \"")
            .Append(picture).Append("\" ").Append('\u0014')
            .Append(result).Append('\u0015');
    }

    public override void VisitDayShort(DayShort day, DxpIDocumentContext context) =>
        AppendLegacyDateBlock("dd");

    public override void VisitDayLong(DayLong day, DxpIDocumentContext context) =>
        AppendLegacyDateBlock("dddd");

    public override void VisitMonthShort(MonthShort month, DxpIDocumentContext context) =>
        AppendLegacyDateBlock("MM");

    public override void VisitMonthLong(MonthLong month, DxpIDocumentContext context) =>
        AppendLegacyDateBlock("MMMM");

    public override void VisitYearShort(YearShort year, DxpIDocumentContext context) =>
        AppendLegacyDateBlock("yy");

    public override void VisitYearLong(YearLong year, DxpIDocumentContext context) =>
        AppendLegacyDateBlock("yyyy");

    public override void VisitPageNumber(PageNumber pageNumber, DxpIDocumentContext context)
    {
        // The legacy page-number block is always decimal, even when the
        // section uses another page-number format.
        _paragraphText?.Append("\u0013 PAGE \\* Arabic \u00141\u0015");
    }

    public override void VisitTab(TabChar tab, DxpIDocumentContext context)
    {
        _paragraphText?.Append('\t');
    }

    public override void VisitBreak(Break lineBreak, DxpIDocumentContext context)
    {
        _paragraphText?.Append(lineBreak.Type?.Value == BreakValues.Page ? '\f' :
            lineBreak.Type?.Value == BreakValues.Column ? '\u000E' : '\v');
        if (_paragraphText != null) _positionalLineStart = _paragraphText.Length;
    }

    public override void VisitCarriageReturn(CarriageReturn carriageReturn, DxpIDocumentContext context)
    {
        _paragraphText?.Append('\v');
        if (_paragraphText != null) _positionalLineStart = _paragraphText.Length;
    }

    public override void VisitNoBreakHyphen(NoBreakHyphen hyphen, DxpIDocumentContext context)
    {
        _paragraphText?.Append('\u001E');
    }

    public override void VisitSoftHyphen(SoftHyphen hyphen, DxpIDocumentContext context)
    {
        _paragraphText?.Append('\u001F');
    }

    public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target,
        DxpIDocumentContext context)
    {
        var text = _paragraphText;
        if (text == null) return DxpDisposable.Empty;
        if (target == null)
            throw new InvalidDataException("A DOCX hyperlink has no resolvable target.");
        var destination = target.internalRef ?? target.uri;
        if (string.IsNullOrWhiteSpace(destination) || destination.Any(x =>
            x == '"' || char.IsControl(x)))
            throw new NotSupportedException("The DOCX hyperlink target cannot be encoded as a DOC field.");
        var instruction = target.internalRef != null
            ? $" HYPERLINK \\l \"{destination}\" "
            : $" HYPERLINK \"{destination}\" ";
        text.Append('\u0013').Append(instruction).Append('\u0014');
        return DxpDisposable.Create(() => text.Append('\u0015'));
    }

    public override void VisitBookmarkStart(BookmarkStart bookmark, DxpIDocumentContext context)
    {
        if (_suppressBookmarkCapture) return;
        var story = _paragraphText ?? _storyText ??
            (context.CurrentPart == _mainPart ? _text : null);
        if (story == null) return;
        var id = bookmark.Id?.Value ?? throw new InvalidDataException("A DOCX bookmark has no ID.");
        var name = bookmark.Name?.Value ?? throw new InvalidDataException("A DOCX bookmark has no name.");
        if (_bookmarks.ContainsKey(id))
            throw new InvalidDataException($"DOCX bookmark ID '{id}' is duplicated.");
        _bookmarks.Add(id, new BookmarkCapture(name, story, story.Length));
    }

    public override void VisitBookmarkEnd(BookmarkEnd bookmark, DxpIDocumentContext context)
    {
        if (_suppressBookmarkCapture) return;
        var story = _paragraphText ?? _storyText ??
            (context.CurrentPart == _mainPart ? _text : null);
        if (story == null) return;
        var id = bookmark.Id?.Value;
        if (id == null || !_bookmarks.TryGetValue(id, out var start) ||
            !ReferenceEquals(start.Story, story) || start.End != null)
            throw new InvalidDataException("A DOCX bookmark end is unmatched.");
        start.End = story.Length;
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField field, DxpIDocumentContext context)
    {
        var target = _paragraphText;
        if (target == null) return DxpDisposable.Empty;
        target.Append('\u0013');
        target.Append(field.Instruction?.Value ?? string.Empty);
        target.Append('\u0014');
        return DxpDisposable.Create(() => target.Append('\u0015'));
    }

    public override void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext context)
        => _paragraphText?.Append('\u0013');

    public override void VisitComplexFieldInstruction(FieldCode instruction, string text,
        DxpIDocumentContext context) => _paragraphText?.Append(text);

    public override void VisitComplexFieldCachedResultText(string text,
        DxpIDocumentContext context) => _paragraphText?.Append(text);

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext context)
        => _paragraphText?.Append('\u0014');

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext context)
        => _paragraphText?.Append('\u0015');

    private static int? ReadStatistic(string? text) =>
        int.TryParse(text, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var value) && value >= 0
            ? value : null;
}
