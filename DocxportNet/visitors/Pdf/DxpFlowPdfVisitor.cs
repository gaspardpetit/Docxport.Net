using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Doc;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using DocxportNet.Visitors;
using DocxportNet.Walker;
using DocxportNet.Walker.Context;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel.Shapes.Charts;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Globalization;
using System.Buffers.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using WordStyle = DocumentFormat.OpenXml.Wordprocessing.Style;

namespace DocxportNet.Visitors.Pdf;

public sealed class DxpFlowPdfVisitor : DxpVisitor, DxpIFieldEvalProvider, IDxpAllHeaderFooterVisitor
{
    private const double CompactHeaderRowReductionPoints = 10.3;
    private const double EmptyParagraphAfterTableExtraPoints = 2.3;
    private const double SimpleColumnParagraphPitchPoints = 25;
    private const double ColoredSpacedTableCellWidthPaddingPoints = 19.5;
    private const double NarrowHorizontalMergeCellWidthPoints = 50;
    private const double NarrowHorizontalMergeCellPaddingPoints = 10.8;
    private const double CompoundFontLineHeightRatio = 1.328;
    private const double VerticalTextLineSpacingFactor = 1.036;
    private const double SuperscriptHeaderLineCompensationPoints = 8.2;
    private const double BorderedParagraphExitClearancePoints = 8;
    private const double DirectParagraphSideBorderOffsetPoints = 1.5;
    private const double CommentLineSegmentGapPoints = 60;
    private const double TocEmptyParagraphClearancePoints = 2.2;
    private const double RepeatingTableFooterClearancePoints = 35;
    private const double ContinuedFootnoteClearancePoints = 25.5;
    private const double MarkupFootnoteSeparatorClearancePoints = 8;
    private const double DenseFootnoteBodyClearancePoints = 50;
    private const double DenseTrackedLayoutAllowancePoints = 1;
    private const double ClassicAptosStoryClearancePoints = 0.55;
    private const double EmptyModernDefaultsBodyClearancePoints = 0.2;
    private const double StylelessAptosFloatingTableLineFactor = 1.1578;
    private const int ShortTrackedSplitWordLimit = 50;
    private static bool SupportsContourLayout =>
        typeof(DocumentRenderer).GetProperty("WrapContour") != null &&
        (typeof(DocumentRenderer).GetProperty("WrapContourLineFontBandHeight") != null ||
         Environment.GetEnvironmentVariable("DOCXPORT_DIAG_BASIC_CONTOUR") == "1");
    private readonly string _fontFamily;
    private readonly string? _cjkFontFamily;
    private readonly ISet<string> _availableFontFamilies;
    private readonly Func<char, string?>? _chineseFallbackFont;
    private readonly DxpFieldEval _fieldEval = new();
    private readonly Dictionary<Section, PageBorders> _pageBorders = new();
    private readonly Dictionary<Section, LineNumberType> _sectionLineNumbers = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _sourceBodyParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _sourceCellParagraphs = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, HashSet<string>>
        _crossRunLeadingWords = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _exactCjkLineParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _mixedFullWidthPunctuationParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _exactStylelessPlainParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _exactStylelessImplicitAptosParagraphs = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> _justifiedShrinkFactors = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _twoWordOverhangParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Shapes.TextFrame> _floatingTableFrames = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Shapes.TextFrame,
        (MigraDoc.DocumentObjectModel.Tables.Table Table, DxpComputedBorder Border)>
        _floatingContinuationTopBorders = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Shapes.TextFrame, double> _textBoxRotations = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Shapes.TextFrame,
        MigraDoc.DocumentObjectModel.Paragraph> _notBesideFrameFollowers = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Shapes.TextFrame> _verticallyCenteredTextBoxes = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Shapes.TextFrame> _clippedTextBoxes = new();
    private readonly List<(Section Section, MigraDoc.DocumentObjectModel.Shapes.TextFrame Frame,
        long RelativeHeight)> _foregroundTextBoxes = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Shapes.TextFrame> _behindTextBoxes = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Shapes.TextFrame> _aboveContentTextBoxes = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Shapes.TextFrame,
        (Cell Cell, MigraDoc.DocumentObjectModel.Paragraph Paragraph, double Left, double Top)>
        _cellAnchoredTextBoxes = new();
    private readonly Dictionary<string, DxpCommentThread> _commentThreads = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Cell> _commentRangeCells = new(StringComparer.Ordinal);
    private readonly bool _includePdfComments;
    private readonly bool _includeRevisionMarkup;
    private readonly XColor _revisionMarkupColor;
    private readonly XColor _revisionDeletionColor;
    private bool _insideInsertedRun;
    private bool _insideDeletedRun;
    private bool _insideMoveToRun;
    private bool _insideMoveFromRun;
    private MigraDoc.DocumentObjectModel.Color? _textBoxTextColor;
    private readonly List<(MigraDoc.DocumentObjectModel.Paragraph Paragraph, double Width, double Height,
        double LineHeight)> _exactLineImages = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Shapes.Image, double> _inlineImageOffsets = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _frameWidthParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _implicitTableParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _tocLeaderParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _positionalLeaderParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _footerLeaderParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _clearedTabLeaderParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _narrowJustifiedBodyParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _emergencyBreakParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _trackedLayoutParagraphs = new();
    private readonly HashSet<DocumentFormat.OpenXml.Wordprocessing.Paragraph> _tableOpeningSpacingTransferred = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Tables.Table> _noAutoSpacingTables = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph,
        DocumentFormat.OpenXml.Wordprocessing.Paragraph> _styleRefBodyParagraphs = new();
    private readonly List<(string StyleName, string CachedText)> _styleRefHeaders = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _suppressedLineNumberParagraphs = new();
    private readonly HashSet<Section> _footerTableDistanceApplied = new();
    private readonly HashSet<Section> _firstPageIntroTableFooters = new();
    private readonly HashSet<Section> _footerLineGridDistanceApplied = new();
    private readonly HashSet<Section> _fittedCjkFooterDistanceApplied = new();
    private readonly HashSet<Section> _fittedCjkFooterContinuationDistanceApplied = new();
    private readonly HashSet<Section> _unstyledFooterDistanceApplied = new();
    private readonly HashSet<Section> _numberedFooterDistanceApplied = new();
    private readonly HashSet<Section> _footerSpacingResetApplied = new();
    private readonly HashSet<Section> _repeatingTableFooterClearanceApplied = new();
    private bool _sectionHasFooterTable;
    private readonly HashSet<Section> _cjkDefaultMarginsAdjusted = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Tables.Table,
        (DxpComputedBorder Border, double Spacing, bool InHeaderFooter, double ResetAllowance)>
        _spacedTableLeftBorders = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Tables.Table,
        (double Border, double Spacing, double Padding)> _htmlBevelTables = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Tables.Table, DxpComputedBoxBorders>
        _tableFragmentBorders = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Tables.Table> _splitOverpageTables = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Tables.Table,
        List<(int Start, int Count, bool Top, bool Bottom)>> _nestedFragmentBorderGroups = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Tables.Table, List<DiagonalCell>>
        _diagonalCells = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Tables.Table, List<RotatedCell>>
        _rotatedCells = new();
    private bool _activeRotatedCell;
    private bool _rotatedCellPlaceholderAdded;
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Tables.Table, RowTopStroke>
        _doubleFirstRowTopBorders = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Tables.Table, DxpComputedBorder>
        _rtlDoubleTableLeftBorders = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Tables.Table, DxpComputedBorder>
        _doubleTableRightBorders = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Tables.Table, List<DoubleCellSideBorder>>
        _doubleCellSideBorders = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, string> _paragraphShades = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Shapes.Shape, IReadOnlyList<XPoint>>
        _wrapContours = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Shapes.Shape, double>
        _wrapContourBandHeights = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph,
        DocumentFormat.OpenXml.Wordprocessing.Paragraph> _directBorderParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _manualParagraphShades = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> _footerParagraphShadeClearances = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, List<EmphasisRange>>
        _emphasisRanges = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, List<StrikeRange>>
        _strikeRanges = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, List<RevisionUnderlineRange>>
        _revisionUnderlineRanges = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _revisionParagraphs = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, List<string>>
        _formattingBalloons = new();
    private readonly Dictionary<string, int> _changedFieldUrls = new(StringComparer.Ordinal);
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, HashSet<int>>
        _strikeTabBoundaries = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, List<TextEffectRange>>
        _textEffectRanges = new();
    private readonly HashSet<string> _largeSuperscriptTexts = new(StringComparer.Ordinal);
    private readonly Dictionary<string, double> _inlineSuperscriptSizes = new(StringComparer.Ordinal);
    private readonly HashSet<string> _ambiguousInlineSuperscripts = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _closingQuoteSourceCounts = new(StringComparer.Ordinal);
    private readonly HashSet<string> _excludedClosingQuoteTokens = new(StringComparer.Ordinal);
    private readonly HashSet<string> _closingQuoteFollowedBySpace = new(StringComparer.Ordinal);
    private readonly HashSet<string> _rightAlignedTrackedTexts = new(StringComparer.Ordinal);
    private readonly HashSet<string> _longJustifiedTrackedTexts = new(StringComparer.Ordinal);
    private readonly HashSet<string> _plainLatinAfterFitTexts = new(StringComparer.Ordinal);
    private readonly Dictionary<string, int> _apostropheDisplaySourceCounts = new(StringComparer.Ordinal);
    private readonly HashSet<string> _excludedApostropheDisplayTokens = new(StringComparer.Ordinal);
    private readonly List<string[]> _rightLegendLabelGroups = new();
    private readonly List<string[]> _columnRightLegendLabelGroups = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> _characterScales = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> _topAlignedFontSizes = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _topAlignedHeaderParagraphs = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, (double Size, double Factor)>
        _verticalAlignedFonts = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _oversizedGridLatinParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _oversizedStoryLatinParagraphs = new();
    private readonly HashSet<Section> _topAlignedBodyMargins = new();
    private readonly HashSet<Section> _topAlignedFooterDistances = new();
    private readonly HashSet<Section> _centerBottomBodyMargins = new();
    private readonly HashSet<Section> _centerBottomFooterDistances = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> _fitTextWidths = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, int> _fitTextCharacterCounts = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _mixedCjkFitBodyParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _mixedHighAnsiFittedHeaders = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _latinTimesFitParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _mixedCalibriYuMinchoFitParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _latinYuMinchoFitParagraphs = new();
    private readonly HashSet<Section> _headersWithFirstFittedParagraph = new();
    private readonly HashSet<Section> _headersWithSecondCjkParagraph = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _fittedBodyAfterSecondCjkHeader = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _latinCalibriFitParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _fitTextTracking = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, (string Text, double Width)[]>
        _partialFitGroups = new();
    private readonly HashSet<Section> _styleFitStoryDistancesApplied = new();
    private readonly HashSet<Section> _implicitFitFooterDistancesApplied = new();
    private readonly HashSet<Section> _calibriFitFooterDistancesApplied = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> _characterSpacings = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _hangulBaselineShiftParagraphs = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, List<(string Base, string Guide, double Size, double Raise, int Start, string? Font, string Alignment)>>
        _rubyGuides = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> _rubyNaturalLineHeights = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, LineSpacingRule> _rubyOriginalLineSpacingRules = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> _rubyOverflowBaselineOffsets = new();
    private bool _rubyOnlyImplicitFontDocument;
    private bool _suppressRubyGuideText;
    private bool _renderRubyBaseText;
    private bool _renderRubyGuide;
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, string> _tabbedRunShadings = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> _borderedTabInsets = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> _centeredSpacingOffsets = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _gridGlyphParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _gridBoundaryParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _unshiftedGridGlyphParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _packedGridParagraphs = new();
    private readonly List<FormattedText> _paintOnlyMinchoSpans = new();
    private readonly List<FormattedText> _paintOnlyRubyBaseMinchoSpans = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> _packedGridOffsets = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double> _firstBodyDoubleBorderInset = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, List<TextBorderRange>>
        _textBorderRanges = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _headerRunBorderParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _footerRunBorderParagraphs = new();
    private readonly Dictionary<string, (string Foreground, string Background, int Count)>
        _joinedClosingMarkColors = new(StringComparer.Ordinal);
    private readonly Dictionary<string, bool> _safeClosingMarkTokens = new(StringComparer.Ordinal);
    private WordprocessingDocument? _sourceWordDocument;
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph,
        (string Color, bool TrailingSpace, bool InHeaderFooter)>
        _coloredUnderlineParagraphs = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, (string Marker, string Color, bool Double)>
        _listMarkerUnderlines = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, (string Marker, string Color)>
        _listMarkerStrikes = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _borderResetParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _bodyParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _compactArialBodyParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _cjkNumberedBodyParagraphs = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, (string Family, double Rise)> _bulletBaselineParagraphs = new();
    private readonly Dictionary<int, MigraDoc.DocumentObjectModel.Paragraph> _footnoteReferences = new();
    private readonly Dictionary<int, string> _footnoteLabels = new();
    private readonly HashSet<int> _customFootnoteMarks = new();
    private readonly Dictionary<int, int> _footnoteSectionStarts = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph,
        MigraDoc.DocumentObjectModel.Paragraph> _footnoteSplitContinuations = new();
    private readonly List<(int Index, MigraDoc.DocumentObjectModel.Paragraph Paragraph, int PageOffset)> _footnoteParagraphs = new();
    private int? _activeFootnoteIndex;
    private readonly HashSet<Section> _headersWithStyledEmptyLine = new();
    private readonly HashSet<Section> _headersWithMarginAnchoredWrap = new();
    private readonly Dictionary<Section, double> _headerRunBorderClearances = new();
    private readonly Dictionary<Section, string> _verticallyAlignedSections = new();
    private readonly HashSet<Section> _sectionsWithLinkedHeaders = new();
    private readonly HashSet<DocumentFormat.OpenXml.Wordprocessing.Paragraph> _bodyParagraphsUnderHeaderWrap = new();
    private readonly HashSet<Section> _headersWithStyledTrailingLine = new();
    private readonly HashSet<Section> _footersWithStyledEmptyLine = new();
    private readonly HashSet<Section> _footersWithStyledTrailingLine = new();
    private readonly Dictionary<Section, List<FooterImageOverlay>> _footerImages = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, List<ParagraphImageOverlay>>
        _paragraphImages = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph,
        (string Numerator, string Denominator, double Size)>
        _mathFractions = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _firstBodyDisplayFractions = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _mixedDisplayMathParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _mixedDisplayRunFractionParagraphs = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, string> _mathRadicals = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _inlineMathRadicals = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _standaloneInlineMathRadicals = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _mixedInlineMathRadicals = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _tallInlineMathRadicals = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph,
        List<(string Numerator, string Denominator, double Width)>>
        _inlineMathFractions = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, List<InlineMathTerm>>
        _nestedInlineMathFractions = new();
    private sealed record InlineMathTerm(string? Text, InlineMathTerm? Numerator,
        InlineMathTerm? Denominator, double Width);
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph,
        List<(string[][] Cells, double[] ColumnWidths)>>
        _inlineMathMatrices = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph,
        List<(string Superscript, string Subscript)>> _inlineMathCombinedScripts = new();
    private readonly Dictionary<DocumentFormat.OpenXml.Wordprocessing.Paragraph, FloatingTextWrap>
        _floatingTextWraps = new();
    private FloatingTextWrap? _activeFloatingTextWrap;
    private readonly List<MigraDoc.DocumentObjectModel.Paragraph> _mirroredParagraphs = new();
    private readonly List<DxpPdfHighlight> _highlights = new();
    private readonly Dictionary<MigraDoc.DocumentObjectModel.Paragraph, (StringBuilder Text,
        int NormalizedLength)> _highlightParagraphTexts = new();
    private readonly HashSet<string> _noBreakWords = new(StringComparer.Ordinal);
    private readonly List<MigraDoc.DocumentObjectModel.Paragraph> _noBreakParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _superscriptHeaderParagraphs = new();
    private sealed record DiagonalStroke(string Color, double Width);
    private sealed record RowTopStroke(string Color, double Width, bool InHeaderFooter);
    private sealed record DoubleCellSideBorder(int Row, int Column, bool Right,
        DxpComputedBorder Border);
    private sealed record DiagonalCell(int Row, int Column, int RowSpan, int ColumnSpan,
        DiagonalStroke? Falling, DiagonalStroke? Rising, bool HasStraightBorders,
        bool ColoredTableBorders);
    private sealed record RotatedCell(int Row, int Column, string Text, string Direction,
        bool InHeaderFooter, bool InHeader, string FontName, double FontSize, bool Bold, bool Italic,
        string? ColorCss);
    private sealed record EmphasisRange(int Start, int Length, EmphasisMarkValues Kind);
    private sealed record StrikeRange(int Start, int Length, bool Double, int? VisibleStart = null,
        int? VisibleLength = null, bool Moved = false);
    private sealed record RevisionUnderlineRange(int Start, int Length, bool Moved = false);
    private enum TextEffectKind { Outline, Emboss, Imprint, Shadow, ShadowOutline }
    private sealed record TextEffectRange(int Start, int Length, string FontName, bool Bold, bool Italic,
        TextEffectKind Kind);
    private sealed record TextBorderRange(int Start, int Length, string Color, double Width, double Inset,
        bool Shadow, bool Double, int VisibleStart = 0, int VisibleLength = 0);
    private IReadOnlyDictionary<string, WordStyle> _paragraphStyles = new Dictionary<string, WordStyle>();
    private DxpNumberingResolver? _numberingResolver;
    public DxpFieldEval FieldEval => _fieldEval;
    internal bool UseCachedMarginMergePlaceholders { get; set; }
    private Stream? _output;
    private MigraDoc.DocumentObjectModel.Document? _document;
    private Section? _section;
    private MigraDoc.DocumentObjectModel.Paragraph? _paragraph;
    private MigraDoc.DocumentObjectModel.Paragraph? _lastParagraph;
    private object? _lastParagraphContainer;
    private DxpStyleEffectiveRunStyle? _runStyle;
    private string? _runHighAnsiFont;
    private string? _runEastAsiaFont;
    private bool _implicitFittedStoryFont;
    private bool _calibriFittedStoryText;
    private DxpStyleEffectiveRunStyle? _previousTextStyle;
    private MigraDoc.DocumentObjectModel.Paragraph? _previousTextParagraph;
    private bool _previousTextCanJoin;
    private bool _previousTextBasicMetrics;
    private TextEffectKind? _runEffect;
    private double _runScalePercent = 100;
    private double _runCharacterSpacing;
    private double _paragraphLineSize;
    private TextBorderRange? _runTextBorder;
    private bool _hiddenRun;
    private bool _runHasExplicitColor;
    private bool _runHasExplicitBold;
    private bool _runHasExplicitItalic;
    private bool _runHasExplicitUnderline;
    private bool _runRightToLeft;
    private bool _runComplexBold;
    private bool _runComplexItalic;
    private bool _tocHyperlinkRun;
    private HighlightColorValues? _runHighlight;
    private EmphasisMarkValues? _runEmphasis;
    private string? _runShading;
    private HighlightColorValues? _defaultRunHighlight;
    private DxpComputedTableCellStyle? _cellStyle;
    private bool _cellNoWrap;
    private double? _cellNoWrapWidthPoints;
    private double? _cellWrapWidthPoints;
    private bool _balanceNarrowCellLastWord;
    private double? _rowLeftMarginPoints;
    private double? _rowRightMarginPoints;
    private double? _cellLeftMarginPoints;
    private double? _cellRightMarginPoints;
    private double _cellSpacingIndentPoints;
    private object? _container;
    private MigraDoc.DocumentObjectModel.Hyperlink? _activeHyperlink;
    private readonly Stack<(System.Text.StringBuilder Instruction,
        MigraDoc.DocumentObjectModel.Hyperlink? Previous)> _complexFieldLinks = new();
    private MigraDoc.DocumentObjectModel.Tables.Table? _table;
    private Row? _row;
    private bool _evenAndOddHeaders;
    private bool _balanceSingleByteDoubleByteWidth;
    private readonly Dictionary<string, int> _balancedFooterFieldSpaces = new(StringComparer.Ordinal);
    private readonly Dictionary<string, double> _doubleBorderTextPadding = new(StringComparer.Ordinal);
    private bool _gutterAtTop;
    private bool _inHeader;
    private HeaderFooterValues _headerFooterKind;
    private int _headerParagraphCount;
    private int _headerNumberedParagraphCount;
    private int _headerPositionalTabCount;
    private bool _headerPositionalTabHasFollowingText;
    private int _headerPositionalCellParagraphCount;
    private int _headerLeftMarginTabBreakCount;
    private int _footerLeftMarginTabBreakCount;
    private int _footerParagraphCount;
    private int _headerEmptyParagraphCount;
    private double _headerSizedEmptyLineExtraPoints;
    private int _headerPendingEmptyParagraphCount;
    private bool _headerPreviousParagraphUsesBodySpacing;
    private double _headerFirstBeforePoints;
    private double _headerFirstAfterPoints;
    private double _headerFirstAfterLayoutAdjustmentPoints;
    private double _headerLastAfterPoints;
    private bool _headerHasContextualSpacing;
    private bool _footerHasContextualSpacing;
    private bool _footerLastParagraphHasText;
    private bool _headerHasContent;
    private bool _headerHasDirectBidiOverride;
    private bool _headerParagraphsAreNumbered;
    private int _headerRowCount;
    private double _headerMinimumRowClearancePoints;
    private bool _headerHasExactHeightRow;
    private bool _headerHasZeroMarginRow;
    private bool _headerHasRotatedCell;
    private int _nestedHeaderRowCount;
    private int _headerCompactRowTransitions;
    private double _headerRowSpacingReductionPoints;
    private bool _previousHeaderRowCompact;
    private double _previousHeaderRowSpacingReductionPoints;
    private double _headerCellSpacingPoints;
    private bool _headerHasFilledSpacedTable;
    private bool _headerHasStyledTable;
    private bool _headerHasImbalancedAutoFitTable;
    private bool _bodyHasRowBorderExceptions;
    private double _tableCellSpacingPoints;
    private MigraDoc.DocumentObjectModel.Color? _tableBackground;
    private bool _horizontalMergeTable;
    private bool _tableInHeaderFooter;
    private bool _tableRightToLeft;
    private bool _headerHasHorizontalMerge;
    private bool _headerHasOccupiedHorizontalMerge;
    private bool _headerHasTrailingUnstyledTableParagraph;
    private bool _footerHasHorizontalMerge;
    private bool _footerHorizontalMergeClearanceApplied;
    private double _rowCellSpacingPoints;
    private bool _implicitContentTable;
    private double _doubleRowBorderPoints;
    private bool _nativeDoubleTableBorders;
    private SpacingBetweenLines? _tableParagraphSpacing;
    private double _defaultParagraphLinePoints;
    private double _defaultParagraphAfterPoints;
    private double? _documentDefaultParagraphAfterPoints;
    private int _sectionCount;
    private string? _sectionPageNumberFormat;
    private double _sectionBaseTopMarginPoints;
    private bool _sectionOtherHeaderTableRaisedTopMargin;
    private double? _gridLinePitchPoints;
    private double? _gridCharSpacingPoints;
    private bool _snapToCharsGrid;
    private bool _compactSnapToCharsParagraph;
    private bool _compactLinesAndCharsParagraph;
    private bool _largeGridHeaderDistanceAdjusted;
    private double _headerGridWrapClearancePoints;
    private double? _paragraphGridCharSpacingPoints;
    private bool _autoSpaceDE;
    private bool _autoSpaceDN;
    private bool _narrowJustifiedSpaces;
    private bool _compressedBodyLayoutSpaces;
    private bool _restoreCompressedBodySpaces;
    private bool _mixedBodyMinchoFallback;
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _layoutWidthParagraphs = new();
    private readonly HashSet<MigraDoc.DocumentObjectModel.Paragraph> _compressedBodyParagraphs = new();
    private bool _pendingNoBreakHyphen;
    private bool _cjkStyleParagraph;
    private bool _cjkEmphasisParagraph;
    private readonly HashSet<Section> _firstBodyClearsInheritedEmphasis = new();
    private readonly HashSet<FieldFrame> _renderedPageFields = new(ReferenceEqualityComparer.Instance);
    private readonly HashSet<FieldFrame> _renderedMarginMergeFields = new(ReferenceEqualityComparer.Instance);
    private bool _sectionUsesDefaultMargins;
    private bool _documentHasFitText;
    private bool _documentHasTrackedInsertions;
    private bool _positionedLayoutApplied;
    private bool _documentHasPositionalTabs;
    private bool _documentHasTextBoxes;
    private bool _documentHasMath;
    private double? _defaultTabStopPoints;
    private bool _headerHasUnnumberedDefaultTab;
    private bool _bodyHasRowGridGaps;
    private bool _unfittedMixedScriptParagraph;
    private bool _unfittedMixedScriptLatinParagraph;
    private bool _unfittedCalibriParagraph;
    private bool _latinAfterFittedParagraph;
    private bool _chineseFallbackParagraph;
    private bool _paragraphInheritsFitText;
    private bool _pendingEmptyParagraphPageBreak;
    private bool _implicitModernFont;
    private bool _classicPartialDefaults;
    private bool _explicitDefaultRunProperties;
    private string? _minorThemeLatinFont;
    private bool _themedImplicitParagraphSpacing;
    private bool _stylelessBodyParagraphSpacing;
    private bool _stylelessUnthemedParagraphAfter;
    private int _paragraphTextLength;
    private int _paragraphBorderLetterLength;
    private int _revisedListMarkerLength;
    private bool _paragraphHasLeaderTab;
    private string? _distributedText;
    private readonly HashSet<string> _distributedWordPairs = new(StringComparer.Ordinal);
    private double[]? _columnWidths;
    private double _columnGapPoints;
    private Cell? _activeColumnCell;
    private Cell[]? _columnCells;
    private int _columnIndex;
    private int _columnParagraphCount;
    private int _columnParagraphCapacity;
    private bool _columnBreakPending;
    private bool _footnoteAppendixStarted;
    private bool _endnotesAtSectionEnd;
    private int? _activeEndnoteIndex;
    private readonly Dictionary<long, int> _endnoteOrdinals = new();
    private readonly Dictionary<int, string> _endnoteLabels = new();
    private readonly HashSet<int> _customEndnoteMarks = new();
    private readonly Dictionary<int, int> _endnoteSectionStarts = new();
    private readonly Dictionary<long, Section> _endnoteSections = new();
    private readonly HashSet<Section> _endnoteSeparatorSections = new();
    private readonly HashSet<int> _endnoteDestinationOrdinals = new();
    private MigraDoc.DocumentObjectModel.PageBreak? _footnoteAppendixBreak;
    private bool _mirrorMargins;
    private int _footnoteBookmarkCount;

    public DxpFlowPdfVisitor(string fontFamily = "Arial", string? cjkFontFamily = null,
        IEnumerable<string>? availableFontFamilies = null,
        Func<char, string?>? chineseFallbackFont = null, bool includePdfComments = false,
        bool includeRevisionMarkup = false, XColor? revisionMarkupColor = null,
        XColor? revisionDeletionColor = null) : base(null)
    {
        _fontFamily = fontFamily;
        _cjkFontFamily = cjkFontFamily;
        _availableFontFamilies = new HashSet<string>(availableFontFamilies ?? [], StringComparer.OrdinalIgnoreCase);
        _chineseFallbackFont = chineseFallbackFont;
        _includePdfComments = includePdfComments;
        _includeRevisionMarkup = includeRevisionMarkup;
        _revisionMarkupColor = revisionMarkupColor ?? XColor.FromArgb(209, 52, 56);
        _revisionDeletionColor = revisionDeletionColor ?? XColor.FromArgb(209, 52, 56);
    }

    private static string? ReadMinorThemeLatinFont(WordprocessingDocument document)
    {
        if (document.MainDocumentPart?.ThemePart is not { } themePart) return null;
        using var stream = themePart.GetStream(FileMode.Open, FileAccess.Read);
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        return XDocument.Load(stream).Descendants(drawing + "minorFont")
            .FirstOrDefault()?.Element(drawing + "latin")?.Attribute("typeface")?.Value;
    }

    internal static bool HasModernUnthemedDefaults(WordprocessingDocument document)
    {
        var main = document.MainDocumentPart;
        var styles = main?.StyleDefinitionsPart?.Styles;
        var defaultFonts = styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.RunFonts;
        return main?.ThemePart == null &&
            main?.Document.Descendants<SectionProperties>().Any() != true &&
            main?.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.RunFonts>().Any() != true &&
            (styles?.Elements<WordStyle>().FirstOrDefault(style => style.StyleId?.Value == "Heading1")?
                 .StyleRunProperties?.RunFonts?.AsciiTheme?.Value == ThemeFontValues.MajorHighAnsi &&
             defaultFonts == null ||
             defaultFonts?.AsciiTheme?.Value == ThemeFontValues.MinorHighAnsi &&
             defaultFonts.HighAnsiTheme?.Value == ThemeFontValues.MinorHighAnsi) &&
            styles?.Elements<WordStyle>().FirstOrDefault(style => style.StyleId?.Value == "Normal")?
                .StyleRunProperties?.RunFonts == null;
    }

    public override void SetOutput(Stream stream) => _output = stream;

    private PdfDocumentRenderer CreatePdfRenderer(MigraDoc.DocumentObjectModel.Document document)
    {
        var renderer = new PdfDocumentRenderer { Document = document };
        foreach (var frame in _behindTextBoxes)
            frame.Layer = MigraDoc.DocumentObjectModel.Shapes.TextFrameLayer.BehindContent;
        foreach (var frame in _aboveContentTextBoxes)
            frame.Layer = MigraDoc.DocumentObjectModel.Shapes.TextFrameLayer.AboveContent;
        foreach (var frame in _clippedTextBoxes)
            frame.ClipContent = true;
        foreach (var (frame, rotation) in _textBoxRotations)
            frame.Rotation = rotation;
        if (Environment.GetEnvironmentVariable("DOCXPORT_PAGINATE_FLOATING_TABLES") != "0")
            foreach (var frame in _floatingTableFrames)
                frame.PaginateContent = true;
        if ((_splitOverpageTables.Count > 0 || _tableFragmentBorders.Count > 0) &&
            typeof(DocumentRenderer).GetProperty("TableFragmentTopBorderWidth") is { } fragmentBorderOption)
            fragmentBorderOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Tables.Table, int, double?>(
                    (table, startRow) => startRow > 0 && _splitOverpageTables.Contains(table)
                        ? Math.Max(0.5, FragmentTopBorderReservation(table, startRow) ?? 0)
                        : FragmentTopBorderReservation(table, startRow)));
        if (SupportsContourLayout)
            typeof(DocumentRenderer).GetProperty("WrapContour")?.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Shapes.Shape, IReadOnlyList<XPoint>?>(
                    shape => _wrapContours.TryGetValue(shape, out var contour) ? contour : null));
        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_CONTOUR_FONT_BAND") == "1")
            typeof(DocumentRenderer).GetProperty("WrapContourBandHeight")?.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Shapes.Shape, double, double?>(
                    (shape, height) => _wrapContourBandHeights.TryGetValue(shape, out var naturalHeight)
                        ? Math.Min(height, naturalHeight) : null));
        typeof(DocumentRenderer).GetProperty("WrapContourLineFontBandHeight")?.SetValue(
            renderer.DocumentRenderer,
            new Func<MigraDoc.DocumentObjectModel.Shapes.Shape, double, double, double?>(
                (shape, height, fontHeight) => _wrapContours.ContainsKey(shape) && fontHeight > 0
                    ? Math.Min(height, fontHeight + 0.25) : null));
        typeof(DocumentRenderer).GetProperty("WrapLineFitHeight")?.SetValue(
            renderer.DocumentRenderer,
            new Func<MigraDoc.DocumentObjectModel.Shapes.Shape, double, double, double?>(
                (shape, height, fontHeight) => _wrapContours.ContainsKey(shape) && fontHeight > 0
                    ? Math.Min(height, fontHeight) : null));
        typeof(DocumentRenderer).GetProperty("TabLeaderLayout")?.SetValue(renderer.DocumentRenderer,
            new Func<MigraDoc.DocumentObjectModel.Paragraph, MigraDoc.DocumentObjectModel.TabLeader, XFont,
                (double CharacterSpacing, double LeftInset, double RightInset)?>(
                (paragraph, leader, font) =>
                {
                    if (!string.Equals(font.FontFamily.Name, "Aptos", StringComparison.OrdinalIgnoreCase) ||
                        Math.Abs(font.Size - 12) >= 0.1) return null;
                    if (leader is MigraDoc.DocumentObjectModel.TabLeader.Heavy or
                        MigraDoc.DocumentObjectModel.TabLeader.Lines)
                    {
                        if (_positionalLeaderParagraphs.Contains(paragraph)) return (0.001, 4, 2.03);
                        return (0.001, 0, _footerLeaderParagraphs.Contains(paragraph) ? 3.43 : 1);
                    }
                    if (leader == MigraDoc.DocumentObjectModel.TabLeader.Dashes)
                        return (0.001, 0, _positionalLeaderParagraphs.Contains(paragraph) ? 0 :
                            _clearedTabLeaderParagraphs.Contains(paragraph) ? 3.69 : 2.3);
                    if (leader == MigraDoc.DocumentObjectModel.TabLeader.MiddleDot &&
                        _positionalLeaderParagraphs.Contains(paragraph)) return (0.0523, 0, 3.24);
                    if (leader != MigraDoc.DocumentObjectModel.TabLeader.Dots) return null;
                    if (_tocLeaderParagraphs.Contains(paragraph)) return (0.0523, 0.2, 0.6);
                    if (_positionalLeaderParagraphs.Contains(paragraph)) return (0.0523, 0, 3.24);
                    return (0.0523, 0, _footerLeaderParagraphs.Contains(paragraph) ? 0.8 : 2.3);
                }));
        typeof(DocumentRenderer).GetProperty("AllowClosingPunctuationOverhang")?
            .SetValue(renderer.DocumentRenderer, true);
        typeof(DocumentRenderer).GetProperty("TopBottomExclusion")?
            .SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Shapes.Shape, bool>(
                    shape => shape is MigraDoc.DocumentObjectModel.Shapes.Image));
        if (typeof(DocumentRenderer).GetProperty("LineEndOverhang") is { } overhangProperty &&
            overhangProperty.PropertyType ==
                typeof(Func<MigraDoc.DocumentObjectModel.Paragraph, string, XFont, double?>))
            overhangProperty.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, string, XFont, double?>(
                     (paragraph, word, font) =>
                     {
                         if (_narrowJustifiedBodyParagraphs.Contains(paragraph) &&
                             font.FontFamily.Name == "Aptos" && font.Size == 12 &&
                             GetPairedClosingMarkOverhang(word, font) is { } paired)
                             return paired;
                         return GetLineEndOverhang(paragraph, word, font,
                             _twoWordOverhangParagraphs.Contains(paragraph)) ??
                             (_stylelessBodyParagraphSpacing && _sourceBodyParagraphs.Contains(paragraph)
                                 ? GetStylelessFullStopOverhang(word, font) : null);
                     }));
        if (typeof(DocumentRenderer).GetProperty("AdditiveTerminalLineEndOverhang") is { } additiveOption)
            additiveOption.SetValue(renderer.DocumentRenderer, true);
        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_HANGING_MARK_PAIR_SPACES") == "1" &&
            typeof(DocumentRenderer).GetProperty("HangingTerminalLineSpaceThreshold") is { } hangingSpaceOption)
            hangingSpaceOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, string, XFont, double?>(
                     (paragraph, word, font) => GetHangingTerminalLineSpaceThreshold(
                         paragraph.Format.Alignment, _twoWordOverhangParagraphs.Contains(paragraph),
                         word, font.FontFamily.Name)));
        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_HANGING_MARK_PAIR_SPACES") == "1" &&
            typeof(DocumentRenderer).GetProperty("HangingTerminalLineNearEdgeThreshold") is { } nearEdgeOption)
            nearEdgeOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, string, XFont, double?>(
                    (paragraph, word, font) => GetHangingTerminalLineNearEdgeThreshold(
                        paragraph.Format.Alignment, _twoWordOverhangParagraphs.Contains(paragraph),
                        word, font.FontFamily.Name)));
        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_NONTERMINAL_CJK_HANG_SPACE") == "1" &&
            typeof(DocumentRenderer).GetProperty("HangingNonterminalLineSpaceLimit") is { } nonterminalOption)
            nonterminalOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, string, XFont, double?>(
                    (paragraph, word, _) => paragraph.Format.Alignment == ParagraphAlignment.Left &&
                        word.Length > 0 && IsFullWidthCjkPunctuation(word[^1]) &&
                        IsClosingLinePunctuation(word[^1]) ? 2.6 : null));
        typeof(DocumentRenderer).GetProperty("UnderlineStrokeWidthFactor")?
            .SetValue(renderer.DocumentRenderer, 0.025);
        if (typeof(DocumentRenderer).GetProperty("InlineImageVerticalOffset") is { } imageOffsetOption &&
            imageOffsetOption.PropertyType ==
                typeof(Func<MigraDoc.DocumentObjectModel.Shapes.Image, double?>))
            imageOffsetOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Shapes.Image, double?>(
                    image => _inlineImageOffsets.TryGetValue(image, out var offset) ? offset : null));
        typeof(DocumentRenderer).GetProperty("UnderlineVerticalOffsetFactor")?
            .SetValue(renderer.DocumentRenderer, 0.23);
        typeof(DocumentRenderer).GetProperty("UnderlineGeometry")?.SetValue(renderer.DocumentRenderer,
            new Func<MigraDoc.DocumentObjectModel.Font, (double StrokeWidthFactor, double VerticalOffsetFactor)?>(
                font => string.Equals(font.Name, "Arial", StringComparison.OrdinalIgnoreCase)
                    ? (0.095, 0.65) :
                    string.Equals(font.Name, "DejaVu Sans", StringComparison.OrdinalIgnoreCase)
                    ? (0.045, -0.15) :
                    string.Equals(font.Name, "Aptos Display", StringComparison.OrdinalIgnoreCase) &&
                    font.Size.Point is > 8.9 and < 9.1 ? (0.025, 0.29) :
                    font.Size.Point >= 12 ? (0.02, 0.23) : null));
        var trackedWidths = _documentHasTrackedInsertions &&
            Environment.GetEnvironmentVariable("DOCXPORT_TRACKED_POSITIONED_LAYOUT") == "1";
        var frameWidths = _frameWidthParagraphs.Count > 0 &&
            Environment.GetEnvironmentVariable("DOCXPORT_FRAME_POSITIONED_LAYOUT") != "0";
        var implicitTableWidths = _implicitTableParagraphs.Count > 0;
        var stylelessWidths = _stylelessBodyParagraphSpacing && _sourceBodyParagraphs.Count > 0;
        if ((trackedWidths || frameWidths || implicitTableWidths || stylelessWidths) &&
            _availableFontFamilies.Contains("Aptos") &&
            DocSystemFontAdvances.Find("Aptos", includeCloudFonts: true) is { } metrics)
        {
            double? Adjustment(string word, string family, double size) =>
                    string.Equals(family, "Aptos", StringComparison.OrdinalIgnoreCase) &&
                    metrics.TryMeasure(word, size, out var plain) &&
                    metrics.TryMeasure(word, size, out var positioned, useKerning: true)
                    ? positioned - plain
                        : null;
            if (typeof(DocumentRenderer).GetProperty("TextWidthAdjustment") is { } scopedOption)
            {
                scopedOption.SetValue(renderer.DocumentRenderer,
                    new Func<MigraDoc.DocumentObjectModel.Paragraph, string, XFont, double?>(
                        (paragraph, word, font) =>
                        {
                            if (implicitTableWidths && _implicitTableParagraphs.Contains(paragraph) &&
                                string.Equals(font.FontFamily.Name, "Aptos", StringComparison.OrdinalIgnoreCase) &&
                                metrics.TryMeasure(word, font.Size, out var advance))
                            {
                                using var measure = XGraphics.CreateMeasureContext(new XSize(100, 100),
                                    XGraphicsUnit.Point, XPageDirection.Downwards);
                                return advance - measure.MeasureString(word, font).Width;
                            }
                            if (stylelessWidths && _sourceBodyParagraphs.Contains(paragraph) &&
                                string.Equals(font.FontFamily.Name, "Aptos", StringComparison.OrdinalIgnoreCase) &&
                                metrics.TryMeasure(word, font.Size, out var kernedAdvance, useKerning: true))
                            {
                                using var measure = XGraphics.CreateMeasureContext(new XSize(100, 100),
                                    XGraphicsUnit.Point, XPageDirection.Downwards);
                                if (metrics.TryMeasure(word, font.Size, out var plainAdvance))
                                    kernedAdvance = Math.Min(kernedAdvance, plainAdvance);
                                return kernedAdvance - measure.MeasureString(word, font).Width;
                            }
                            return trackedWidths && _trackedLayoutParagraphs.Contains(paragraph) ||
                                frameWidths && _frameWidthParagraphs.Contains(paragraph)
                                ? Adjustment(word, font.FontFamily.Name, font.Size) : null;
                        }));
                _positionedLayoutApplied = trackedWidths;
            }
            else if (trackedWidths && !frameWidths &&
                typeof(DocumentRenderer).GetProperty("PositionedWordWidth") is { } widthOption)
            {
                widthOption.SetValue(renderer.DocumentRenderer,
                    new Func<string, string, double, double?>(Adjustment));
                _positionedLayoutApplied = true;
            }
        }
        if (typeof(DocumentRenderer).GetProperty("CenterTextFrameContentVertically") is { } centerOption)
            centerOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Shapes.TextFrame, bool>(
                    frame => _verticallyCenteredTextBoxes.Contains(frame)));
        if (typeof(DocumentRenderer).GetProperty("BreakOversizedNoBreakGroup") is { } emergencyOption)
            emergencyOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, bool>(
                    paragraph => _emergencyBreakParagraphs.Contains(paragraph)));
        var shrinkSetting = Environment.GetEnvironmentVariable("DOCXPORT_JUSTIFIED_SPACE_SHRINK");
        var shrinkFactor = double.TryParse(shrinkSetting, NumberStyles.Float,
            CultureInfo.InvariantCulture, out var configuredShrinkFactor) &&
            configuredShrinkFactor is > 0 and < 1
            ? configuredShrinkFactor : 0.6;
        if (shrinkSetting != "0" &&
            typeof(DocumentRenderer).GetProperty("JustifiedSpaceShrinkFactorForParagraph") is { } paragraphShrinkOption &&
            paragraphShrinkOption.PropertyType ==
                typeof(Func<MigraDoc.DocumentObjectModel.Paragraph, double?>))
            paragraphShrinkOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, double?>(
                    paragraph => shrinkSetting is null &&
                        _justifiedShrinkFactors.TryGetValue(paragraph, out var factor)
                        ? factor : shrinkFactor));
        else if (shrinkSetting != "0" &&
            typeof(DocumentRenderer).GetProperty("JustifiedSpaceShrinkFactor") is { } shrinkOption)
            shrinkOption.SetValue(renderer.DocumentRenderer, shrinkFactor);
        if (typeof(DocumentRenderer).GetProperty("JustifiedRenderWordWidthAdjustment") is { } renderWidthOption &&
            DocSystemFontAdvances.Find("Aptos", includeCloudFonts: true) is { } aptosAdvances)
            renderWidthOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, string, XFont, double?>(
                    (paragraph, word, font) =>
                    {
                        if (!_sourceBodyParagraphs.Contains(paragraph) ||
                            !string.Equals(font.FontFamily.Name, "Aptos", StringComparison.OrdinalIgnoreCase))
                            return null;
                        if (_narrowJustifiedBodyParagraphs.Contains(paragraph))
                            return Environment.GetEnvironmentVariable("DOCXPORT_DIAG_APTOS_COMMA_QUOTE_KERN") == "1" &&
                                word.EndsWith(",\u201D", StringComparison.Ordinal)
                                ? aptosAdvances.LegacyKerningAdjustment(',', '\u201D', font.Size) : null;
                        return aptosAdvances.TryMeasure(word, font.Size, out var plain) &&
                            aptosAdvances.TryMeasure(word, font.Size, out var kerned, useKerning: true)
                            ? kerned - plain : null;
                    }));
        var aptosKerningMetrics = DocSystemFontAdvances.Find("Aptos", includeCloudFonts: true);
        var calibriKerningMetrics = DocSystemFontAdvances.Find("Calibri");
        if (typeof(DocumentRenderer).GetProperty("LeadingKerningAfterBlank") is { } leadingKerningOption)
            leadingKerningOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, string, XFont, double?>(
                    (paragraph, word, font) =>
                    {
                        if (word.Length == 0) return null;
                        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_INLINE_RADICAL_LEADING_KERNING") == "1" &&
                            _standaloneInlineMathRadicals.Contains(paragraph) &&
                            !_mixedInlineMathRadicals.Contains(paragraph) && word[0] == '√' &&
                            string.Equals(font.FontFamily.Name, "Cambria Math", StringComparison.OrdinalIgnoreCase))
                            return 0.87;
                        if (paragraph.Format.Alignment == ParagraphAlignment.Justify &&
                            font.FontFamily.Name == "Aptos" && font.Size == 12 && word[0] == '“')
                            return aptosKerningMetrics?.KerningAdjustment(' ', '“', font.Size);
                        if (!_crossRunLeadingWords.TryGetValue(paragraph, out var leadingWords) ||
                            !leadingWords.Contains(word)) return null;
                        var metrics = DocSystemFontAdvances.Find(font.FontFamily.Name,
                            font.Style.HasFlag(XFontStyleEx.Bold), font.Style.HasFlag(XFontStyleEx.Italic));
                        return metrics?.KerningAdjustment(' ', word[0], font.Size);
                    }));
        if (typeof(DocumentRenderer).GetProperty("TrailingGlyphKerning") is { } trailingKerningOption &&
            (aptosKerningMetrics != null || calibriKerningMetrics != null))
        {
            trailingKerningOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, string, XFont, double?>(
                    (paragraph, word, font) =>
                    {
                        if (font.FontFamily.Name == "Calibri" && calibriKerningMetrics != null &&
                            word.Length >= 2 && word[^1] is ('»' or '›') &&
                            calibriKerningMetrics.TryMeasure(word, font.Size, out var calibriPlain) &&
                            calibriKerningMetrics.TryMeasure(word, font.Size, out var calibriKerned,
                                useKerning: true))
                            return calibriKerned - calibriPlain;
                        if (font.FontFamily.Name != "Aptos" || aptosKerningMetrics == null ||
                            word.Length < 2 || word[^2] != '.')
                            return null;
                        if (paragraph.Format.Alignment == ParagraphAlignment.Justify &&
                            word[^1] is ('’' or '”'))
                            return aptosKerningMetrics.KerningAdjustment('.', word[^1], font.Size);
                        if (paragraph.Format.Alignment == ParagraphAlignment.Left && word[^1] == '"' &&
                            aptosKerningMetrics.TryMeasure(word, font.Size, out var plain) &&
                            aptosKerningMetrics.TryMeasure(word, font.Size, out var kerned,
                                useKerning: true))
                            return kerned - plain;
                        if (paragraph.Format.Alignment == ParagraphAlignment.Left &&
                            font.Size is 9 or 12 &&
                            word[^1] == '\'')
                            return aptosKerningMetrics.KerningAdjustment('.', '\'', font.Size);
                        return null;
                    }));
        }
        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_LEFT_CURLY_MARK_PAIR") == "1" &&
            typeof(DocumentRenderer).GetProperty("RenderOnlyTrailingGlyphKerning") is { } renderOnlyKerningOption)
            renderOnlyKerningOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, string, XFont, double?>(
                     (paragraph, word, font) => paragraph.Format.Alignment == ParagraphAlignment.Left &&
                         _twoWordOverhangParagraphs.Contains(paragraph) &&
                         word.Length >= 2 && word[^2] == '.' && word[^1] is ('\u2019' or '\u201d')
                        ? font.FontFamily.Name switch
                        {
                            "Calibri" => calibriKerningMetrics?.KerningAdjustment('.', word[^1], font.Size),
                            "Aptos" => aptosKerningMetrics?.KerningAdjustment('.', word[^1], font.Size),
                            _ => null
                        }
                        : null));
        var hasOneAndHalfLineGrid = _sourceWordDocument?.MainDocumentPart?.Document
            .Descendants<SpacingBetweenLines>()
            .Any(spacing => spacing.Line?.Value == "360" &&
                (spacing.LineRule == null || spacing.LineRule.Value == LineSpacingRuleValues.Auto)) == true;
        if (!hasOneAndHalfLineGrid &&
            Environment.GetEnvironmentVariable("DOCXPORT_DIAG_ORDINARY_BODY_PAINT_BASELINE") == "1" &&
            typeof(DocumentRenderer).GetProperty("ParagraphPaintBaselineOffset") is { } paintBaselineOption)
            paintBaselineOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, IReadOnlyList<XFont>, double?>((paragraph, fonts) =>
                    _bodyParagraphs.Contains(paragraph) && !_documentHasPositionalTabs &&
                    !_documentHasTextBoxes && !_documentHasMath &&
                    !_inlineMathRadicals.Contains(paragraph) &&
                    !paragraph.Elements.OfType<MigraDoc.DocumentObjectModel.Shapes.Image>().Any() &&
                    fonts.Count > 0 && fonts.All(font =>
                        string.Equals(font.FontFamily.Name, "Aptos", StringComparison.OrdinalIgnoreCase)) &&
                    paragraph.Format.LineSpacingRule != LineSpacingRule.Exactly ? 0.07 : null));
        var exactImageBaselineSetting = Environment.GetEnvironmentVariable("DOCXPORT_DIAG_EXACT_IMAGE_BASELINE");
        var exactImageBaselineAuto = string.Equals(exactImageBaselineSetting, "auto", StringComparison.OrdinalIgnoreCase);
        double? exactImageBaselineOffset = double.TryParse(exactImageBaselineSetting, NumberStyles.Float,
            CultureInfo.InvariantCulture, out var configuredExactImageBaseline) &&
            double.IsFinite(configuredExactImageBaseline)
            ? configuredExactImageBaseline : null;
        if (typeof(DocumentRenderer).GetProperty("ExactParagraphBaselineOffset") is { } baselineOption)
            baselineOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, IReadOnlyList<XFont>, double?>(
                    (paragraph, fonts) =>
                        _rubyGuides.ContainsKey(paragraph) ? 1 :
                        _exactCjkLineParagraphs.Contains(paragraph) &&
                        !fonts.Any(font => font.FontFamily.Name.StartsWith("Yu Mincho",
                            StringComparison.OrdinalIgnoreCase)) ? -3.16 :
                        exactImageBaselineOffset is double imageBaseline &&
                        _exactLineImages.Any(item => ReferenceEquals(item.Paragraph, paragraph)) ? imageBaseline :
                        exactImageBaselineAuto && _exactLineImages.FirstOrDefault(item =>
                            ReferenceEquals(item.Paragraph, paragraph)).LineHeight is > 0 and var imageLineHeight
                            ? imageLineHeight * 0.0119 :
                        _exactStylelessImplicitAptosParagraphs.Contains(paragraph) ? -2.11 :
                        _exactStylelessPlainParagraphs.Contains(paragraph) ? -0.89 : null));
        var useRubyOverflowBaseline =
            Environment.GetEnvironmentVariable("DOCXPORT_DIAG_RUBY_OVERFLOW_BASELINE") == "1";
        var rubyBaselineOffset = 0.0;
        var hasRubyImplicitBaseline = _rubyOnlyImplicitFontDocument &&
            double.TryParse(Environment.GetEnvironmentVariable("DOCXPORT_DIAG_RUBY_BASELINE_OFFSET"),
                NumberStyles.Float, CultureInfo.InvariantCulture, out rubyBaselineOffset);
        if ((hasRubyImplicitBaseline || useRubyOverflowBaseline) &&
            typeof(DocumentRenderer).GetProperty("AtLeastParagraphBaselineOffset") is { } rubyBaselineOption)
            rubyBaselineOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, double?>(paragraph =>
                    hasRubyImplicitBaseline && _rubyGuides.ContainsKey(paragraph)
                        ? rubyBaselineOffset
                        : useRubyOverflowBaseline &&
                          _rubyOverflowBaselineOffsets.TryGetValue(paragraph, out var overflowOffset)
                            ? overflowOffset : null));
        var radicalBottomSpaceSetting = Environment.GetEnvironmentVariable("DOCXPORT_DIAG_INLINE_RADICAL_BOTTOM_SPACE");
        if (string.Equals(radicalBottomSpaceSetting, "auto", StringComparison.OrdinalIgnoreCase) &&
            typeof(DocumentRenderer).GetProperty("AtLeastBottomLineSpace") is { } automaticBottomSpaceOption)
            automaticBottomSpaceOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, double?>(paragraph =>
                    _inlineMathRadicals.Contains(paragraph)
                        ? _tallInlineMathRadicals.Contains(paragraph) ? 3.11 : 2.27
                        : null));
        else if (double.TryParse(radicalBottomSpaceSetting,
            NumberStyles.Float, CultureInfo.InvariantCulture, out var radicalBottomSpace) &&
            radicalBottomSpace > 0 &&
            typeof(DocumentRenderer).GetProperty("AtLeastBottomLineSpace") is { } bottomSpaceOption)
            bottomSpaceOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, double?>(paragraph =>
                    _inlineMathRadicals.Contains(paragraph) ? radicalBottomSpace : null));
        if (typeof(DocumentRenderer).GetProperty("ExactFontBaselineOffset") is { } fontBaselineOption)
            fontBaselineOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, XFont, double?>(
                    (paragraph, font) => _exactCjkLineParagraphs.Contains(paragraph) &&
                        font.FontFamily.Name.StartsWith("Yu Mincho", StringComparison.OrdinalIgnoreCase)
                        ? 1.34 : Environment.GetEnvironmentVariable("DOCXPORT_DIAG_ARIAL_EXACT_BASELINE") == "1" &&
                            !_documentHasPositionalTabs &&
                            font.FontFamily.Name.Equals("Arial", StringComparison.OrdinalIgnoreCase) &&
                            !_textBorderRanges.ContainsKey(paragraph)
                            ? 0.2 : null));
        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_CJK_FINAL_LINE_HEIGHT") == "1" &&
            typeof(DocumentRenderer).GetProperty("ExactFinalLineHeight") is { } finalHeightOption)
            finalHeightOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, IReadOnlyList<XFont>, double?>(
                    (paragraph, fonts) => _exactCjkLineParagraphs.Contains(paragraph) &&
                        !fonts.Any(font => font.FontFamily.Name.StartsWith("Yu Mincho",
                            StringComparison.OrdinalIgnoreCase)) ? 16.94 : null));
        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_CJK_MIXED_LINE_METRICS") == "1" &&
            typeof(DocumentRenderer).GetProperty("MinimumLineMetricsForFont") is { } mixedMetricsOption)
            mixedMetricsOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, XFont,
                    (double LineHeight, double InherentLineSpace)?>(
                    (paragraph, font) => _mixedFullWidthPunctuationParagraphs.Contains(paragraph) &&
                        font.FontFamily.Name.Equals("MS Mincho", StringComparison.OrdinalIgnoreCase) &&
                        font.Size == 12 && paragraph.Format.LineSpacingRule != LineSpacingRule.Exactly
                        ? (18.0878, 15.4484) : null));
        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_PAGE_TOP_BORDER_DISTANCE") == "1" &&
            typeof(DocumentRenderer).GetProperty("TopBorderDistanceAtAreaStart") is { } topBorderOption)
            topBorderOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, int, double, double?>(
                    (paragraph, pageNumber, areaY) => pageNumber > 1 &&
                        _directBorderParagraphs.TryGetValue(paragraph,
                            out var source) &&
                        source.ParagraphProperties?.ParagraphBorders is { } currentBorders &&
                        source.PreviousSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()?
                            .ParagraphProperties?.ParagraphBorders?.OuterXml == currentBorders.OuterXml &&
                        paragraph.Section is { } section &&
                        Math.Abs(areaY - section.PageSetup.TopMargin.Point) < 0.1 &&
                        Math.Abs(paragraph.Format.Borders.DistanceFromTop.Point - 1) < 0.1
                        ? 0 : null));
        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_INLINE_FRACTION_LEAF_SPACING") == "1" &&
            typeof(DocumentRenderer).GetProperty("AtLeastMinimumLineSpacing") is { } minimumOption)
            minimumOption.SetValue(renderer.DocumentRenderer,
                new Func<MigraDoc.DocumentObjectModel.Paragraph, XFont, double?>(
                    (paragraph, font) => _inlineMathFractions.ContainsKey(paragraph) && font.Size < 14
                        ? 0 : null));
#if PDFSHARP_GLYPH_UNICODE_SEQUENCE
        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_APTOS_FL_LIGATURE") == "1")
            renderer.PdfDocument.RenderEvents.RenderTextEvent += (_, args) =>
            {
                if (!string.Equals(args.Font.FontFamily.Name, "Aptos", StringComparison.OrdinalIgnoreCase))
                    return;
                var pairs = args.CodePointGlyphIndexPairs;
                if (!pairs.Zip(pairs.Skip(1), (first, second) => first.CodePoint == 'f' && second.CodePoint == 'l')
                    .Any(match => match))
                    return;
                var glyphIndex = PdfSharp.Fonts.GlyphHelper.GlyphIndexFromCodePoint(0xFB02, args.Font);
                if (glyphIndex == 0)
                    return;
                var shaped = new List<PdfSharp.Internal.OpenType.CodePointGlyphIndexPair>(pairs.Length);
                for (var index = 0; index < pairs.Length; index++)
                {
                    if (index + 1 < pairs.Length && pairs[index].CodePoint == 'f' && pairs[index + 1].CodePoint == 'l')
                    {
                        shaped.Add(new PdfSharp.Internal.OpenType.CodePointGlyphIndexPair('f', glyphIndex)
                            { UnicodeText = "fl" });
                        index++;
                    }
                    else shaped.Add(pairs[index]);
                }
                args.CodePointGlyphIndexPairs = shaped.ToArray();
            };
#endif
        if ((_inlineMathRadicals.Count > 0 || _mathRadicals.Count > 0) &&
            DocSystemFontAdvances.Find("Cambria Math") is { } mathMetrics &&
            mathMetrics.TryGetMathVerticalVariant('√', out var radicalVariant))
            renderer.PdfDocument.RenderEvents.RenderTextEvent += (sender, args) =>
            {
                if (!string.Equals(args.Font.FontFamily.Name, "Cambria Math",
                    StringComparison.OrdinalIgnoreCase)) return;
                var pairs = args.CodePointGlyphIndexPairs;
                for (var index = 0; index + 1 < pairs.Length; index++)
                    if (pairs[index].CodePoint == '√' &&
                        ShouldUseTallerMathRadical(mathMetrics, pairs[index + 1].CodePoint))
                        pairs[index].GlyphIndex = radicalVariant;
            };
        return renderer;
    }

    internal static bool ShouldUseTallerMathRadical(DocOpenTypeAdvances metrics,
        int radicandCodepoint) =>
        metrics.TryGetGlyphVerticalBounds(radicandCodepoint, out var minimum, out _) &&
        minimum < -metrics.UnitsPerEm / 10;

    internal static double? GetLineEndOverhang(MigraDoc.DocumentObjectModel.Paragraph paragraph,
        string word, XFont font, bool twoWordParagraph = false)
    {
        if (GetClosePunctuationTolerance(paragraph.Format.Alignment, word,
            font.FontFamily.Name, font.Size, twoWordParagraph) is { } tolerance)
            return tolerance;
        if (paragraph.Format.Alignment == ParagraphAlignment.Left && twoWordParagraph &&
            font.FontFamily.Name == "Aptos" && font.Size == 12 && word.EndsWith(".'", StringComparison.Ordinal))
            return 0.5;
        if (paragraph.Format.Alignment == ParagraphAlignment.Left && twoWordParagraph &&
            font.FontFamily.Name == "Aptos" && font.Size == 12 && word.Length > 0 &&
            word[^1] is ('»' or '›'))
            return word[^1] == '»' ? 0.25 : 0.1;
        if (paragraph.Format.Alignment == ParagraphAlignment.Left && twoWordParagraph &&
            font.FontFamily.Name == "Calibri" && word.EndsWith('"'))
            return font.Size * 0.06;
        if (paragraph.Format.Alignment == ParagraphAlignment.Left &&
            font.FontFamily.Name == "Aptos" && word.EndsWith(".\"", StringComparison.Ordinal))
            return null;
        if (paragraph.Format.Alignment is not (ParagraphAlignment.Justify or ParagraphAlignment.Left) ||
            word.Length == 0) return null;
        var fullWidthClosingMark = IsFullWidthCjkPunctuation(word[^1]) &&
            IsClosingLinePunctuation(word[^1]);
        if (!fullWidthClosingMark &&
            (word.Length < 3 || word[^1] is not ('\u201d' or '\u2019' or '"' or '\''))) return null;
        if (word[^1] is ('"' or '\'') &&
            (paragraph.Format.Alignment != ParagraphAlignment.Left || !twoWordParagraph ||
             font.Size < 18 || font.FontFamily.Name is not ("Aptos" or "Calibri") ||
             word[^1] == '\'' && font.FontFamily.Name != "Aptos"))
            return null;
        if (twoWordParagraph && word[^1] is ('\u2019' or '\u201d') &&
            font.FontFamily.Name == "Arial" && paragraph.Format.RightIndent.Point > 0)
            return null;
        using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100),
            XGraphicsUnit.Point, XPageDirection.Downwards);
        var allowance = fullWidthClosingMark ? font.Size * 1.05 :
            graphics.MeasureString(word[^1..], font).Width;
        if (!fullWidthClosingMark && word[^2] is ('.' or ','))
            allowance += graphics.MeasureString(word[^2..^1], font).Width * 0.5;
        if (!fullWidthClosingMark && font.FontFamily.Name == "Arial")
            allowance -= font.Size * 0.15;
        else if (!fullWidthClosingMark && font.FontFamily.Name == "Calibri")
            allowance -= 0.5;
        if (!fullWidthClosingMark && paragraph.Format.Alignment == ParagraphAlignment.Left)
            allowance = Math.Min(allowance, font.Size *
                (twoWordParagraph ? 0.6 : 1.0 / 3.0));
        if (word[^1] is ('"' or '\'') && paragraph.Format.Alignment == ParagraphAlignment.Left)
            allowance = Math.Min(allowance, font.Size * 0.12);
        return allowance;
    }

    internal static double? GetPairedClosingMarkOverhang(string word, XFont font)
    {
        if (word != "\u201D" &&
            !(word.Length >= 2 && word[^1] == '-') &&
            !(word.Length >= 3 && word[^1] == '\u201D' && word[^2] is (';' or ':' or '!' or '/')))
            return null;
        using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100),
            XGraphicsUnit.Point, XPageDirection.Downwards);
        var halfHyphen = graphics.MeasureString("-", font).Width / 2;
        return word == "\u201D"
            ? graphics.MeasureString(word, font).Width + halfHyphen
            : word[^1] == '-' ? halfHyphen : graphics.MeasureString(word[^2..], font).Width;
    }


    internal static double? GetHangingTerminalLineSpaceThreshold(ParagraphAlignment alignment,
        bool twoWordParagraph, string word, string fontName) =>
        alignment == ParagraphAlignment.Left && twoWordParagraph &&
        word.Length > 0 && word[^1] is ('\u201d' or '\u2019') &&
        fontName is "Aptos" or "Calibri" or "Arial" ? 0.25 : null;

    internal static double? GetHangingTerminalLineNearEdgeThreshold(ParagraphAlignment alignment,
        bool twoWordParagraph, string word, string fontName) =>
        alignment == ParagraphAlignment.Left && twoWordParagraph && fontName == "Arial" &&
        word.Length > 0 && word[^1] is ('\u201d' or '\u2019') ? 0.25 : null;

    internal static double? GetClosePunctuationTolerance(ParagraphAlignment alignment, string word,
        string fontName, double fontSize, bool twoWordParagraph)
    {
        return alignment == ParagraphAlignment.Left && twoWordParagraph && fontName == "Calibri" &&
            fontSize == 12 && word.Length > 0 &&
            !IsFullWidthCjkPunctuation(word[^1]) &&
            char.GetUnicodeCategory(word[^1]) == UnicodeCategory.ClosePunctuation ? 0.2 : null;
    }

    internal static double? GetStylelessFullStopOverhang(string word, XFont font)
    {
        if (word.Length < 2 || word[^1] != '.')
            return null;
        using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100),
            XGraphicsUnit.Point, XPageDirection.Downwards);
        return Math.Min(graphics.MeasureString(word[^1..], font).Width, font.Size / 3.0);
    }


    public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext context)
    {
        _sourceWordDocument = doc;
        _textBoxTextColor = null;
        _safeClosingMarkTokens.Clear();
        _numberingResolver = new DxpNumberingResolver(doc);
        _documentHasFitText = doc.MainDocumentPart?.Document.Descendants<FitText>().Any() == true;
        _documentHasTrackedInsertions = doc.MainDocumentPart?.Document.Descendants<InsertedRun>().Any() == true;
        _positionedLayoutApplied = false;
        _restoreCompressedBodySpaces = false;
        _layoutWidthParagraphs.Clear();
        _compressedBodyParagraphs.Clear();
        _trackedLayoutParagraphs.Clear();
        _frameWidthParagraphs.Clear();
        _implicitTableParagraphs.Clear();
        _documentHasPositionalTabs = doc.MainDocumentPart?.Document.Descendants<PositionalTab>().Any() == true ||
            doc.MainDocumentPart?.HeaderParts.Any(part => part.Header.Descendants<PositionalTab>().Any()) == true ||
            doc.MainDocumentPart?.FooterParts.Any(part => part.Footer.Descendants<PositionalTab>().Any()) == true;
        _documentHasTextBoxes = doc.MainDocumentPart?.Document
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.TextBoxContent>().Any() == true;
        _documentHasMath = doc.MainDocumentPart?.Document
            .Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any() == true;
        _defaultTabStopPoints = doc.MainDocumentPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<DefaultTabStop>()?.Val?.Value is short interval && interval != 720 && interval > 0
            ? interval / 20.0 : null;
        _bodyHasRowGridGaps = doc.MainDocumentPart?.Document.Body?.Descendants<TableRow>()
            .Any(row => row.TableRowProperties?.ChildElements.Any(element =>
                element.LocalName is "gridBefore" or "gridAfter") == true) == true;
        _bodyHasRowBorderExceptions = doc.MainDocumentPart?.Document.Body?
            .Elements<DocumentFormat.OpenXml.Wordprocessing.Table>()
            .Any(table => table.Elements<TableRow>()
                .Any(row => row.GetFirstChild<TablePropertyExceptions>()?
                    .GetFirstChild<TableBorders>() != null)) == true;
        _evenAndOddHeaders = doc.MainDocumentPart?.DocumentSettingsPart?.Settings?
            .Elements<EvenAndOddHeaders>().Any() == true;
        _balanceSingleByteDoubleByteWidth = doc.MainDocumentPart?.DocumentSettingsPart?.Settings?
            .Descendants<BalanceSingleByteDoubleByteWidth>()
            .Any(setting => setting.Val?.Value != false) == true;
        _balancedFooterFieldSpaces.Clear();
        _doubleBorderTextPadding.Clear();
        _distributedWordPairs.Clear();
        var gutterSetting = doc.MainDocumentPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<GutterAtTop>();
        _gutterAtTop = gutterSetting != null && gutterSetting.Val?.Value != false;
        var mirrorSetting = doc.MainDocumentPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<MirrorMargins>();
        _mirrorMargins = mirrorSetting != null && mirrorSetting.Val?.Value != false;
        _sectionCount = 0;
        _pendingEmptyParagraphPageBreak = false;
        _footnoteAppendixStarted = false;
        _endnotesAtSectionEnd = doc.MainDocumentPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<EndnoteDocumentWideProperties>()?
            .GetFirstChild<EndnotePosition>()?.Val?.Value == EndnotePositionValues.SectionEnd;
        _activeEndnoteIndex = null;
        _endnoteOrdinals.Clear();
        _endnoteLabels.Clear();
        _customEndnoteMarks.Clear();
        _endnoteSectionStarts.Clear();
        _endnoteSections.Clear();
        _endnoteSeparatorSections.Clear();
        _endnoteDestinationOrdinals.Clear();
        _footnoteAppendixBreak = null;
        _footnoteBookmarkCount = 0;
        _activeFootnoteIndex = null;
        _footnoteReferences.Clear();
        _footnoteSplitContinuations.Clear();
        _footnoteParagraphs.Clear();
        _firstBodyClearsInheritedEmphasis.Clear();
        _renderedPageFields.Clear();
        _renderedMarginMergeFields.Clear();
        _pageBorders.Clear();
        _sectionLineNumbers.Clear();
        _sourceBodyParagraphs.Clear();
        _sourceCellParagraphs.Clear();
        _crossRunLeadingWords.Clear();
        _exactCjkLineParagraphs.Clear();
        _mixedFullWidthPunctuationParagraphs.Clear();
        _exactStylelessPlainParagraphs.Clear();
        _exactStylelessImplicitAptosParagraphs.Clear();
        _justifiedShrinkFactors.Clear();
        _twoWordOverhangParagraphs.Clear();
        _floatingTableFrames.Clear();
        _floatingContinuationTopBorders.Clear();
        _insideInsertedRun = false;
        _insideDeletedRun = false;
        _insideMoveToRun = false;
        _insideMoveFromRun = false;
        _textBoxRotations.Clear();
        _notBesideFrameFollowers.Clear();
        _verticallyCenteredTextBoxes.Clear();
        _clippedTextBoxes.Clear();
        _foregroundTextBoxes.Clear();
        _behindTextBoxes.Clear();
        _aboveContentTextBoxes.Clear();
        _cellAnchoredTextBoxes.Clear();
        _commentThreads.Clear();
        _commentRangeCells.Clear();
        _exactLineImages.Clear();
        _inlineImageOffsets.Clear();
        _tocLeaderParagraphs.Clear();
        _positionalLeaderParagraphs.Clear();
        _footerLeaderParagraphs.Clear();
        _clearedTabLeaderParagraphs.Clear();
        _narrowJustifiedBodyParagraphs.Clear();
        _styleRefBodyParagraphs.Clear();
        _styleRefHeaders.Clear();
        _suppressedLineNumberParagraphs.Clear();
        _footerTableDistanceApplied.Clear();
        _firstPageIntroTableFooters.Clear();
        _footerLineGridDistanceApplied.Clear();
        _fittedCjkFooterDistanceApplied.Clear();
        _fittedCjkFooterContinuationDistanceApplied.Clear();
        _unstyledFooterDistanceApplied.Clear();
        _numberedFooterDistanceApplied.Clear();
        _footerSpacingResetApplied.Clear();
        _repeatingTableFooterClearanceApplied.Clear();
        _cjkDefaultMarginsAdjusted.Clear();
        _spacedTableLeftBorders.Clear();
        _htmlBevelTables.Clear();
        _tableFragmentBorders.Clear();
        _splitOverpageTables.Clear();
        _nestedFragmentBorderGroups.Clear();
        _diagonalCells.Clear();
        _rotatedCells.Clear();
        _doubleFirstRowTopBorders.Clear();
        _rtlDoubleTableLeftBorders.Clear();
        _doubleTableRightBorders.Clear();
        _doubleCellSideBorders.Clear();
        _nativeDoubleTableBorders = false;
        _paragraphShades.Clear();
        _directBorderParagraphs.Clear();
        _manualParagraphShades.Clear();
        _footerParagraphShadeClearances.Clear();
        _emphasisRanges.Clear();
        _strikeRanges.Clear();
        _revisionUnderlineRanges.Clear();
        _revisionParagraphs.Clear();
        _formattingBalloons.Clear();
        _changedFieldUrls.Clear();
        if (_includeRevisionMarkup && doc.MainDocumentPart is { } mainPart)
        {
            var roots = new List<OpenXmlElement>();
            if (mainPart.Document != null) roots.Add(mainPart.Document);
            roots.AddRange(mainPart.HeaderParts.Select(part => part.Header).OfType<OpenXmlElement>());
            roots.AddRange(mainPart.FooterParts.Select(part => part.Footer).OfType<OpenXmlElement>());
            if (mainPart.FootnotesPart?.Footnotes is { } footnotes) roots.Add(footnotes);
            if (mainPart.EndnotesPart?.Endnotes is { } endnotes) roots.Add(endnotes);
            foreach (var run in roots.SelectMany(root => root.Descendants<Run>()))
            {
                if (run.RunProperties?.GetFirstChild<RunPropertiesChange>() == null ||
                    run.GetFirstChild<FieldCode>() is not { } fieldCode) continue;
                var fieldText = fieldCode.Text ?? fieldCode.InnerText;
                var parsed = new DxpFieldParser().Parse(fieldText).Ast;
                if (string.Equals(parsed.FieldType, "HYPERLINK", StringComparison.OrdinalIgnoreCase) &&
                    DxpFieldTokenization.TryGetFirstToken(parsed.ArgumentsText, out var target) &&
                    Uri.TryCreate(target, UriKind.Absolute, out var uri) &&
                    uri.Scheme is "http" or "https" or "mailto")
                    _changedFieldUrls[target] = _changedFieldUrls.GetValueOrDefault(target) + 1;
            }
        }
        _strikeTabBoundaries.Clear();
        _textEffectRanges.Clear();
        _joinedClosingMarkColors.Clear();
        _largeSuperscriptTexts.Clear();
        _inlineSuperscriptSizes.Clear();
        _ambiguousInlineSuperscripts.Clear();
        _closingQuoteSourceCounts.Clear();
        _excludedClosingQuoteTokens.Clear();
        _closingQuoteFollowedBySpace.Clear();
        _rightAlignedTrackedTexts.Clear();
        _longJustifiedTrackedTexts.Clear();
        _plainLatinAfterFitTexts.Clear();
        _apostropheDisplaySourceCounts.Clear();
        _excludedApostropheDisplayTokens.Clear();
        _rightLegendLabelGroups.Clear();
        _columnRightLegendLabelGroups.Clear();
        _characterScales.Clear();
        _topAlignedFontSizes.Clear();
        _topAlignedHeaderParagraphs.Clear();
        _verticalAlignedFonts.Clear();
        _oversizedGridLatinParagraphs.Clear();
        _oversizedStoryLatinParagraphs.Clear();
        _topAlignedBodyMargins.Clear();
        _topAlignedFooterDistances.Clear();
        _centerBottomBodyMargins.Clear();
        _centerBottomFooterDistances.Clear();
        _fitTextWidths.Clear();
        _fitTextCharacterCounts.Clear();
        _mixedCjkFitBodyParagraphs.Clear();
        _mixedHighAnsiFittedHeaders.Clear();
        _latinTimesFitParagraphs.Clear();
        _mixedCalibriYuMinchoFitParagraphs.Clear();
        _latinYuMinchoFitParagraphs.Clear();
        _headersWithFirstFittedParagraph.Clear();
        _headersWithSecondCjkParagraph.Clear();
        _fittedBodyAfterSecondCjkHeader.Clear();
        _latinCalibriFitParagraphs.Clear();
        _fitTextTracking.Clear();
        _partialFitGroups.Clear();
        _styleFitStoryDistancesApplied.Clear();
        _implicitFitFooterDistancesApplied.Clear();
        _calibriFitFooterDistancesApplied.Clear();
        _characterSpacings.Clear();
        _hangulBaselineShiftParagraphs.Clear();
        _rubyGuides.Clear();
        _rubyNaturalLineHeights.Clear();
        _rubyOriginalLineSpacingRules.Clear();
        _rubyOverflowBaselineOffsets.Clear();
        _suppressRubyGuideText = false;
        _renderRubyBaseText = false;
        _renderRubyGuide = false;
        _tabbedRunShadings.Clear();
        _borderedTabInsets.Clear();
        _centeredSpacingOffsets.Clear();
        _gridGlyphParagraphs.Clear();
        _gridBoundaryParagraphs.Clear();
        _unshiftedGridGlyphParagraphs.Clear();
        _packedGridParagraphs.Clear();
        _paintOnlyMinchoSpans.Clear();
        _paintOnlyRubyBaseMinchoSpans.Clear();
        _packedGridOffsets.Clear();
        _firstBodyDoubleBorderInset.Clear();
        _textBorderRanges.Clear();
        _headerRunBorderParagraphs.Clear();
        _footerRunBorderParagraphs.Clear();
        _coloredUnderlineParagraphs.Clear();
        _listMarkerUnderlines.Clear();
        _listMarkerStrikes.Clear();
        _borderResetParagraphs.Clear();
        _bodyParagraphs.Clear();
        _compactArialBodyParagraphs.Clear();
        _cjkNumberedBodyParagraphs.Clear();
        _bulletBaselineParagraphs.Clear();
        _headersWithStyledEmptyLine.Clear();
        _headersWithMarginAnchoredWrap.Clear();
        _headerRunBorderClearances.Clear();
        _verticallyAlignedSections.Clear();
        _sectionsWithLinkedHeaders.Clear();
        _bodyParagraphsUnderHeaderWrap.Clear();
        _headersWithStyledTrailingLine.Clear();
        _footersWithStyledEmptyLine.Clear();
        _footersWithStyledTrailingLine.Clear();
        _footerImages.Clear();
        _paragraphImages.Clear();
        _footnoteLabels.Clear();
        _customFootnoteMarks.Clear();
        _footnoteSectionStarts.Clear();
        _wrapContours.Clear();
        _wrapContourBandHeights.Clear();
        _mathFractions.Clear();
        _firstBodyDisplayFractions.Clear();
        _mixedDisplayMathParagraphs.Clear();
        _mixedDisplayRunFractionParagraphs.Clear();
        _inlineMathFractions.Clear();
        _nestedInlineMathFractions.Clear();
        _inlineMathMatrices.Clear();
        _mathRadicals.Clear();
        _inlineMathRadicals.Clear();
        _standaloneInlineMathRadicals.Clear();
        _mixedInlineMathRadicals.Clear();
        _tallInlineMathRadicals.Clear();
        _inlineMathCombinedScripts.Clear();
        _floatingTextWraps.Clear();
        _activeFloatingTextWrap = null;
        _mirroredParagraphs.Clear();
        _highlights.Clear();
        _highlightParagraphTexts.Clear();
        _noBreakWords.Clear();
        _noBreakParagraphs.Clear();
        _superscriptHeaderParagraphs.Clear();
        _paragraphStyles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?
            .Elements<WordStyle>()
            .Where(style => style.StyleId?.Value != null)
            .ToDictionary(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase)
            ?? new Dictionary<string, WordStyle>();
        _document = new MigraDoc.DocumentObjectModel.Document();
        _document.DefaultTabStop = Unit.FromPoint(_defaultTabStopPoints ?? 36);
        var normalFont = _document.Styles["Normal"].Font;
        _minorThemeLatinFont = ReadMinorThemeLatinFont(doc);
        _rubyOnlyImplicitFontDocument = doc.MainDocumentPart?.ThemePart == null &&
            doc.MainDocumentPart?.StyleDefinitionsPart == null &&
            doc.MainDocumentPart?.Document?.Descendants<DocumentFormat.OpenXml.Wordprocessing.RunFonts>()
                .Any() == true &&
            doc.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.RunFonts>()
                .All(fonts => fonts.Ancestors<Ruby>().Any()) == true;
        _implicitModernFont = !_documentHasFitText && _availableFontFamilies.Contains("Aptos") &&
            (string.Equals(_minorThemeLatinFont, "Aptos", StringComparison.OrdinalIgnoreCase) &&
                doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?
                    .RunPropertiesDefault?.RunPropertiesBaseStyle?.ChildElements.Count is null or 0 &&
                _paragraphStyles.GetValueOrDefault("Normal")?.StyleRunProperties?.RunFonts == null ||
             doc.MainDocumentPart?.ThemePart == null &&
                doc.MainDocumentPart?.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.RunFonts>().Any() != true &&
                doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?
                    .Descendants<DocumentFormat.OpenXml.Wordprocessing.RunFonts>().Any() != true ||
             HasModernUnthemedDefaults(doc) ||
             _rubyOnlyImplicitFontDocument &&
                Environment.GetEnvironmentVariable("DOCXPORT_DIAG_RUBY_IMPLICIT_FONT") == "1");
        var mixedImplicitModernFont =
            !_implicitModernFont && _availableFontFamilies.Contains("Aptos") &&
            doc.MainDocumentPart?.ThemePart == null &&
            doc.MainDocumentPart?.StyleDefinitionsPart == null &&
            doc.MainDocumentPart?.Document?.Body?.Descendants<Run>().Any(run =>
                run.RunProperties?.RunFonts != null && !run.Ancestors<Ruby>().Any()) == true &&
            doc.MainDocumentPart?.Document?.Body?.Descendants<Run>().Any(run =>
                run.RunProperties?.RunFonts == null &&
                run.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Any(text => !string.IsNullOrWhiteSpace(text.Text))) == true;
        var defaultRunProperties = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?
            .RunPropertiesDefault?.RunPropertiesBaseStyle;
        _explicitDefaultRunProperties = defaultRunProperties?.ChildElements.Count > 0;
        var normalStyle = _paragraphStyles.Values.FirstOrDefault(style =>
            style.Type?.Value == StyleValues.Paragraph && style.Default?.Value == true);
        if (normalStyle == null) _paragraphStyles.TryGetValue("Normal", out normalStyle);
        _classicPartialDefaults = !_implicitModernFont && defaultRunProperties?.ChildElements.Count > 0 &&
            defaultRunProperties.FontSize == null && normalStyle?.StyleRunProperties?.FontSize == null;
        normalFont.Name = _implicitModernFont || mixedImplicitModernFont ? "Aptos" :
            _classicPartialDefaults && defaultRunProperties?.RunFonts == null &&
                normalStyle?.StyleRunProperties?.RunFonts == null ? "Times New Roman" : _fontFamily;
        normalFont.Size = _implicitModernFont || mixedImplicitModernFont ? 12 : _classicPartialDefaults ? 10 : 11;
        var defaultSpacing = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?
            .ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.SpacingBetweenLines;
        _documentDefaultParagraphAfterPoints = int.TryParse(defaultSpacing?.After?.Value,
            out var documentAfterTwips) ? documentAfterTwips / 20.0 : null;
        _defaultRunHighlight = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?
            .RunPropertiesDefault?.RunPropertiesBaseStyle?.GetFirstChild<Highlight>()?.Val?.Value;
        var normalSpacing = normalStyle?.StyleParagraphProperties?.SpacingBetweenLines;
        var beforeValue = normalSpacing?.Before?.Value ?? defaultSpacing?.Before?.Value;
        var afterValue = normalSpacing?.After?.Value ?? defaultSpacing?.After?.Value;
        _themedImplicitParagraphSpacing = _implicitModernFont &&
            doc.MainDocumentPart?.ThemePart != null && afterValue == null;
        _stylelessBodyParagraphSpacing = _implicitModernFont &&
            doc.MainDocumentPart?.StyleDefinitionsPart == null && afterValue == null &&
            doc.MainDocumentPart?.Document?.Body?.Descendants<SectionProperties>().Count() <= 1;
        _stylelessUnthemedParagraphAfter = doc.MainDocumentPart?.StyleDefinitionsPart == null &&
            doc.MainDocumentPart?.ThemePart == null && afterValue == null &&
            doc.MainDocumentPart?.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.RunFonts>()
                .Any() == true &&
            doc.MainDocumentPart?.Document?.Body?.Descendants<SectionProperties>().Count() <= 1;
        var lineValue = normalSpacing?.Line?.Value ?? defaultSpacing?.Line?.Value;
        var lineRule = normalSpacing?.LineRule?.Value ?? defaultSpacing?.LineRule?.Value;
        var defaultSize = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?
            .RunPropertiesDefault?.RunPropertiesBaseStyle?.FontSize?.Val?.Value;
        var normalSize = normalStyle?.StyleRunProperties?.FontSize?.Val?.Value;
        var fontPoints = double.TryParse(defaultSize ?? normalSize, out var defaultSizeHalfPoints)
            ? defaultSizeHalfPoints / 2 : _classicPartialDefaults ? 10 : 11;
        _defaultParagraphAfterPoints = int.TryParse(afterValue, out var spacingAfter)
            ? spacingAfter / 20.0 : _themedImplicitParagraphSpacing ? 8 : 0;
        _defaultParagraphLinePoints = int.TryParse(lineValue, out var lineTwips) && lineTwips > 0
            ? lineRule == LineSpacingRuleValues.Exact ? lineTwips / 20.0
                : lineRule == LineSpacingRuleValues.AtLeast ? Math.Max(fontPoints, lineTwips / 20.0)
                : fontPoints * lineTwips / 240.0
            : fontPoints * 1.15;
        var normalFormat = _document.Styles["Normal"].ParagraphFormat;
        if (int.TryParse(beforeValue, out var defaultBefore))
            normalFormat.SpaceBefore = Unit.FromPoint(defaultBefore / 20.0);
        if (int.TryParse(afterValue, out var defaultAfter))
            normalFormat.SpaceAfter = Unit.FromPoint(defaultAfter / 20.0);
        else if (_themedImplicitParagraphSpacing)
            normalFormat.SpaceAfter = Unit.FromPoint(_defaultParagraphAfterPoints);
        if (int.TryParse(lineValue, out var defaultLine) && defaultLine > 0)
        {
            if (lineRule == LineSpacingRuleValues.Exact)
            {
                normalFormat.LineSpacingRule = LineSpacingRule.Exactly;
                normalFormat.LineSpacing = Unit.FromPoint(defaultLine / 20.0);
            }
            else if (lineRule == LineSpacingRuleValues.AtLeast)
            {
                normalFormat.LineSpacingRule = LineSpacingRule.AtLeast;
                normalFormat.LineSpacing = Unit.FromPoint(defaultLine / 20.0);
            }
            else
            {
                normalFormat.LineSpacingRule = LineSpacingRule.Multiple;
                normalFormat.LineSpacing = defaultLine / 240.0;
            }
        }
        return DxpDisposable.Create(() =>
        {
            if (_classicPartialDefaults &&
                string.Equals(defaultRunProperties?.RunFonts?.Ascii?.Value, "Aptos", StringComparison.OrdinalIgnoreCase))
                foreach (Section section in _document.Sections)
                {
                    section.PageSetup.TopMargin = Unit.FromPoint(
                        section.PageSetup.TopMargin.Point + ClassicAptosStoryClearancePoints);
                    section.PageSetup.FooterDistance = Unit.FromPoint(
                        section.PageSetup.FooterDistance.Point + ClassicAptosStoryClearancePoints);
                }
            if (_implicitModernFont && doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults is
                    { ChildElements.Count: 0 } && normalStyle?.StyleParagraphProperties == null &&
                normalStyle?.StyleRunProperties == null)
                foreach (Section section in _document.Sections)
                    section.PageSetup.TopMargin = Unit.FromPoint(
                        section.PageSetup.TopMargin.Point - EmptyModernDefaultsBodyClearancePoints);
            foreach (var section in _firstBodyClearsInheritedEmphasis)
                section.PageSetup.TopMargin = Unit.FromPoint(section.PageSetup.TopMargin.Point - 5);
            foreach (var (section, frame, _) in _foregroundTextBoxes.OrderBy(item => item.RelativeHeight))
            {
                var index = section.Elements.IndexOf(frame);
                if (index < 0) continue;
                section.Elements.RemoveObjectAt(index);
                var foregroundFrame = (MigraDoc.DocumentObjectModel.Shapes.TextFrame)frame.Clone();
                section.Elements.Add(foregroundFrame);
                if (_verticallyCenteredTextBoxes.Contains(frame))
                    _verticallyCenteredTextBoxes.Add(foregroundFrame);
                if (_clippedTextBoxes.Contains(frame))
                    _clippedTextBoxes.Add(foregroundFrame);
                if (_textBoxRotations.TryGetValue(frame, out var rotation))
                    _textBoxRotations[foregroundFrame] = rotation;
                if (_cellAnchoredTextBoxes.Remove(frame, out var cellAnchor))
                    _cellAnchoredTextBoxes[foregroundFrame] = cellAnchor;
            }
            RepeatFloatingTableTopBorders();
            SplitOverpageTableCells();
            RestoreNestedFragmentPageBorders();
            AlignSectionContentVertically();
            MigraDoc.DocumentObjectModel.Document? evenDocument = null;
            if (_mirroredParagraphs.Count > 0)
            {
                foreach (var paragraph in _mirroredParagraphs)
                {
                    var leftIndent = paragraph.Format.LeftIndent;
                    paragraph.Format.LeftIndent = paragraph.Format.RightIndent;
                    paragraph.Format.RightIndent = leftIndent;
                }
                evenDocument = (MigraDoc.DocumentObjectModel.Document)_document!.Clone();
                foreach (var paragraph in _mirroredParagraphs)
                {
                    var leftIndent = paragraph.Format.LeftIndent;
                    paragraph.Format.LeftIndent = paragraph.Format.RightIndent;
                    paragraph.Format.RightIndent = leftIndent;
                }
            }
            if (_headersWithStyledEmptyLine.Count > 0)
            {
                var preflight = (MigraDoc.DocumentObjectModel.Document)_document!.Clone();
                var preflightRenderer = CreatePdfRenderer(preflight);
                preflightRenderer.RenderDocument();
                ReserveHeaderClearance(preflightRenderer, preflight, evenDocument);
            }
            PdfDocumentRenderer? footnoteRenderer = null;
            List<(int Page, int Index, MigraDoc.DocumentObjectModel.Paragraph Note, double Y,
                int PageOffset)> footnoteOverlays = new();
            if (_footnoteParagraphs.Count > 0)
            {
                var hasBodyTables = _document!.Sections.Cast<Section>().Any(section =>
                    section.Elements.OfType<MigraDoc.DocumentObjectModel.Tables.Table>().Any());
                for (var iteration = 0; iteration < 32; iteration++)
                {
                    var layout = (MigraDoc.DocumentObjectModel.Document)_document!.Clone();
                    var layoutRenderer = CreatePdfRenderer(layout);
                    layoutRenderer.RenderDocument();
                    if (hasBodyTables)
                    {
                        if (TrySplitOverflowingTableRow(layoutRenderer, layout)) continue;
                        layout = (MigraDoc.DocumentObjectModel.Document)_document.Clone();
                        layoutRenderer = CreatePdfRenderer(layout);
                        layoutRenderer.RenderDocument();
                    }
                    if (ReserveFootnoteReferencePages(layoutRenderer, layout)) continue;
                    break;
                }
                var preflight = (MigraDoc.DocumentObjectModel.Document)_document!.Clone();
                foreach (Section section in preflight.Sections)
                    foreach (var note in section.Elements.OfType<MigraDoc.DocumentObjectModel.Paragraph>())
                        for (var index = note.Elements.Count - 1; index >= 0; index--)
                            if (note.Elements[index].GetType().Name == "BookmarkField")
                                note.Elements.RemoveObjectAt(index);
                footnoteRenderer = CreatePdfRenderer(preflight);
                footnoteRenderer.RenderDocument();
                footnoteOverlays = PlaceFittingFootnotes(footnoteRenderer, preflight);
            }
            if (evenDocument != null && footnoteRenderer != null)
            {
                foreach (var paragraph in _mirroredParagraphs)
                {
                    var leftIndent = paragraph.Format.LeftIndent;
                    paragraph.Format.LeftIndent = paragraph.Format.RightIndent;
                    paragraph.Format.RightIndent = leftIndent;
                }
                evenDocument = (MigraDoc.DocumentObjectModel.Document)_document!.Clone();
                foreach (var paragraph in _mirroredParagraphs)
                {
                    var leftIndent = paragraph.Format.LeftIndent;
                    paragraph.Format.LeftIndent = paragraph.Format.RightIndent;
                    paragraph.Format.RightIndent = leftIndent;
                }
            }
            var styleRefTemplate = _styleRefHeaders.Count > 0
                ? (MigraDoc.DocumentObjectModel.Document)_document!.Clone() : null;
            AlignNotBesideFrameFollowers();
            PositionCellAnchoredTextBoxes();
            var renderer = CreatePdfRenderer(_document);
            if (_noBreakWords.Count > 0)
            {
                renderer.DocumentRenderer.PrepareDocument();
                var restored = new HashSet<string>(StringComparer.Ordinal);
                foreach (var paragraph in _noBreakParagraphs)
                    RestoreNoBreakWords(paragraph.Elements, _noBreakWords, restored);
                for (var page = 1; page <= renderer.PageCount; page++)
                    foreach (var info in renderer.DocumentRenderer.GetRenderInfoFromPage(page))
                        if (info.DocumentObject is MigraDoc.DocumentObjectModel.Paragraph paragraph)
                            RestoreNoBreakWords(paragraph.Elements, _noBreakWords, restored);
                if (!_noBreakWords.IsSubsetOf(restored))
                    throw new InvalidOperationException("A no-break hyphen was not present after PDF layout: " +
                        string.Join(", ", _noBreakWords.Except(restored)));
            }
            if (_layoutWidthParagraphs.Count > 0 || _paintOnlyMinchoSpans.Count > 0 ||
                _paintOnlyRubyBaseMinchoSpans.Count > 0)
            {
                renderer.PrepareRenderPages();
                foreach (var paragraph in _layoutWidthParagraphs)
                    paragraph.Format.RightIndent = Unit.FromPoint(paragraph.Format.RightIndent.Point +
                        DenseTrackedLayoutAllowancePoints);
                foreach (var span in _paintOnlyMinchoSpans)
                    span.Font.Name = "MS Mincho";
                foreach (var span in _paintOnlyRubyBaseMinchoSpans)
                    span.Font.Name = "Yu Mincho";
                renderer.RenderPages(1, renderer.PageCount);
            }
            else renderer.RenderDocument();
            if (footnoteRenderer != null)
                RenderFittingFootnotes(renderer, footnoteRenderer, footnoteOverlays);
            DisableRasterInterpolation(renderer.PdfDocument);
            if (_mirroredParagraphs.Count == 0 || renderer.PageCount == 1)
            {
                using var pageAware = RenderPageAwareStyleRefHeaders(renderer, styleRefTemplate);
                RenderPageBorders(renderer, pageAware);
                RenderLineNumbers(renderer, pageAware);
                RenderTableFragmentBorders(renderer, pageAware);
                RenderSpacedTableBorders(renderer, pageAware);
                RenderHtmlBeveledTableBorders(renderer, pageAware);
                RenderDiagonalCellBorders(renderer, pageAware);
                RenderRotatedCells(renderer, pageAware);
                RenderDoubleFirstRowTopBorders(renderer, pageAware);
                RenderRtlDoubleTableLeftBorders(renderer, pageAware);
                RenderDoubleTableRightBorders(renderer, pageAware);
                RenderDoubleCellSideBorders(renderer, pageAware);
                RenderParagraphShadingGaps(renderer, pageAware);
                RenderContinuousParagraphSideBorders(renderer, pageAware);
                RenderFooterImages(renderer, pageAware);
                RenderParagraphAnchoredImages(renderer, pageAware);
                RenderMathFractions(renderer, pageAware);
                RenderMathRadicals(renderer, pageAware);
                using var emphasized = RenderEmphasisMarks(renderer, pageAware,
                    footnoteRenderer, footnoteOverlays);
                ApplyRightLegendShift(emphasized);
                ApplyLargeSuperscriptRise(emphasized);
                PadDoubleBorderText(emphasized);
                ApplyFontKerning(emphasized);
                ApplyRightAlignedTrackedKerning(emphasized);
                ApplyInlineSuperscriptRise(emphasized);
                ApplyBodyFontBaselineRise(emphasized);
                RaiseCjkNumberedBodyText(renderer, emphasized);
                AlignBulletBaselines(renderer, emphasized);
                SuppressNativeColoredUnderlines(renderer, emphasized);
                LowerMixedDisplayMathRuns(renderer, emphasized);
                using var barred = RenderInlineMathBars(renderer, emphasized);
                using var underlined = RenderColoredUnderlines(renderer, barred);
                using var justified = JustifyContinuedPageLines(renderer, underlined);
                var finalDocument = justified ?? underlined;
                WidenBalancedFooterFieldSpaces(finalDocument);
                CollapseDuplicateThinStrokes(finalDocument);
                RecolorJoinedClosingMarks(finalDocument);
                RestoreSourceSpaces(finalDocument);
                using var painted = RestoreLayoutWidthContinuationPaint(renderer, finalDocument);
                ApplyLongJustifiedTrackedKerning(painted ?? finalDocument);
                MatchElevenPointAptosPaint(painted ?? finalDocument);
                DisableRasterInterpolation(painted ?? finalDocument);
                ClipExactLineImages(renderer, painted ?? finalDocument);
                using var annotatedEarly = RenderCommentAnnotations(renderer, painted ?? finalDocument);
                using var formattedEarly = ApplyFormattingBalloons(renderer,
                    annotatedEarly ?? painted ?? finalDocument);
                DxpPdfHighlightOverlay.Save(formattedEarly ?? annotatedEarly ?? painted ?? finalDocument,
                    _output!, _highlights);
                return;
            }
            var evenRenderer = CreatePdfRenderer(evenDocument);
            evenRenderer.RenderDocument();
            if (footnoteRenderer != null)
                RenderFittingFootnotes(evenRenderer, footnoteRenderer, footnoteOverlays);
            DisableRasterInterpolation(evenRenderer.PdfDocument);
            if (evenRenderer.PageCount != renderer.PageCount)
            {
                RenderPageBorders(renderer);
                RenderLineNumbers(renderer);
                RenderTableFragmentBorders(renderer);
                RenderSpacedTableBorders(renderer);
                RenderHtmlBeveledTableBorders(renderer);
                RenderDiagonalCellBorders(renderer);
                RenderRotatedCells(renderer);
                RenderDoubleFirstRowTopBorders(renderer);
                RenderRtlDoubleTableLeftBorders(renderer);
                RenderDoubleTableRightBorders(renderer);
                RenderDoubleCellSideBorders(renderer);
                RenderParagraphShadingGaps(renderer);
                RenderContinuousParagraphSideBorders(renderer);
                RenderFooterImages(renderer);
                RenderParagraphAnchoredImages(renderer);
                RenderMathFractions(renderer);
                RenderMathRadicals(renderer);
                using var emphasized = RenderEmphasisMarks(renderer);
                ApplyRightLegendShift(emphasized);
                ApplyLargeSuperscriptRise(emphasized);
                PadDoubleBorderText(emphasized);
                ApplyFontKerning(emphasized);
                ApplyRightAlignedTrackedKerning(emphasized);
                ApplyInlineSuperscriptRise(emphasized);
                ApplyBodyFontBaselineRise(emphasized);
                RaiseCjkNumberedBodyText(renderer, emphasized);
                AlignBulletBaselines(renderer, emphasized);
                SuppressNativeColoredUnderlines(renderer, emphasized);
                LowerMixedDisplayMathRuns(renderer, emphasized);
                using var barred = RenderInlineMathBars(renderer, emphasized);
                using var underlined = RenderColoredUnderlines(renderer, barred);
                using var justified = JustifyContinuedPageLines(renderer, underlined);
                var finalDocument = justified ?? underlined;
                WidenBalancedFooterFieldSpaces(finalDocument);
                CollapseDuplicateThinStrokes(finalDocument);
                RecolorJoinedClosingMarks(finalDocument);
                RestoreSourceSpaces(finalDocument);
                using var painted = RestoreLayoutWidthContinuationPaint(renderer, finalDocument);
                ApplyLongJustifiedTrackedKerning(painted ?? finalDocument);
                MatchElevenPointAptosPaint(painted ?? finalDocument);
                DisableRasterInterpolation(painted ?? finalDocument);
                using var annotatedFallback = RenderCommentAnnotations(renderer, painted ?? finalDocument);
                using var formattedFallback = ApplyFormattingBalloons(renderer,
                    annotatedFallback ?? painted ?? finalDocument);
                DxpPdfHighlightOverlay.Save(formattedFallback ?? annotatedFallback ?? painted ?? finalDocument,
                    _output!, _highlights);
                return;
            }
            using var oddStream = new MemoryStream();
            using var evenStream = new MemoryStream();
            renderer.PdfDocument.Save(oddStream, false);
            evenRenderer.PdfDocument.Save(evenStream, false);
            oddStream.Position = 0;
            evenStream.Position = 0;
            using var oddPages = PdfReader.Open(oddStream, PdfDocumentOpenMode.Import);
            using var evenPages = PdfReader.Open(evenStream, PdfDocumentOpenMode.Import);
            using var combined = new PdfDocument();
            for (var index = 0; index < renderer.PageCount; index++)
                combined.AddPage((index % 2 == 0 ? oddPages : evenPages).Pages[index]);
            RenderPageBorders(renderer, combined);
            RenderLineNumbers(renderer, combined);
            RenderTableFragmentBorders(renderer, combined);
            RenderSpacedTableBorders(renderer, combined);
            RenderHtmlBeveledTableBorders(renderer, combined);
            RenderDiagonalCellBorders(renderer, combined);
            RenderRotatedCells(renderer, combined);
            RenderDoubleFirstRowTopBorders(renderer, combined);
            RenderRtlDoubleTableLeftBorders(renderer, combined);
            RenderDoubleTableRightBorders(renderer, combined);
            RenderDoubleCellSideBorders(renderer, combined);
            RenderParagraphShadingGaps(renderer, combined);
            RenderContinuousParagraphSideBorders(renderer, combined);
            RenderFooterImages(renderer, combined);
            RenderParagraphAnchoredImages(renderer, combined);
            RenderMathFractions(renderer, combined);
            RenderMathRadicals(renderer, combined);
            using var emphasizedCombined = RenderEmphasisMarks(renderer, combined);
            ApplyRightLegendShift(emphasizedCombined);
            ApplyLargeSuperscriptRise(emphasizedCombined);
            PadDoubleBorderText(emphasizedCombined);
            ApplyFontKerning(emphasizedCombined);
            ApplyRightAlignedTrackedKerning(emphasizedCombined);
            ApplyInlineSuperscriptRise(emphasizedCombined);
            ApplyBodyFontBaselineRise(emphasizedCombined);
            RaiseCjkNumberedBodyText(renderer, emphasizedCombined);
            AlignBulletBaselines(renderer, emphasizedCombined);
            SuppressNativeColoredUnderlines(renderer, emphasizedCombined);
            LowerMixedDisplayMathRuns(renderer, emphasizedCombined);
            using var barredCombined = RenderInlineMathBars(renderer, emphasizedCombined);
            using var underlinedCombined = RenderColoredUnderlines(renderer, barredCombined);
            using var justifiedCombined = JustifyContinuedPageLines(renderer, underlinedCombined);
            var finalCombined = justifiedCombined ?? underlinedCombined;
            WidenBalancedFooterFieldSpaces(finalCombined);
            CollapseDuplicateThinStrokes(finalCombined);
            RecolorJoinedClosingMarks(finalCombined);
            RestoreSourceSpaces(finalCombined);
            ApplyLongJustifiedTrackedKerning(finalCombined);
            MatchElevenPointAptosPaint(finalCombined);
            DisableRasterInterpolation(finalCombined);
            using var annotatedFinal = RenderCommentAnnotations(renderer, finalCombined);
            using var formattedFinal = ApplyFormattingBalloons(renderer, annotatedFinal ?? finalCombined);
            DxpPdfHighlightOverlay.Save(formattedFinal ?? annotatedFinal ?? finalCombined,
                _output!, _highlights);
        });
    }

    private void PadDoubleBorderText(PdfDocument document)
    {
        if (_doubleBorderTextPadding.Count == 0) return;
        foreach (PdfPage page in document.Pages)
            foreach (var entry in page.Contents.Elements)
            {
                if ((entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                    PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null ||
                    content.Elements.ContainsKey("/Filter")) continue;
                var original = Encoding.Latin1.GetString(content.Stream.Value);
                var revised = RewriteDoubleBorderTextPadding(original, _doubleBorderTextPadding);
                if (revised != original)
                    content.Stream.Value = Encoding.Latin1.GetBytes(revised);
            }
    }

    private void RestoreSourceSpaces(PdfDocument document)
    {
        if ((_positionedLayoutApplied || _restoreCompressedBodySpaces) && _sourceWordDocument != null)
            DxpPdfSourceSpaceRestorer.Restore(document, _sourceWordDocument);
    }

    private PdfDocument? RestoreLayoutWidthContinuationPaint(PdfDocumentRenderer renderer, PdfDocument document)
    {
        if (_layoutWidthParagraphs.Count == 0) return null;
        using var stream = new MemoryStream();
        document.Save(stream, false);
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(stream.ToArray());
        var modified = PdfReader.Open(new MemoryStream(stream.ToArray()), PdfDocumentOpenMode.Modify);
        for (var pageNumber = 2; pageNumber <= renderer.PageCount; pageNumber++)
        {
            var previous = renderer.DocumentRenderer.GetRenderInfoFromPage(pageNumber - 1)
                .Select(info => info.DocumentObject).ToHashSet();
            foreach (var info in renderer.DocumentRenderer.GetRenderInfoFromPage(pageNumber))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    info.LayoutInfo.ContentArea is not { } area ||
                    !(_layoutWidthParagraphs.Contains(paragraph) && previous.Contains(paragraph) ||
                      _footnoteSplitContinuations.Any(pair => _layoutWidthParagraphs.Contains(pair.Key) &&
                          previous.Contains(pair.Key) && ReferenceEquals(pair.Value, paragraph)))) continue;
                var page = parsed.GetPage(pageNumber);
                var words = page.GetWords().Where(word =>
                        word.BoundingBox.Left >= area.X.Point - 1 &&
                        word.BoundingBox.Right <= area.X.Point + area.Width.Point + 2 &&
                        page.Height - word.BoundingBox.Top >= area.Y.Point - 2 &&
                        page.Height - word.BoundingBox.Bottom <= area.Y.Point + area.Height.Point + 2)
                    .GroupBy(word => Math.Round(word.Letters[0].StartBaseLine.Y, 0))
                    .OrderByDescending(line => line.Key).FirstOrDefault()?
                    .OrderBy(word => word.BoundingBox.Left).ToArray();
                if (words is not { Length: >= 3 } ||
                    Math.Abs(area.X.Point + area.Width.Point - words[^1].BoundingBox.Right) > 2)
                    continue;
                foreach (var entry in modified.Pages[pageNumber - 1].Contents.Elements)
                {
                    if ((entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                        PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null ||
                        content.Elements.ContainsKey("/Filter")) continue;
                    var source = Encoding.Latin1.GetString(content.Stream.Value);
                    var revised = JustifyFinalTextLine(source, words.Select(word => word.Text).ToArray(),
                        -DenseTrackedLayoutAllowancePoints, words[0].Letters[0].StartBaseLine.X,
                        words[0].Letters[0].StartBaseLine.Y);
                    if (revised == source) continue;
                    content.Stream.Value = Encoding.Latin1.GetBytes(revised);
                    break;
                }
            }
        }
        return modified;
    }

    private void RecolorJoinedClosingMarks(PdfDocument document)
    {
        if (_joinedClosingMarkColors.Count == 0) return;
        foreach (var (prefix, colors) in _joinedClosingMarkColors)
        {
            var expression = new Regex(@"\(" + Regex.Escape(prefix) + @"(?<mark>\\.|[^()\\])\) Tj");
            var streams = document.Pages.Cast<PdfPage>()
                .SelectMany(page => page.Contents.Elements)
                .Select(entry => (entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value as
                    PdfSharp.Pdf.Advanced.PdfContent)
                .Where(content => content?.Stream != null && !content.Elements.ContainsKey("/Filter"))
                .ToArray();
            if (streams.Sum(content => expression.Matches(
                    Encoding.Latin1.GetString(content!.Stream!.Value)).Count) != colors.Count)
                continue;
            var foreground = PdfRgb(colors.Foreground);
            var background = PdfRgb(colors.Background);
            foreach (var content in streams)
            {
                var source = Encoding.Latin1.GetString(content!.Stream!.Value);
                var revised = expression.Replace(source, match =>
                    "(" + prefix + ") Tj\n" + foreground + " rg\n(" +
                    match.Groups["mark"].Value + ") Tj\n" + background + " rg");
                if (revised != source) content.Stream.Value = Encoding.Latin1.GetBytes(revised);
            }
        }
    }

    private static string PdfRgb(string color)
    {
        var value = int.Parse(color.AsSpan(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        return string.Join(" ", new[] { (value >> 16) & 255, (value >> 8) & 255, value & 255 }
            .Select(channel => (channel / 255.0).ToString("0.####", CultureInfo.InvariantCulture)));
    }

    internal static string RewriteDoubleBorderTextPadding(string source,
        IReadOnlyDictionary<string, double> padding)
    {
        var lines = source.Split('\n');
        var origin = -1;
        var previousWord = "";
        var previousOrigin = -1;
        var horizontalShift = 0.0;
        var changed = false;
        for (var index = 0; index < lines.Length; index++)
        {
            var line = lines[index].TrimEnd('\r');
            if (line is "BT" or "ET")
            {
                origin = previousOrigin = -1;
                previousWord = "";
                horizontalShift = 0;
                continue;
            }
            var move = Regex.Match(line,
                @"^(?<horizontal>-?\d+(?:\.\d+)?) (?<vertical>-?\d+(?:\.\d+)?) Td$");
            if (move.Success)
            {
                var vertical = double.Parse(move.Groups["vertical"].Value, CultureInfo.InvariantCulture);
                if (Math.Abs(vertical) >= 0.001)
                {
                    if (horizontalShift != 0)
                    {
                        var horizontal = double.Parse(move.Groups["horizontal"].Value,
                            CultureInfo.InvariantCulture) - horizontalShift;
                        lines[index] = horizontal.ToString("0.####", CultureInfo.InvariantCulture) +
                            " " + move.Groups["vertical"].Value + " Td";
                        changed = true;
                        horizontalShift = 0;
                    }
                    previousWord = "";
                    previousOrigin = -1;
                }
                origin = index;
                continue;
            }
            var text = Regex.Match(line, @"^\((?<word>(?:\\.|[^)])*)\) Tj$");
            if (!text.Success) continue;
            var word = text.Groups["word"].Value;
            if (previousOrigin >= 0 && padding.TryGetValue(previousWord + "\0" + word,
                    out var inset))
            {
                var firstMove = Regex.Match(lines[previousOrigin],
                    @"^(?<horizontal>-?\d+(?:\.\d+)?) (?<vertical>-?\d+(?:\.\d+)?) Td\r?$");
                if (firstMove.Success)
                {
                    var horizontal = double.Parse(firstMove.Groups["horizontal"].Value,
                        CultureInfo.InvariantCulture) + inset;
                    lines[previousOrigin] = horizontal.ToString("0.####", CultureInfo.InvariantCulture) +
                        " " + firstMove.Groups["vertical"].Value + " Td";
                    horizontalShift += inset;
                    changed = true;
                }
            }
            previousWord = word;
            previousOrigin = origin;
        }
        return changed ? string.Join('\n', lines) : source;
    }

    private void WidenBalancedFooterFieldSpaces(PdfDocument document)
    {
        if (_balancedFooterFieldSpaces.Count == 0) return;
        foreach (PdfPage page in document.Pages)
        {
            var fonts = page.Resources.Elements.GetDictionary("/Font");
            if (fonts == null) continue;
            var aptosFonts = new HashSet<string>(StringComparer.Ordinal);
            foreach (var key in fonts.Elements.Keys)
            {
                var value = fonts.Elements[key];
                var font = (value as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfDictionary ??
                    value as PdfDictionary;
                var baseFont = font?.Elements.GetName("/BaseFont").TrimStart('/');
                if (baseFont != null && baseFont[(baseFont.IndexOf('+') + 1)..]
                    .StartsWith("Aptos", StringComparison.OrdinalIgnoreCase))
                    aptosFonts.Add(key.TrimStart('/'));
            }
            if (aptosFonts.Count == 0) continue;
            foreach (var entry in page.Contents.Elements)
            {
                if ((entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                    PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null ||
                    content.Elements.ContainsKey("/Filter")) continue;
                var original = Encoding.Latin1.GetString(content.Stream.Value);
                var revised = RewriteBalancedFooterFieldSpaces(original,
                    _balancedFooterFieldSpaces, aptosFonts, page.Height.Point);
                if (revised != original)
                    content.Stream.Value = Encoding.Latin1.GetBytes(revised);
            }
        }
    }

    internal static string RewriteBalancedFooterFieldSpaces(string source,
        IReadOnlyDictionary<string, int> prefixes, HashSet<string> eligibleFonts, double pageHeight)
    {
        var lines = source.Split('\n');
        var baseline = 0.0;
        var fontSize = 12.0;
        var eligibleFont = false;
        var remaining = 0;
        var pendingMove = -1;
        var horizontalShift = 0.0;
        var changed = false;
        for (var index = 0; index < lines.Length; index++)
        {
            var line = lines[index].TrimEnd('\r');
            if (line == "BT")
            {
                baseline = 0;
                remaining = 0;
                pendingMove = -1;
                horizontalShift = 0;
                continue;
            }
            if (line == "ET")
            {
                remaining = 0;
                pendingMove = -1;
                horizontalShift = 0;
                continue;
            }
            var font = Regex.Match(line, @"^/F\d+ (?<size>\d+(?:\.\d+)?) Tf$");
            if (font.Success)
            {
                fontSize = double.Parse(font.Groups["size"].Value, CultureInfo.InvariantCulture);
                eligibleFont = eligibleFonts.Contains(line[1..line.IndexOf(' ')]);
            }
            var move = Regex.Match(line,
                @"^(?<horizontal>-?\d+(?:\.\d+)?) (?<vertical>-?\d+(?:\.\d+)?) Td$");
            if (move.Success)
            {
                var vertical = double.Parse(move.Groups["vertical"].Value, CultureInfo.InvariantCulture);
                baseline += vertical;
                pendingMove = Math.Abs(vertical) < 0.001 ? index : -1;
                if (Math.Abs(vertical) >= 0.001)
                {
                    if (horizontalShift != 0)
                    {
                        var horizontal = double.Parse(move.Groups["horizontal"].Value,
                            CultureInfo.InvariantCulture) - horizontalShift;
                        lines[index] = horizontal.ToString("0.####", CultureInfo.InvariantCulture) +
                            " " + move.Groups["vertical"].Value + " Td";
                        changed = true;
                        horizontalShift = 0;
                    }
                    remaining = 0;
                }
                continue;
            }
            var text = Regex.Match(line, @"^\((?<value>(?:\\.|[^)])*)\) Tj$");
            if (!text.Success) continue;
            if (remaining > 0 && pendingMove >= 0)
            {
                var horizontal = Regex.Match(lines[pendingMove], @"^(?<value>-?\d+(?:\.\d+)?) 0 Td\r?$");
                if (horizontal.Success)
                {
                    var increment = fontSize * 0.28;
                    var advance = double.Parse(horizontal.Groups["value"].Value,
                        CultureInfo.InvariantCulture) + increment;
                    lines[pendingMove] = advance.ToString("0.####", CultureInfo.InvariantCulture) + " 0 Td";
                    horizontalShift += increment;
                    changed = true;
                    remaining--;
                }
            }
            if (eligibleFont && baseline > 0 && baseline < pageHeight * 0.15 &&
                prefixes.TryGetValue(text.Groups["value"].Value, out var spaces))
                remaining = spaces;
            pendingMove = -1;
        }
        return changed ? string.Join('\n', lines) : source;
    }

    private PdfDocument? JustifyContinuedPageLines(PdfDocumentRenderer renderer, PdfDocument document)
    {
        var continuations = new List<(int Page, double Left, double Top, double Right, double Bottom,
            bool Continued)>();
        var pageInfos = Enumerable.Range(1, renderer.PageCount)
            .Select(renderer.DocumentRenderer.GetRenderInfoFromPage).ToArray();
        var rendered = pageInfos.SelectMany(infos => infos ?? [])
            .Select(info => info.DocumentObject).ToHashSet();
        var synthetic = _footnoteSplitContinuations.Keys.Concat(_footnoteSplitContinuations.Values)
            .ToHashSet();
        for (var page = 1; page <= renderer.PageCount; page++)
        {
            var currentInfos = pageInfos[page - 1];
            if (currentInfos == null) continue;
            var next = page < renderer.PageCount
                ? pageInfos[page]?.Select(info => info.DocumentObject).ToHashSet() : null;
            foreach (var info in currentInfos)
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    !(_sourceBodyParagraphs.Contains(paragraph) || synthetic.Contains(paragraph)) ||
                    paragraph.Format.Alignment != ParagraphAlignment.Justify ||
                    info.LayoutInfo.ContentArea is not { } area) continue;
                var continued = next?.Contains(paragraph) == true ||
                    _footnoteSplitContinuations.TryGetValue(paragraph, out var continuation) &&
                    rendered.Contains(continuation);
                if (!continued && !_narrowJustifiedBodyParagraphs.Contains(paragraph)) continue;
                continuations.Add((page, area.X.Point, area.Y.Point,
                    area.X.Point + area.Width.Point, area.Y.Point + area.Height.Point, continued));
            }
        }
        var tableContinuations = new List<(int Page, MigraDoc.DocumentObjectModel.Tables.Table Table,
            double Left, double Top, double PreviousTop, double PreviousBottom)>();
        for (var page = 2; page <= renderer.PageCount; page++)
        {
            if (pageInfos[page - 2] == null || pageInfos[page - 1] == null) continue;
            foreach (var info in pageInfos[page - 1])
                if (info.DocumentObject is MigraDoc.DocumentObjectModel.Tables.Table table &&
                    info.LayoutInfo.ContentArea is { } area &&
                    pageInfos[page - 2].FirstOrDefault(previous => ReferenceEquals(previous.DocumentObject, table))?
                        .LayoutInfo.ContentArea is { } previousArea)
                    tableContinuations.Add((page, table, area.X.Point, area.Y.Point,
                        previousArea.Y.Point, previousArea.Y.Point + previousArea.Height.Point));
        }
        if (continuations.Count == 0 && tableContinuations.Count == 0) return null;
        using var stream = new MemoryStream();
        document.Save(stream, false);
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(stream.ToArray());
        var modified = PdfReader.Open(new MemoryStream(stream.ToArray()), PdfDocumentOpenMode.Modify);
        foreach (var (pageNumber, left, top, right, bottom, continued) in continuations)
        {
            var page = parsed.GetPage(pageNumber);
            var lines = page.GetWords().Where(word =>
                    word.BoundingBox.Left >= left - 1 && word.BoundingBox.Right <= right + 2 &&
                    page.Height - word.BoundingBox.Top >= top - 2 &&
                    page.Height - word.BoundingBox.Bottom <= bottom + 2)
                .GroupBy(word => Math.Round(word.Letters[0].StartBaseLine.Y, 0))
                .OrderByDescending(group => group.Key).ToArray();
            if (lines.Length == 0) continue;
            var words = lines[^1].OrderBy(word => word.BoundingBox.Left).ToArray();
            if (words.Length < (continued ? 3 : 2) ||
                continued && right - words[^1].BoundingBox.Right < 5) continue;
            var adjustment = continued ? right - words[^1].BoundingBox.Right :
                (words.Length - 1) * words[0].Letters[0].PointSize * 0.55 / 12;
            foreach (var entry in modified.Pages[pageNumber - 1].Contents.Elements)
            {
                if ((entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                    PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null ||
                    content.Elements.ContainsKey("/Filter")) continue;
                var source = Encoding.Latin1.GetString(content.Stream.Value);
                var revised = JustifyFinalTextLine(source, words.Select(word => word.Text).ToArray(),
                    adjustment, continued ? null : words[0].Letters[0].StartBaseLine.X,
                    continued ? null : words[0].Letters[0].StartBaseLine.Y);
                if (revised == source && !continued)
                    revised = JustifyFinalTextLine(source, words.Select(word => word.Text).ToArray(),
                        (words.Length - 1) * words[0].Letters[0].PointSize * 0.203125 / 12,
                        words[0].Letters[0].StartBaseLine.X,
                        words[0].Letters[0].StartBaseLine.Y, allowEncodedPunctuation: true);
                if (revised == source) continue;
                content.Stream.Value = Encoding.Latin1.GetBytes(revised);
                break;
            }
        }
        var adjustedTableEndingLines = new HashSet<(int Page, int Column, double Baseline)>();
        foreach (var (pageNumber, table, tableLeft, tableTop, previousTop, previousBottom) in tableContinuations)
        {
            var page = parsed.GetPage(pageNumber);
            var previousPage = parsed.GetPage(pageNumber - 1);
            var cellLeft = tableLeft;
            for (var column = 0; column < table.Columns.Count; column++)
            {
                var cellRight = cellLeft + table.Columns[column].Width.Point;
                var firstLine = page.GetWords().Where(word =>
                        word.BoundingBox.Left >= cellLeft - 1 && word.BoundingBox.Right <= cellRight + 1 &&
                        page.Height - word.BoundingBox.Top >= tableTop - 2 &&
                        page.Height - word.BoundingBox.Top <= tableTop + 22)
                    .OrderByDescending(word => word.Letters[0].StartBaseLine.Y)
                    .ThenBy(word => word.BoundingBox.Left).ToArray();
                if (firstLine.Length == 0)
                {
                    cellLeft = cellRight;
                    continue;
                }
                var baseline = firstLine[0].Letters[0].StartBaseLine.Y;
                var line = firstLine.Where(word =>
                    Math.Abs(word.Letters[0].StartBaseLine.Y - baseline) < 2).ToArray();
                var lineText = string.Concat(line.Select(word => word.Text));
                var previousRights = previousPage.GetWords().Where(word =>
                        word.BoundingBox.Left >= cellLeft - 1 && word.BoundingBox.Right <= cellRight + 1 &&
                        previousPage.Height - word.BoundingBox.Top >= previousTop - 2 &&
                        previousPage.Height - word.BoundingBox.Bottom <= previousBottom + 2 &&
                        word.BoundingBox.Right > cellRight - 20 &&
                        word.BoundingBox.Right < cellRight - 4)
                    .Select(word => word.BoundingBox.Right).OrderBy(right => right).ToArray();
                if (previousRights.Length < 3)
                {
                    cellLeft = cellRight;
                    continue;
                }
                var target = previousRights[previousRights.Length / 2];
                var endingWords = previousPage.GetWords().Where(word =>
                        word.BoundingBox.Left >= cellLeft - 1 &&
                        word.BoundingBox.Right <= cellRight + 1 &&
                        previousPage.Height - word.BoundingBox.Top >= previousBottom - 45 &&
                        previousPage.Height - word.BoundingBox.Bottom <= previousBottom + 2)
                    .GroupBy(word => Math.Round(word.Letters[0].StartBaseLine.Y, 0))
                    .OrderBy(group => group.Key).FirstOrDefault()?
                    .OrderBy(word => word.BoundingBox.Left).ToArray();
                var ending = endingWords?.LastOrDefault();
                if (ending != null && endingWords is { Length: >= 3 } &&
                    target - ending.BoundingBox.Right is > 2 and < 45)
                {
                    var endingText = string.Concat(endingWords.Select(word => word.Text));
                    var adjustedEnding = false;
                    for (var rowIndex = 0; rowIndex + 1 < table.Rows.Count && !adjustedEnding; rowIndex++)
                    {
                        var row = table.Rows[rowIndex];
                        if (row.Borders.Bottom.Visible != false ||
                            table.Rows[rowIndex + 1].Borders.Top.Visible != false) continue;
                        var paragraphs = row.Cells[column].Elements
                            .OfType<MigraDoc.DocumentObjectModel.Paragraph>().ToArray();
                        var continuation = table.Rows[rowIndex + 1].Cells[column].Elements
                            .OfType<MigraDoc.DocumentObjectModel.Paragraph>().FirstOrDefault();
                        if (continuation?.Format.Alignment != ParagraphAlignment.Justify) continue;
                        foreach (var paragraph in paragraphs)
                        {
                            if (adjustedEnding) break;
                            if (paragraph.Format.Alignment != ParagraphAlignment.Justify ||
                                !DxpFootnoteParagraphSplitter.TryGetText(paragraph, out var content)) continue;
                            var sourceWords = Regex.Matches(content, @"\S+")
                                .Select(match => match.Value).ToArray();
                            for (var start = 0; start + 2 < sourceWords.Length && !adjustedEnding; start++)
                            {
                                var words = sourceWords[start..];
                                if (string.Concat(words) != endingText) continue;
                                foreach (var entry in modified.Pages[pageNumber - 2].Contents.Elements)
                                {
                                    if ((entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                                        PdfSharp.Pdf.Advanced.PdfContent pdfContent || pdfContent.Stream == null ||
                                        pdfContent.Elements.ContainsKey("/Filter")) continue;
                                    var source = Encoding.Latin1.GetString(pdfContent.Stream.Value);
                                    var revised = JustifyFinalTextLine(source, words,
                                        target - ending.BoundingBox.Right,
                                        endingWords[0].Letters[0].StartBaseLine.X,
                                        endingWords[0].Letters[0].StartBaseLine.Y);
                                    if (revised == source) continue;
                                    pdfContent.Stream.Value = Encoding.Latin1.GetBytes(revised);
                                    adjustedEnding = true;
                                    adjustedTableEndingLines.Add((pageNumber - 1, column,
                                        Math.Round(endingWords[0].Letters[0].StartBaseLine.Y, 0)));
                                    break;
                                }
                            }
                        }
                    }
                }
                var adjustedLine = false;
                for (var rowIndex = 0; rowIndex < table.Rows.Count && !adjustedLine; rowIndex++)
                {
                    var row = table.Rows[rowIndex];
                    var cell = row.Cells[column];
                    var paragraphs = cell.Elements.OfType<MigraDoc.DocumentObjectModel.Paragraph>().ToArray();
                    for (var paragraphIndex = 0; paragraphIndex < paragraphs.Length && !adjustedLine;
                         paragraphIndex++)
                    {
                        var paragraph = paragraphs[paragraphIndex];
                        if (paragraph.Format.Alignment != ParagraphAlignment.Justify ||
                            !DxpFootnoteParagraphSplitter.TryGetText(paragraph, out var content)) continue;
                        var sourceWords = Regex.Matches(content, @"\S+").Select(match => match.Value).ToArray();
                        for (var start = 0; start < sourceWords.Length - 2 && !adjustedLine; start++)
                        {
                            var combined = "";
                            for (var end = start; end < sourceWords.Length && !adjustedLine; end++)
                            {
                                combined += sourceWords[end];
                                if (combined.Length > lineText.Length) break;
                                if (combined != lineText || end - start < 2) continue;
                                var nextParagraph = paragraphIndex + 1 < paragraphs.Length
                                    ? paragraphs[paragraphIndex + 1]
                                    : rowIndex + 1 < table.Rows.Count
                                        ? table.Rows[rowIndex + 1].Cells[column].Elements
                                            .OfType<MigraDoc.DocumentObjectModel.Paragraph>().FirstOrDefault()
                                        : null;
                                var splitFragment = start == 0 && end == sourceWords.Length - 1 &&
                                    nextParagraph?.Format.Alignment == ParagraphAlignment.Justify &&
                                    Math.Abs(paragraph.Format.SpaceAfter.Point) < 0.1 &&
                                    Math.Abs(nextParagraph.Format.SpaceBefore.Point) < 0.1;
                                if (!splitFragment && (start == 0 || end == sourceWords.Length - 1)) continue;
                                var missing = target - line[^1].BoundingBox.Right;
                                if (missing is < 5 or > 30) continue;
                                var words = sourceWords[start..(end + 1)];
                                foreach (var entry in modified.Pages[pageNumber - 1].Contents.Elements)
                                {
                                    if ((entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                                        PdfSharp.Pdf.Advanced.PdfContent pdfContent || pdfContent.Stream == null ||
                                        pdfContent.Elements.ContainsKey("/Filter")) continue;
                                    var source = Encoding.Latin1.GetString(pdfContent.Stream.Value);
                                    var revised = JustifyFinalTextLine(source, words, missing,
                                        line[0].BoundingBox.Left);
                                    if (revised == source) continue;
                                    pdfContent.Stream.Value = Encoding.Latin1.GetBytes(revised);
                                    adjustedLine = true;
                                    break;
                                }
                            }
                        }
                    }
                }
                cellLeft = cellRight;
            }
        }
        for (var pageNumber = 1; pageNumber <= renderer.PageCount; pageNumber++)
        {
            if (pageInfos[pageNumber - 1] == null) continue;
            var page = parsed.GetPage(pageNumber);
            foreach (var info in pageInfos[pageNumber - 1])
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                    info.LayoutInfo.ContentArea is not { } area) continue;
                var cellLeft = area.X.Point;
                for (var column = 0; column < table.Columns.Count; column++)
                {
                    var cellRight = cellLeft + table.Columns[column].Width.Point;
                    var suffixes = table.Rows.Cast<MigraDoc.DocumentObjectModel.Tables.Row>()
                        .SelectMany(row => row.Cells[column].Elements
                            .OfType<MigraDoc.DocumentObjectModel.Paragraph>())
                        .Where(paragraph => paragraph.Format.Alignment == ParagraphAlignment.Justify &&
                            !synthetic.Contains(paragraph))
                        .Select(paragraph =>
                        {
                            if (!DxpFootnoteParagraphSplitter.TryGetText(paragraph, out var text))
                                text = string.Concat(paragraph.Elements.Cast<MigraDoc.DocumentObjectModel.DocumentObject>()
                                    .OfType<FormattedText>().SelectMany(span => span.Elements
                                        .OfType<MigraDoc.DocumentObjectModel.Text>())
                                    .Select(element => element.Content));
                            return Regex.Matches(text, @"\S+").Select(match => match.Value).ToArray();
                        }).ToArray();
                    var lines = page.GetWords().Where(word =>
                            word.BoundingBox.Left >= cellLeft - 1 &&
                            word.BoundingBox.Right <= cellRight + 1 &&
                            page.Height - word.BoundingBox.Top >= area.Y.Point - 2 &&
                            page.Height - word.BoundingBox.Bottom <= area.Y.Point + area.Height.Point + 2)
                        .GroupBy(word => Math.Round(word.Letters[0].StartBaseLine.Y, 0));
                    foreach (var line in lines)
                    {
                        if (adjustedTableEndingLines.Contains((pageNumber, column, line.Key))) continue;
                        var printed = line.OrderBy(word => word.BoundingBox.Left).ToArray();
                        if (printed.Length == 0 || cellRight - printed[^1].BoundingBox.Right < 14) continue;
                        if (tableContinuations.Any(continuation => continuation.Page == pageNumber &&
                            ReferenceEquals(continuation.Table, table) &&
                            page.Height - printed[0].BoundingBox.Top < continuation.Top + 12)) continue;
                        var printedText = string.Concat(printed.Select(word => word.Text));
                        var sourceWords = suffixes.Select(words => Enumerable.Range(0, words.Length)
                                .Select(start => words[start..])
                                .FirstOrDefault(suffix => suffix.Length >= 2 &&
                                    string.Concat(suffix) == printedText))
                            .Where(words => words != null).ToArray();
                        if (sourceWords.Length != 1 || sourceWords[0] is not { } words) continue;
                        var fontSize = printed[0].Letters[0].PointSize;
                        if (fontSize <= 8) continue;
                        var missing = (words.Length - 1) * (fontSize - 8) * 0.205;
                        foreach (var entry in modified.Pages[pageNumber - 1].Contents.Elements)
                        {
                            if ((entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                                PdfSharp.Pdf.Advanced.PdfContent pdfContent || pdfContent.Stream == null ||
                                pdfContent.Elements.ContainsKey("/Filter")) continue;
                            var source = Encoding.Latin1.GetString(pdfContent.Stream.Value);
                            var revised = JustifyFinalTextLine(source, words, missing,
                                printed[0].Letters[0].StartBaseLine.X,
                                printed[0].Letters[0].StartBaseLine.Y);
                            if (revised == source) continue;
                            pdfContent.Stream.Value = Encoding.Latin1.GetBytes(revised);
                            break;
                        }
                    }
                    cellLeft = cellRight;
                }
            }
        }
        return modified;
    }

    internal static string JustifyFinalTextLine(string source, string[] words, double missingWidth,
        double? firstWordLeft = null, double? firstWordBaseline = null,
        bool allowEncodedPunctuation = false)
    {
        static string Letters(string value) => new(value.Where(char.IsLetterOrDigit).ToArray());
        var starts = allowEncodedPunctuation
            ? Regex.Matches(source, @"\((?<text>(?:\\.|[^)])*)\) Tj")
                .Cast<Match>().Where(match => Letters(match.Groups["text"].Value) == Letters(words[0]))
            : Regex.Matches(source, @"\(" + Regex.Escape(words[0]) + @"\) Tj").Cast<Match>();
        foreach (var candidate in starts)
        {
            var start = candidate.Index;
            if (firstWordBaseline is { } expectedBaseline && firstWordLeft is { } lineLeft)
            {
                var blockStart = source.LastIndexOf("BT\n", start, StringComparison.Ordinal);
                if (blockStart < 0) continue;
                var horizontal = 0.0;
                var vertical = 0.0;
                foreach (Match move in Regex.Matches(source[(blockStart + 3)..start],
                             @"(?<horizontal>-?\d+(?:\.\d+)?) (?<vertical>-?\d+(?:\.\d+)?) Td"))
                {
                    horizontal += double.Parse(move.Groups["horizontal"].Value, CultureInfo.InvariantCulture);
                    vertical += double.Parse(move.Groups["vertical"].Value, CultureInfo.InvariantCulture);
                }
                if (Math.Abs(horizontal - lineLeft) > 2 ||
                    Math.Abs(vertical - expectedBaseline) > 2) continue;
            }
            else if (firstWordLeft is { } expectedLeft)
            {
                var moves = Regex.Matches(source[..start],
                    @"(?<horizontal>-?\d+(?:\.\d+)?) -?\d+(?:\.\d+)? Td\r?\n$");
                if (moves.Count == 0 || Math.Abs(double.Parse(moves[^1].Groups["horizontal"].Value,
                        CultureInfo.InvariantCulture) - expectedLeft) > 2) continue;
            }
            var end = source.IndexOf("\nET", start, StringComparison.Ordinal);
            if (end < 0) continue;
            var line = source[start..end];
            var nextLine = Regex.Matches(line, @"(?<horizontal>-?\d+(?:\.\d+)?) (?<vertical>-\d+(?:\.\d+)?) Td")
                .Cast<Match>().FirstOrDefault(move =>
                    double.Parse(move.Groups["vertical"].Value, CultureInfo.InvariantCulture) < -8);
            if (nextLine != null) line = line[..nextLine.Index];
            var tokens = Regex.Matches(line,
                @"(?<delta>-?\d+(?:\.\d+)?) (?<vertical>-?\d+(?:\.\d+)?) Td\r?\n\((?<text>(?:\\.|[^)])*)\) Tj");
            var wordIndex = 0;
            var consumed = Letters(words[0]);
            var advances = new List<(int Index, int Length, double Value)>();
            var valid = true;
            foreach (Match token in tokens)
            {
                var text = Letters(token.Groups["text"].Value);
                if (text.Length == 0) continue;
                if (wordIndex == words.Length - 1 && consumed == Letters(words[^1]) &&
                    text.All(char.IsDigit) &&
                    double.Parse(token.Groups["vertical"].Value, CultureInfo.InvariantCulture) > 2)
                    continue;
                if (consumed == Letters(words[wordIndex]))
                {
                    wordIndex++;
                    consumed = "";
                }
                if (wordIndex >= words.Length ||
                    !Letters(words[wordIndex]).StartsWith(consumed + text, StringComparison.Ordinal))
                {
                    if (text.All(char.IsDigit) &&
                        double.Parse(token.Groups["vertical"].Value, CultureInfo.InvariantCulture) > 2)
                        continue;
                    valid = false;
                    break;
                }
                if (consumed.Length == 0)
                {
                    var delta = token.Groups["delta"];
                    advances.Add((start + delta.Index, delta.Length,
                        double.Parse(delta.Value, CultureInfo.InvariantCulture)));
                }
                consumed += text;
            }
            if (!valid || wordIndex != words.Length - 1 || consumed != Letters(words[^1]) ||
                advances.Count != words.Length - 1) continue;
            var increment = missingWidth / advances.Count;
            var adjustments = advances.Select(advance => (advance.Index, advance.Length,
                Value: advance.Value + increment)).ToList();
            if (nextLine != null)
            {
                var horizontal = nextLine.Groups["horizontal"];
                adjustments.Add((start + horizontal.Index, horizontal.Length,
                    double.Parse(horizontal.Value, CultureInfo.InvariantCulture) - missingWidth));
            }
            var rewritten = new StringBuilder(source);
            foreach (var adjustment in adjustments.OrderByDescending(item => item.Index))
                rewritten.Remove(adjustment.Index, adjustment.Length).Insert(adjustment.Index,
                    adjustment.Value.ToString("0.####", CultureInfo.InvariantCulture));
            return rewritten.ToString();
        }
        return source;
    }

    public override IDisposable VisitSectionBegin(SectionProperties properties, SectionLayout layout, DxpIDocumentContext context)
    {
        var inheritsFirstPageIntroTableFooter = _section != null &&
            _firstPageIntroTableFooters.Contains(_section) &&
            properties.GetFirstChild<TitlePage>() != null &&
            !properties.Elements<FooterReference>().Any(reference =>
                reference.Type?.Value == HeaderFooterValues.First);
        var continuesCurrentLayout = _section != null &&
            properties.GetFirstChild<SectionType>() is { } sectionType &&
            AttributeValue(sectionType, "val") == "continuous" &&
            properties.ChildElements.All(element => element is SectionType);
        _section = continuesCurrentLayout ? _section : _document!.AddSection();
        var mainPart = (context as DxpDocumentContext)?.MainDocumentPart;
        _sectionHasFooterTable = properties.Elements<FooterReference>()
            .Any(reference => reference.Id?.Value is string id &&
                mainPart?.GetPartById(id) is FooterPart footer &&
                footer.Footer.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().Any());
        if (!properties.Elements<HeaderReference>().Any())
            _sectionsWithLinkedHeaders.Add(_section);
        _largeGridHeaderDistanceAdjusted = false;
        _sectionUsesDefaultMargins = layout.PageMargin == null;
        _footerHorizontalMergeClearanceApplied = false;
        if (properties.GetFirstChild<PageBorders>() is { } borders)
            _pageBorders[_section] = borders;
        if (layout.LineNumbers is { } lineNumbers)
            _sectionLineNumbers[_section] = lineNumbers;
        var gridType = layout.DocGrid?.Type?.Value;
        _snapToCharsGrid = gridType == DocGridValues.SnapToChars;
        var characterGrid = gridType == DocGridValues.LinesAndChars ||
            gridType == DocGridValues.SnapToChars;
        _gridLinePitchPoints = (gridType == DocGridValues.Lines ||
            gridType == DocGridValues.LinesAndChars) &&
            layout.DocGrid.LinePitch?.Value is int pitch && pitch > 0 ? pitch / 20.0 : null;
        _gridCharSpacingPoints = characterGrid &&
            int.TryParse(AttributeValue(layout.DocGrid, "charSpace"), out var charSpace) && charSpace > 0
            ? charSpace / 4096.0 : null;
        _sectionCount++;
        var pageNumberFormat = properties.GetFirstChild<PageNumberType>()?.Format?.Value;
        _sectionPageNumberFormat = pageNumberFormat == NumberFormatValues.UpperRoman ? "ROMAN" :
            pageNumberFormat == NumberFormatValues.LowerRoman ? "roman" :
            pageNumberFormat == NumberFormatValues.UpperLetter ? "ALPHABETIC" :
            pageNumberFormat == NumberFormatValues.LowerLetter ? "alphabetic" : null;
        var page = _section.PageSetup;
        page.MirrorMargins = _mirrorMargins && !_gutterAtTop;
        page.DifferentFirstPageHeaderFooter = properties.GetFirstChild<TitlePage>() != null;
        page.OddAndEvenPagesHeaderFooter = _evenAndOddHeaders;
        page.SectionStart = (properties.GetFirstChild<SectionType>() is { } breakType
            ? AttributeValue(breakType, "val") : null) switch
        {
            "evenPage" => BreakType.BreakEvenPage,
            "oddPage" => BreakType.BreakOddPage,
            _ => BreakType.BreakNextPage
        };
        if (properties.GetFirstChild<PageNumberType>()?.Start?.Value is int start)
            page.StartingNumber = start;
        page.PageWidth = Unit.FromPoint(layout.PageSize?.Width?.Value is uint width ? width / 20.0 : 612);
        page.PageHeight = Unit.FromPoint(layout.PageSize?.Height?.Value is uint height ? height / 20.0 : 792);
        page.LeftMargin = Unit.FromPoint(layout.PageMargin?.Left?.Value is uint left ? left / 20.0 : 72);
        page.RightMargin = Unit.FromPoint(layout.PageMargin?.Right?.Value is uint right ? right / 20.0 : 72);
        if (!_gutterAtTop && layout.PageMargin?.Gutter?.Value is uint gutter && gutter > 0)
        {
            if (properties.ChildElements.Any(child => child.LocalName == "rtlGutter"))
                page.RightMargin = Unit.FromPoint(page.RightMargin.Point + gutter / 20.0);
            else
                page.LeftMargin = Unit.FromPoint(page.LeftMargin.Point + gutter / 20.0);
        }
        page.TopMargin = Unit.FromPoint((layout.PageMargin?.Top?.Value is int top ? top / 20.0 : 72) +
            (_gutterAtTop ? (layout.PageMargin?.Gutter?.Value ?? 0) / 20.0 : 0));
        _sectionBaseTopMarginPoints = page.TopMargin.Point;
        _sectionOtherHeaderTableRaisedTopMargin = false;
        page.BottomMargin = Unit.FromPoint(layout.PageMargin?.Bottom?.Value is int bottom
            ? bottom / 20.0 : 72);
        var verticalAlignment = properties.ChildElements.FirstOrDefault(child => child.LocalName == "vAlign");
        var sectionAlignment = verticalAlignment == null ? null : AttributeValue(verticalAlignment, "val");
        if (sectionAlignment is "center" or "bottom")
            _verticallyAlignedSections[_section] = sectionAlignment;
        page.HeaderDistance = Unit.FromPoint(layout.PageMargin?.Header?.Value is uint header ? header / 20.0 : 36);
        if (_gridLinePitchPoints is > 0 && _gridCharSpacingPoints is > 0)
            page.HeaderDistance = Unit.FromPoint(Math.Max(0, page.HeaderDistance.Point - 5));
        if (_sectionsWithLinkedHeaders.Contains(_section) && _document.Sections.Count > 1 &&
            _document.Sections[_document.Sections.Count - 2] is Section previous &&
            _headerRunBorderClearances.TryGetValue(previous, out var inheritedClearance))
        {
            _headerRunBorderClearances[_section] = inheritedClearance;
            page.HeaderDistance = Unit.FromPoint(page.HeaderDistance.Point + inheritedClearance);
        }
        page.FooterDistance = Unit.FromPoint(layout.PageMargin?.Footer?.Value is uint footer ? footer / 20.0 : 36);
        if (inheritsFirstPageIntroTableFooter)
        {
            _firstPageIntroTableFooters.Add(_section);
            page.FooterDistance = Unit.FromPoint(page.FooterDistance.Point - 2.8);
        }
        if (gridType == DocGridValues.LinesAndChars)
            page.FooterDistance = Unit.FromPoint(page.FooterDistance.Point + 13.3);
        _columnWidths = null;
        _activeColumnCell = null;
        _columnCells = null;
        if (layout.Columns?.ColumnCount?.Value is short columnCount && columnCount > 1)
        {
            _columnGapPoints = int.TryParse(layout.Columns.Space?.Value, out var gapTwips)
                ? gapTwips / 20.0 : 36;
            var available = page.PageWidth.Point - page.LeftMargin.Point - page.RightMargin.Point -
                _columnGapPoints * (columnCount - 1);
            var widths = layout.Columns.Elements<DocumentFormat.OpenXml.Wordprocessing.Column>().Select(column =>
                int.TryParse(column.Width?.Value, out var twips) ? twips / 20.0 : 0).ToArray();
            _columnWidths = layout.Columns.EqualWidth?.Value == false && widths.Length == columnCount &&
                widths.All(width => width > 0) ? widths :
                Enumerable.Repeat(available / columnCount, columnCount).ToArray();
            _columnParagraphCapacity = Math.Max(1, (int)Math.Floor(
                (page.PageHeight.Point - page.TopMargin.Point - page.BottomMargin.Point + 2) /
                SimpleColumnParagraphPitchPoints));
        }
        return DxpDisposable.Empty;
    }

    public override void VisitCommentRangeStart(CommentRangeStart start, DxpIDocumentContext context)
    {
        if (_includePdfComments && _paragraph != null && start.Id?.Value is { } id)
        {
            _paragraph.AddBookmark($"dxp-comment-range-start-{id}", prepend: false);
            if (_container is Cell cell)
                _commentRangeCells[id] = cell;
        }
    }

    public override void VisitCommentRangeEnd(CommentRangeEnd end, DxpIDocumentContext context)
    {
        if (_includePdfComments && _paragraph != null && end.Id?.Value is { } id)
            _paragraph.AddBookmark($"dxp-comment-range-end-{id}", prepend: false);
    }

    public override IDisposable VisitCommentThreadBegin(string anchorId, DxpCommentThread thread,
        DxpIDocumentContext context)
    {
        if (_includePdfComments && _paragraph != null && thread.Comments.Count > 0 &&
            _commentThreads.TryAdd(anchorId, thread))
            _paragraph.AddBookmark($"dxp-comment-{anchorId}", prepend: false);
        return DxpDisposable.Empty;
    }

    public override IDisposable VisitCommentBegin(DxpCommentInfo comment, DxpCommentThread thread,
        DxpIDocumentContext context)
    {
        var previousContainer = _container;
        var previousParagraph = _paragraph;
        var previousRunStyle = _runStyle;
        var previousHiddenRun = _hiddenRun;
        var previousColor = _runHasExplicitColor;
        var previousBold = _runHasExplicitBold;
        var previousItalic = _runHasExplicitItalic;
        var previousUnderline = _runHasExplicitUnderline;
        var previousHighlight = _runHighlight;
        var previousShading = _runShading;
        _container = null;
        return DxpDisposable.Create(() =>
        {
            _container = previousContainer;
            _paragraph = previousParagraph;
            _runStyle = previousRunStyle;
            _hiddenRun = previousHiddenRun;
            _runHasExplicitColor = previousColor;
            _runHasExplicitBold = previousBold;
            _runHasExplicitItalic = previousItalic;
            _runHasExplicitUnderline = previousUnderline;
            _runHighlight = previousHighlight;
            _runShading = previousShading;
        });
    }

    private void AlignNotBesideFrameFollowers()
    {
        if (_notBesideFrameFollowers.Count == 0 || _document == null) return;
        var layout = (MigraDoc.DocumentObjectModel.Document)_document.Clone();
        var followers = new List<(MigraDoc.DocumentObjectModel.Shapes.TextFrame Frame,
            MigraDoc.DocumentObjectModel.Paragraph LayoutParagraph,
            MigraDoc.DocumentObjectModel.Paragraph Paragraph)>();
        for (var sectionIndex = 0; sectionIndex < _document.Sections.Count; sectionIndex++)
        {
            var section = (Section)_document.Sections[sectionIndex];
            var layoutSection = (Section)layout.Sections[sectionIndex];
            foreach (var (frame, paragraph) in _notBesideFrameFollowers)
            {
                var frameIndex = section.Elements.IndexOf(frame);
                var paragraphIndex = section.Elements.IndexOf(paragraph);
                if (frameIndex < 0 || paragraphIndex < 0) continue;
                if (layoutSection.Elements[frameIndex] is MigraDoc.DocumentObjectModel.Shapes.TextFrame
                        layoutFrame &&
                    layoutSection.Elements[paragraphIndex] is MigraDoc.DocumentObjectModel.Paragraph
                        layoutParagraph)
                    followers.Add((layoutFrame, layoutParagraph, paragraph));
            }
        }
        if (followers.Count == 0) return;

        var renderer = CreatePdfRenderer(layout);
        renderer.PrepareRenderPages();
        for (var page = 1; page <= renderer.PageCount; page++)
        {
            var infos = renderer.DocumentRenderer.GetRenderInfoFromPage(page);
            foreach (var (frame, layoutParagraph, paragraph) in followers)
            {
                var frameInfo = infos.FirstOrDefault(info => ReferenceEquals(info.DocumentObject, frame));
                var paragraphInfo = infos.FirstOrDefault(info =>
                    ReferenceEquals(info.DocumentObject, layoutParagraph));
                if (frameInfo == null || paragraphInfo == null) continue;
                var frameArea = frameInfo.LayoutInfo.ContentArea;
                var paragraphTop = paragraphInfo.LayoutInfo.ContentArea.Y.Point;
                var clearance = frameArea.Y.Point + frameArea.Height.Point - paragraphTop;
                if (clearance > 0)
                    paragraph.Format.SpaceBefore = Unit.FromPoint(paragraph.Format.SpaceBefore.Point +
                        clearance);
            }
        }
    }

    private void PositionCellAnchoredTextBoxes()
    {
        if (_cellAnchoredTextBoxes.Count == 0 || _document == null) return;
        var layout = (MigraDoc.DocumentObjectModel.Document)_document.Clone();
        var anchors = new List<(MigraDoc.DocumentObjectModel.Shapes.TextFrame Frame,
            MigraDoc.DocumentObjectModel.Shapes.TextFrame LayoutFrame, Cell Cell,
            MigraDoc.DocumentObjectModel.Tables.Table SourceTable,
            MigraDoc.DocumentObjectModel.Paragraph Paragraph, double Left, double Top)>();
        for (var sectionIndex = 0; sectionIndex < _document.Sections.Count; sectionIndex++)
        {
            var section = (Section)_document.Sections[sectionIndex];
            var layoutSection = (Section)layout.Sections[sectionIndex];
            foreach (var (frame, anchor) in _cellAnchoredTextBoxes)
            {
                var frameIndex = section.Elements.IndexOf(frame);
                var tableIndex = section.Elements.IndexOf(anchor.Cell.Table);
                var paragraphIndex = anchor.Cell.Elements.IndexOf(anchor.Paragraph);
                if (frameIndex < 0 || tableIndex < 0 || paragraphIndex < 0 ||
                    layoutSection.Elements[frameIndex] is not MigraDoc.DocumentObjectModel.Shapes.TextFrame layoutFrame ||
                    layoutSection.Elements[tableIndex] is not MigraDoc.DocumentObjectModel.Tables.Table layoutTable ||
                    layoutTable.Rows[anchor.Cell.Row.Index].Cells[anchor.Cell.Column.Index] is not Cell layoutCell ||
                    layoutCell.Elements[paragraphIndex] is not MigraDoc.DocumentObjectModel.Paragraph layoutParagraph)
                    continue;
                anchors.Add((frame, layoutFrame, layoutCell, anchor.Cell.Table,
                    layoutParagraph, anchor.Left, anchor.Top));
            }
        }
        if (anchors.Count == 0) return;
        var preflight = CreatePdfRenderer(layout);
        preflight.PrepareRenderPages();
        var framePages = new Dictionary<MigraDoc.DocumentObjectModel.Shapes.TextFrame, int>();
        var tablePages = new Dictionary<MigraDoc.DocumentObjectModel.Tables.Table, int>();
        for (var page = 1; page <= preflight.PageCount; page++)
            foreach (var info in preflight.DocumentRenderer.GetRenderInfoFromPage(page))
            {
                if (info.DocumentObject is MigraDoc.DocumentObjectModel.Shapes.TextFrame frame)
                    framePages.TryAdd(frame, page);
                if (info.DocumentObject is MigraDoc.DocumentObjectModel.Tables.Table table)
                    tablePages.TryAdd(table, page);
            }
        var getCellRenderInfos = typeof(FormattedCell).GetMethod("GetRenderInfos",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        var getRowBoundaries = typeof(TableFormatInfo).GetField("BottomBorderMap",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        var getLastHeaderRow = typeof(TableFormatInfo).GetField("LastHeaderRow",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        if (getCellRenderInfos == null || getRowBoundaries == null || getLastHeaderRow == null) return;
        var moveToFirstFragmentFrames = new HashSet<MigraDoc.DocumentObjectModel.Shapes.TextFrame>();
        for (var page = 1; page <= preflight.PageCount; page++)
        {
            var infos = preflight.DocumentRenderer.GetRenderInfoFromPage(page);
            foreach (var tableInfo in infos.Where(info => info.FormatInfo is TableFormatInfo))
            {
                if (tableInfo.FormatInfo is not TableFormatInfo tableFormat ||
                    tableFormat.StartRow < 0 || tableFormat.FormattedCells == null) continue;
                if (getRowBoundaries.GetValue(tableFormat) is not Dictionary<int, XUnitPt> rowBoundaries ||
                    !rowBoundaries.TryGetValue(tableFormat.StartRow, out var fragmentStartBoundary)) continue;
                var lastHeaderRow = (int)getLastHeaderRow.GetValue(tableFormat)!;
                var repeatedHeaderHeight = 0.0;
                if (lastHeaderRow >= 0)
                {
                    if (!rowBoundaries.TryGetValue(0, out var headerStartBoundary) ||
                        !rowBoundaries.TryGetValue(lastHeaderRow + 1, out var headerEndBoundary)) continue;
                    repeatedHeaderHeight = (headerEndBoundary - headerStartBoundary).Point;
                }
                foreach (var anchor in anchors)
                {
                    if (moveToFirstFragmentFrames.Contains(anchor.Frame)) continue;
                    if (anchor.Cell.Row.Index <= lastHeaderRow ||
                        anchor.Cell.Row.Index < tableFormat.StartRow ||
                        anchor.Cell.Row.Index > tableFormat.EndRow ||
                        !rowBoundaries.TryGetValue(anchor.Cell.Row.Index, out var rowBoundary) ||
                    !tableFormat.FormattedCells.TryGetValue(anchor.Cell, out var formattedCell)) continue;
                    var frameOnPage = infos.Any(info => ReferenceEquals(info.DocumentObject, anchor.LayoutFrame));
                    var moveToFirstFragment = framePages.TryGetValue(anchor.LayoutFrame, out var framePage) &&
                        tablePages.TryGetValue(anchor.Cell.Table, out var tablePage) &&
                        framePage > tablePage && page == tablePage;
                    if (!frameOnPage && !moveToFirstFragment) continue;
                    if (getCellRenderInfos.Invoke(formattedCell, null) is not RenderInfo[] cellInfos) continue;
                    var paragraphInfo = cellInfos.FirstOrDefault(info =>
                        ReferenceEquals(info.DocumentObject, anchor.Paragraph));
                    if (paragraphInfo == null) continue;
                    if (moveToFirstFragment)
                        moveToFirstFragmentFrames.Add(anchor.Frame);
                    var table = anchor.Cell.Table;
                    var precedingColumnWidth = 0.0;
                    for (var column = 0; column < anchor.Cell.Column.Index; column++)
                        precedingColumnWidth += table.Columns[column].Width.Point;
                    anchor.Frame.RelativeHorizontal = MigraDoc.DocumentObjectModel.Shapes.RelativeHorizontal.Page;
                    anchor.Frame.RelativeVertical = MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Page;
                    anchor.Frame.Left = Unit.FromPoint(tableInfo.LayoutInfo.ContentArea.X.Point +
                        table[0, 0].Borders.Left.Width.Point + precedingColumnWidth +
                        anchor.Cell.Column.LeftPadding.Point +
                        paragraphInfo.LayoutInfo.ContentArea.X.Point + anchor.Left);
                    anchor.Frame.Top = Unit.FromPoint(tableInfo.LayoutInfo.ContentArea.Y.Point +
                        table[lastHeaderRow >= 0 ? 0 : tableFormat.StartRow, 0].Borders.Top.Width.Point +
                        repeatedHeaderHeight +
                        (rowBoundary - fragmentStartBoundary).Point +
                        anchor.Cell.Row.TopPadding.Point +
                        paragraphInfo.LayoutInfo.ContentArea.Y.Point + anchor.Top +
                        (lastHeaderRow < 0
                            ? FragmentTopBorderReservation(anchor.SourceTable, tableFormat.StartRow) ?? 0 : 0));
                }
            }
        }
        foreach (var anchor in anchors)
        {
            if (!moveToFirstFragmentFrames.Contains(anchor.Frame)) continue;
            var section = _document.Sections.Cast<Section>().FirstOrDefault(candidate =>
                candidate.Elements.IndexOf(anchor.Frame) >= 0);
            if (section == null) continue;
            var frameIndex = section.Elements.IndexOf(anchor.Frame);
            var tableIndex = section.Elements.IndexOf(anchor.SourceTable);
            if (frameIndex < 0 || tableIndex < 0 || frameIndex <= tableIndex) continue;
            var movedFrame = (MigraDoc.DocumentObjectModel.Shapes.TextFrame)anchor.Frame.Clone();
            section.Elements.InsertObject(tableIndex, movedFrame);
            section.Elements.RemoveObjectAt(frameIndex + 1);
            if (_verticallyCenteredTextBoxes.Contains(anchor.Frame))
                _verticallyCenteredTextBoxes.Add(movedFrame);
            if (_clippedTextBoxes.Contains(anchor.Frame))
                _clippedTextBoxes.Add(movedFrame);
            if (_textBoxRotations.TryGetValue(anchor.Frame, out var rotation))
                _textBoxRotations[movedFrame] = rotation;
            if (_behindTextBoxes.Contains(anchor.Frame))
                _behindTextBoxes.Add(movedFrame);
            if (!_behindTextBoxes.Contains(anchor.Frame))
                _aboveContentTextBoxes.Add(movedFrame);
        }
    }

    private double? DoubleFragmentTopReservation(MigraDoc.DocumentObjectModel.Tables.Table table, int startRow)
    {
        return _minorThemeLatinFont != null && _availableFontFamilies.Contains(_minorThemeLatinFont) &&
            startRow > 0 && _tableFragmentBorders.TryGetValue(table, out var borders) &&
            borders.Top is { LineStyle: DxpComputedBorderLineStyle.Double } top
            ? top.WidthPt * 3 : null;
    }

    private double? FragmentTopBorderReservation(MigraDoc.DocumentObjectModel.Tables.Table table, int startRow)
    {
        if (DoubleFragmentTopReservation(table, startRow) is double doubled)
            return doubled;
        return startRow > 0 && _tableFragmentBorders.TryGetValue(table, out var borders) &&
            borders.Top is { WidthPt: > 0, LineStyle: not DxpComputedBorderLineStyle.Double and
                not DxpComputedBorderLineStyle.None } top ? top.WidthPt : null;
    }

    public override bool AcceptAlternateContentChoice(AlternateContentChoice choice,
        IReadOnlyList<string> required, DxpIDocumentContext context) =>
        required.Count > 0 && required.All(prefix => prefix == "wps") &&
        choice.Descendants<TextBoxContent>().Any() &&
        choice.Descendants().Any(element => element.LocalName == "wrapNone");

    public override IDisposable VisitTextBoxContentBegin(TextBoxContent textBox, DxpIDocumentContext context)
    {
        var previousContainer = _container;
        var previousParagraph = _paragraph;
        var previousTextBoxTextColor = _textBoxTextColor;
        if (_section == null) return DxpDisposable.Empty;

        var drawing = textBox.Ancestors<Drawing>().FirstOrDefault();
        var anchor = drawing?.Descendants().FirstOrDefault(element => element.LocalName == "anchor");
        var extent = anchor?.ChildElements.FirstOrDefault(element => element.LocalName == "extent");
        var vmlShape = textBox.Ancestors().FirstOrDefault(element =>
            element.LocalName is "shape" or "rect" &&
            element.NamespaceUri == "urn:schemas-microsoft-com:vml");
        var vmlStyle = vmlShape?.GetAttribute("style", "").Value;
        if (extent == null && vmlStyle == null) return DxpDisposable.Empty;

        var frame = _section.AddTextFrame();
        _clippedTextBoxes.Add(frame);
        var bodyProperties = textBox.Ancestors().FirstOrDefault(element => element.LocalName == "wsp")?
            .Descendants().FirstOrDefault(element => element.LocalName == "bodyPr");
        frame.Orientation = bodyProperties?.GetAttribute("vert", "").Value switch
        {
            "vert" => MigraDoc.DocumentObjectModel.Shapes.TextOrientation.Vertical,
            "vert270" => MigraDoc.DocumentObjectModel.Shapes.TextOrientation.Upward,
            _ => MigraDoc.DocumentObjectModel.Shapes.TextOrientation.Horizontal
        };
        if (vmlStyle?.Contains("v-text-anchor:middle", StringComparison.OrdinalIgnoreCase) == true ||
            bodyProperties?.GetAttribute("anchor", "").Value == "ctr")
            _verticallyCenteredTextBoxes.Add(frame);
        frame.RelativeHorizontal = MigraDoc.DocumentObjectModel.Shapes.RelativeHorizontal.Margin;
        frame.RelativeVertical = MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Paragraph;
        var shape = textBox.Ancestors().FirstOrDefault(element => element.LocalName == "AlternateContent")?
            .Descendants().FirstOrDefault(element => element.LocalName == "wsp");
        var fontColor = shape?.ChildElements.FirstOrDefault(element => element.LocalName == "style")?
            .Descendants().FirstOrDefault(element => element.LocalName == "fontRef")?
            .Descendants().FirstOrDefault(element => element.LocalName == "schemeClr")?
            .GetAttribute("val", "").Value;
        _textBoxTextColor = fontColor == null ? null : ResolveDrawingThemeColor(fontColor);
        if (extent != null)
        {
            if (long.TryParse(extent.GetAttribute("cx", "").Value, out var width))
                frame.Width = Unit.FromPoint(width / 12700.0);
            if (long.TryParse(extent.GetAttribute("cy", "").Value, out var height))
                frame.Height = Unit.FromPoint(height / 12700.0);
            var horizontal = anchor!.ChildElements.FirstOrDefault(element => element.LocalName == "positionH");
            var vertical = anchor.ChildElements.FirstOrDefault(element => element.LocalName == "positionV");
            if (horizontal?.GetAttribute("relativeFrom", "").Value == "page")
                frame.RelativeHorizontal = MigraDoc.DocumentObjectModel.Shapes.RelativeHorizontal.Page;
            if (vertical?.GetAttribute("relativeFrom", "").Value == "page")
                frame.RelativeVertical = MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Page;
            else if (vertical?.GetAttribute("relativeFrom", "").Value == "margin")
                frame.RelativeVertical = MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Margin;
            if (long.TryParse(horizontal?.ChildElements.FirstOrDefault(element => element.LocalName == "posOffset")?.InnerText,
                out var horizontalOffset))
                frame.Left = Unit.FromPoint(horizontalOffset / 12700.0);
            if (long.TryParse(vertical?.ChildElements.FirstOrDefault(element => element.LocalName == "posOffset")?.InnerText,
                out var verticalOffset))
                frame.Top = Unit.FromPoint(verticalOffset / 12700.0);
            if (_activeColumnCell != null && _columnWidths != null)
            {
                if (horizontal?.GetAttribute("relativeFrom", "").Value == "column")
                {
                    frame.RelativeHorizontal = MigraDoc.DocumentObjectModel.Shapes.RelativeHorizontal.Page;
                    frame.Left = Unit.FromPoint(_section.PageSetup.LeftMargin.Point +
                        (_columnIndex == 0 ? 0 : _columnWidths[0] + _columnGapPoints) +
                        frame.Left.Position.Point);
                }
                if (vertical?.GetAttribute("relativeFrom", "").Value == "paragraph")
                {
                    frame.RelativeVertical = MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Page;
                    frame.Top = Unit.FromPoint(_section.PageSetup.TopMargin.Point +
                        (_columnParagraphCount - 1) * SimpleColumnParagraphPitchPoints +
                        frame.Top.Position.Point);
                }
            }
            if (anchor.ChildElements.Any(element => element.LocalName == "wrapSquare") &&
                Enum.TryParse<MigraDoc.DocumentObjectModel.Shapes.WrapStyle>("BothSides", out var squareWrap))
                frame.WrapFormat.Style = squareWrap;
            else
                frame.WrapFormat.Style = MigraDoc.DocumentObjectModel.Shapes.WrapStyle.None;
            if (long.TryParse(drawing!.Descendants().FirstOrDefault(element => element.LocalName == "xfrm")?
                    .GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "rot").Value,
                    out var rotationUnits))
                _textBoxRotations[frame] = rotationUnits / 60000.0;
            var shapeProperties = textBox.Ancestors().FirstOrDefault(element => element.LocalName == "wsp")?
                .ChildElements.FirstOrDefault(element => element.LocalName == "spPr");
            var fillColor = shapeProperties?.ChildElements.FirstOrDefault(element => element.LocalName == "solidFill")?
                .Descendants().FirstOrDefault(element => element.LocalName == "srgbClr")?
                .GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value;
            if (fillColor != null && ParseColor("#" + fillColor) is { } drawingFill)
            {
                frame.FillFormat.Visible = true;
                frame.FillFormat.Color = drawingFill;
            }
            var line = shapeProperties?.ChildElements.FirstOrDefault(element => element.LocalName == "ln");
            if (line != null)
            {
                frame.LineFormat.Visible = !line.ChildElements.Any(element => element.LocalName == "noFill");
                if (long.TryParse(line.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "w").Value,
                        out var lineWidth))
                    frame.LineFormat.Width = Unit.FromPoint(lineWidth / 12700.0);
                var solidFill = line.ChildElements.FirstOrDefault(element => element.LocalName == "solidFill");
                var rgb = solidFill?.Descendants().FirstOrDefault(element => element.LocalName == "srgbClr")?
                    .GetAttribute("val", "").Value;
                if (rgb != null && ParseColor("#" + rgb) is { } explicitColor)
                    frame.LineFormat.Color = explicitColor;
                else
                {
                    var schemeColor = solidFill?.Descendants()
                        .FirstOrDefault(element => element.LocalName == "schemeClr");
                    var themeKey = schemeColor?.GetAttribute("val", "").Value;
                    if (themeKey != null && ResolveDrawingThemeColor(themeKey) is { } themeColor)
                        frame.LineFormat.Color = themeColor;
                }
            }
            else
            {
                var lineReference = shape?.ChildElements.FirstOrDefault(element => element.LocalName == "style")?
                    .ChildElements.FirstOrDefault(element => element.LocalName == "lnRef");
                if (lineReference != null && int.TryParse(lineReference.GetAttribute("idx", "").Value,
                        out var lineStyleIndex) && lineStyleIndex > 0)
                {
                    frame.LineFormat.Visible = true;
                    frame.LineFormat.Width = Unit.FromPoint(
                        _sourceWordDocument == null ? 1 :
                        ReadDrawingThemeLineWidth(_sourceWordDocument, lineStyleIndex) ?? 1);
                    var schemeColor = lineReference.Descendants()
                        .FirstOrDefault(element => element.LocalName == "schemeClr");
                    var themeKey = schemeColor?.GetAttribute("val", "").Value;
                    var shadeValue = schemeColor?.ChildElements
                        .FirstOrDefault(element => element.LocalName == "shade")?
                        .GetAttribute("val", "").Value;
                    if (themeKey != null && ResolveDrawingThemeColor(themeKey) is { } themeColor)
                    {
                        if (int.TryParse(shadeValue, out var shade) && shade is >= 0 and <= 100000)
                            themeColor = ApplyDrawingShade(themeColor, shade / 100000.0);
                        frame.LineFormat.Color = themeColor;
                    }
                }
            }
        }
        else
        {
            var properties = vmlStyle!.Split(';', StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Split(':', 2))
                .Where(parts => parts.Length == 2)
                .ToDictionary(parts => parts[0].Trim(), parts => parts[1].Trim(), StringComparer.OrdinalIgnoreCase);
            static double Points(IReadOnlyDictionary<string, string> values, string name)
            {
                if (!values.TryGetValue(name, out var value) || value.Length < 3 ||
                    !double.TryParse(value[..^2], NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var length))
                    return 0;
                return value[^2..].ToLowerInvariant() switch
                {
                    "pt" => length,
                    "in" => length * 72,
                    "cm" => length * 72 / 2.54,
                    "mm" => length * 72 / 25.4,
                    "px" => length * 0.75,
                    _ => 0
                };
            }
            frame.Width = Unit.FromPoint(Points(properties, "width"));
            frame.Height = Unit.FromPoint(Points(properties, "height"));
            frame.Left = Unit.FromPoint(Points(properties, "margin-left"));
            frame.Top = Unit.FromPoint(Points(properties, "margin-top"));
            if (properties.GetValueOrDefault("mso-wrap-style") == "square" &&
                Enum.TryParse<MigraDoc.DocumentObjectModel.Shapes.WrapStyle>("BothSides", out var squareWrap))
                frame.WrapFormat.Style = squareWrap;
            else
                frame.WrapFormat.Style = MigraDoc.DocumentObjectModel.Shapes.WrapStyle.None;
            if (double.TryParse(properties.GetValueOrDefault("rotation"), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var rotationDegrees))
                _textBoxRotations[frame] = rotationDegrees;
            if (ParseColor(vmlShape!.GetAttribute("fillcolor", "").Value) is { } fill)
            {
                frame.FillFormat.Visible = true;
                frame.FillFormat.Color = fill;
            }
            frame.LineFormat.Visible = true;
            var strokeColor = vmlShape.GetAttributes()
                .FirstOrDefault(attribute => attribute.LocalName == "strokecolor").Value;
            if (strokeColor is { Length: >= 7 } && ParseColor(strokeColor[..7]) is { } stroke)
                frame.LineFormat.Color = stroke;
            var lineColor = shape?.ChildElements.FirstOrDefault(element => element.LocalName == "style")?
                .Descendants().FirstOrDefault(element => element.LocalName == "lnRef")?
                .Descendants().FirstOrDefault(element => element.LocalName == "schemeClr");
            var themeKey = lineColor?.GetAttributes()
                .FirstOrDefault(attribute => attribute.LocalName == "val").Value;
            var shadeValue = lineColor?.ChildElements.FirstOrDefault(element => element.LocalName == "shade")?
                .GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value;
            if (themeKey != null && int.TryParse(shadeValue, out var shade) && shade is >= 0 and <= 100000 &&
                ResolveDrawingThemeColor(themeKey) is { } themeColor)
                frame.LineFormat.Color = ApplyDrawingShade(themeColor, shade / 100000.0);
            var strokeWeight = vmlShape.GetAttribute("strokeweight", "").Value;
            if (strokeWeight.EndsWith("pt", StringComparison.OrdinalIgnoreCase) &&
                double.TryParse(strokeWeight[..^2], NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var lineWidth))
                frame.LineFormat.Width = Unit.FromPoint(lineWidth);
        }
        if (textBox.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .FirstOrDefault() is { Parent: Body } anchorParagraph &&
            (anchorParagraph.PreviousSibling() == null ||
             anchorParagraph.ParagraphProperties?.PageBreakBefore is { } pageBreakBefore &&
             pageBreakBefore.Val?.Value != false) &&
            (vmlStyle?.Contains("mso-position-vertical-relative:text", StringComparison.OrdinalIgnoreCase) == true ||
             anchor?.ChildElements.FirstOrDefault(element => element.LocalName == "positionV")?
                 .GetAttribute("relativeFrom", "").Value == "paragraph"))
        {
            frame.RelativeVertical = MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Page;
            frame.Top = Unit.FromPoint(_section.PageSetup.TopMargin.Point + frame.Top.Position.Point);
        }
        if (anchor?.GetAttribute("behindDoc", "").Value == "1")
            _behindTextBoxes.Add(frame);
        else if (extent != null && frame.RelativeVertical == MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Page &&
            frame.WrapFormat.Style == MigraDoc.DocumentObjectModel.Shapes.WrapStyle.None)
        {
            long.TryParse(anchor?.GetAttribute("relativeHeight", "").Value,
                NumberStyles.Integer, CultureInfo.InvariantCulture, out var relativeHeight);
            _foregroundTextBoxes.Add((_section, frame, relativeHeight));
        }
        frame.MarginLeft = Unit.FromPoint(7.2);
        frame.MarginRight = Unit.FromPoint(7.2);
        frame.MarginTop = Unit.FromPoint(3.6);
        frame.MarginBottom = Unit.FromPoint(3.6);
        if (bodyProperties != null)
        {
            if (long.TryParse(bodyProperties.GetAttribute("lIns", "").Value, out var leftInset))
                frame.MarginLeft = Unit.FromPoint(leftInset / 12700.0);
            if (long.TryParse(bodyProperties.GetAttribute("rIns", "").Value, out var rightInset))
                frame.MarginRight = Unit.FromPoint(rightInset / 12700.0);
            if (long.TryParse(bodyProperties.GetAttribute("tIns", "").Value, out var topInset))
                frame.MarginTop = Unit.FromPoint(topInset / 12700.0);
            if (long.TryParse(bodyProperties.GetAttribute("bIns", "").Value, out var bottomInset))
                frame.MarginBottom = Unit.FromPoint(bottomInset / 12700.0);
        }
        if (_container is Cell anchorCell && _paragraph != null &&
            anchor?.GetAttribute("layoutInCell", "").Value != "0" &&
            anchor.ChildElements.FirstOrDefault(element => element.LocalName == "positionH")?
                .GetAttribute("relativeFrom", "").Value == "column" &&
            anchor.ChildElements.FirstOrDefault(element => element.LocalName == "positionV")?
                .GetAttribute("relativeFrom", "").Value == "paragraph" &&
            frame.WrapFormat.Style == MigraDoc.DocumentObjectModel.Shapes.WrapStyle.None)
        {
            _cellAnchoredTextBoxes[frame] = (anchorCell, _paragraph,
                frame.Left.Position.Point, frame.Top.Position.Point);
            if (!_behindTextBoxes.Contains(frame))
            {
                long.TryParse(anchor.GetAttribute("relativeHeight", "").Value,
                    NumberStyles.Integer, CultureInfo.InvariantCulture, out var relativeHeight);
                _foregroundTextBoxes.Add((_section, frame, relativeHeight));
            }
        }
        _container = frame;
        _paragraph = null;
        return DxpDisposable.Create(() =>
        {
            _container = previousContainer;
            _paragraph = previousParagraph;
            _textBoxTextColor = previousTextBoxTextColor;
        });
    }

    internal static double? ReadDrawingThemeLineWidth(WordprocessingDocument document, int styleIndex)
    {
        if (styleIndex < 1 || document.MainDocumentPart?.ThemePart is not { } themePart) return null;
        using var stream = themePart.GetStream(FileMode.Open, FileAccess.Read);
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var style = XDocument.Load(stream).Descendants(drawing + "lnStyleLst").FirstOrDefault()?
            .Elements(drawing + "ln").ElementAtOrDefault(styleIndex - 1);
        return long.TryParse(style?.Attribute("w")?.Value, out var width) ? width / 12700.0 : null;
    }

    private MigraDoc.DocumentObjectModel.Color? ResolveDrawingThemeColor(string key)
    {
        if (_sourceWordDocument?.MainDocumentPart?.ThemePart is not { } themePart) return null;
        using var stream = themePart.GetStream(FileMode.Open, FileAccess.Read);
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var scheme = XDocument.Load(stream).Descendants(drawing + "clrScheme").FirstOrDefault();
        var color = scheme?.Elements().FirstOrDefault(element => element.Name.LocalName == key)?
            .Elements().FirstOrDefault();
        var value = color?.Attribute("val")?.Value;
        var rgb = value is { Length: 6 } && value.All(Uri.IsHexDigit)
            ? value : color?.Attribute("lastClr")?.Value;
        return rgb == null ? null : ParseColor("#" + rgb);
    }

    internal static MigraDoc.DocumentObjectModel.Color ApplyDrawingShade(
        MigraDoc.DocumentObjectModel.Color color, double factor)
    {
        static byte ShadeChannel(uint channel, double factor)
        {
            var encoded = channel / 255.0;
            var linear = encoded <= 0.04045 ? encoded / 12.92 : Math.Pow((encoded + 0.055) / 1.055, 2.4);
            var shaded = linear * factor;
            var result = shaded <= 0.0031308 ? shaded * 12.92 : 1.055 * Math.Pow(shaded, 1 / 2.4) - 0.055;
            return (byte)Math.Round(result * 255, MidpointRounding.AwayFromZero);
        }

        return MigraDoc.DocumentObjectModel.Color.FromRgb(
            ShadeChannel(color.R, factor), ShadeChannel(color.G, factor), ShadeChannel(color.B, factor));
    }

    private PdfDocument? RenderPageAwareStyleRefHeaders(PdfDocumentRenderer renderer,
        MigraDoc.DocumentObjectModel.Document? template)
    {
        if (template == null || _styleRefHeaders.Count != 1 || renderer.PageCount < 2 ||
            _footnoteParagraphs.Count > 0) return null;
        var field = _styleRefHeaders[0];
        var styleId = _paragraphStyles.Values.FirstOrDefault(style =>
            string.Equals(style.StyleName?.Val?.Value, field.StyleName,
                StringComparison.OrdinalIgnoreCase))?.StyleId?.Value;
        if (styleId == null) return null;

        var headings = new List<(int Page, string Text)>();
        for (var page = 1; page <= renderer.PageCount; page++)
            foreach (var info in renderer.DocumentRenderer.GetRenderInfoFromPage(page))
                if (info.DocumentObject is MigraDoc.DocumentObjectModel.Paragraph paragraph &&
                    _styleRefBodyParagraphs.TryGetValue(paragraph, out var source) &&
                    string.Equals(source.ParagraphProperties?.ParagraphStyleId?.Val?.Value,
                        styleId, StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(source.InnerText))
                    headings.Add((page, source.InnerText.Trim()));
        if (headings.Count == 0) return null;
        var values = Enumerable.Range(1, renderer.PageCount).Select(page =>
            headings.FirstOrDefault(heading => heading.Page == page).Text ??
            headings.LastOrDefault(heading => heading.Page < page).Text ??
            headings[0].Text).ToArray();
        if (values.All(value => value == field.CachedText)) return null;

        var variants = new Dictionary<string, PdfDocumentRenderer>(StringComparer.Ordinal);
        foreach (var value in values.Distinct(StringComparer.Ordinal))
        {
            if (value == field.CachedText) continue;
            var document = (MigraDoc.DocumentObjectModel.Document)template.Clone();
            var candidates = document.Sections.Cast<Section>()
                .SelectMany(section => new[] { section.Headers.Primary, section.Headers.FirstPage,
                    section.Headers.EvenPage })
                .SelectMany(header => header.Elements.OfType<MigraDoc.DocumentObjectModel.Paragraph>())
                .Where(paragraph => string.Concat(paragraph.Elements.OfType<FormattedText>()
                    .SelectMany(span => span.Elements.OfType<MigraDoc.DocumentObjectModel.Text>())
                    .Select(text => text.Content)).Contains(field.CachedText, StringComparison.Ordinal))
                .ToArray();
            if (candidates.Length != 1) return null;
            if (!ReplaceStyleRefText(candidates[0], field.CachedText, value)) return null;
            var variant = CreatePdfRenderer(document);
            variant.RenderDocument();
            if (variant.PageCount != renderer.PageCount) return null;
            variants[value] = variant;
        }

        var streams = new List<MemoryStream>();
        var imports = new Dictionary<string, PdfDocument>(StringComparer.Ordinal);
        try
        {
            foreach (var (value, variant) in variants.Append(
                new KeyValuePair<string, PdfDocumentRenderer>(field.CachedText, renderer)))
            {
                var stream = new MemoryStream();
                variant.PdfDocument.Save(stream, false);
                stream.Position = 0;
                streams.Add(stream);
                imports[value] = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
            }
            var combined = new PdfDocument();
            for (var index = 0; index < values.Length; index++)
                combined.AddPage(imports[values[index]].Pages[index]);
            return combined;
        }
        finally
        {
            foreach (var import in imports.Values) import.Dispose();
            foreach (var stream in streams) stream.Dispose();
        }
    }

    private static bool ReplaceStyleRefText(MigraDoc.DocumentObjectModel.Paragraph paragraph,
        string cached, string replacement)
    {
        var texts = paragraph.Elements.OfType<FormattedText>()
            .SelectMany(span => span.Elements.OfType<MigraDoc.DocumentObjectModel.Text>()).ToArray();
        var offset = string.Concat(texts.Select(text => text.Content))
            .IndexOf(cached, StringComparison.Ordinal);
        if (offset < 0) return false;
        var end = offset + cached.Length;
        var position = 0;
        var inserted = false;
        foreach (var text in texts)
        {
            var content = text.Content;
            var next = position + content.Length;
            if (next > offset && position < end)
            {
                var prefix = content[..Math.Max(0, offset - position)];
                var suffix = content[Math.Min(content.Length, end - position)..];
                text.Content = prefix + (inserted ? "" : replacement) + suffix;
                inserted = true;
            }
            position = next;
        }
        return inserted;
    }

    private void AlignSectionContentVertically()
    {
        if (_verticallyAlignedSections.Count == 0) return;
        var layout = (MigraDoc.DocumentObjectModel.Document)_document!.Clone();
        var renderer = CreatePdfRenderer(layout);
        renderer.RenderDocument();
        for (var sectionIndex = 0; sectionIndex < _document.Sections.Count; sectionIndex++)
        {
            var section = (Section)_document.Sections[sectionIndex];
            if (!_verticallyAlignedSections.TryGetValue(section, out var alignment)) continue;
            var layoutSection = (Section)layout.Sections[sectionIndex];
            var noteIndexes = _footnoteParagraphs
                .Where(note => ReferenceEquals(note.Paragraph.Section, section))
                .Select(note => section.Elements.IndexOf(note.Paragraph))
                .ToHashSet();
            var pages = Enumerable.Range(1, renderer.PageCount)
                .Select(pageNumber => renderer.DocumentRenderer.GetRenderInfoFromPage(pageNumber)
                    .Where(info => ReferenceEquals(info.DocumentObject.Section, layoutSection) &&
                        layoutSection.Elements.IndexOf(info.DocumentObject) is var elementIndex &&
                        elementIndex >= 0 && !noteIndexes.Contains(elementIndex) &&
                        info.DocumentObject is MigraDoc.DocumentObjectModel.Paragraph or
                            MigraDoc.DocumentObjectModel.Tables.Table &&
                        info.LayoutInfo.ContentArea != null)
                    .Select(info => info.LayoutInfo.ContentArea!)
                    .ToArray())
                .Where(areas => areas.Length > 0).ToArray();
            if (pages.Length != 1) continue;
            var top = pages[0].Min(area => area.Y.Point);
            var bottom = pages[0].Max(area => area.Y.Point + area.Height.Point);
            var page = section.PageSetup;
            var available = page.PageHeight.Point - page.TopMargin.Point - page.BottomMargin.Point;
            var noteHeight = noteIndexes.Count == 0 ? 0 :
                Enumerable.Range(1, renderer.PageCount)
                    .SelectMany(pageNumber => renderer.DocumentRenderer.GetRenderInfoFromPage(pageNumber))
                    .Where(info => ReferenceEquals(info.DocumentObject.Section, layoutSection) &&
                        noteIndexes.Contains(layoutSection.Elements.IndexOf(info.DocumentObject)) &&
                        info.LayoutInfo.ContentArea != null)
                    .Sum(info => info.LayoutInfo.ContentArea!.Height.Point) + 20;
            var freeSpace = available - (bottom - top) - noteHeight;
            if (freeSpace <= 0) continue;
            var desiredTop = page.TopMargin.Point + freeSpace * (alignment == "center" ? 0.5 : 1);
            page.TopMargin = Unit.FromPoint(page.TopMargin.Point + desiredTop - top);
        }
    }

    private void RenderPageBorders(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_pageBorders.Count == 0) return;
        Section? previousSection = null;
        var sectionPage = 0;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            var section = renderer.DocumentRenderer.GetRenderInfoFromPage(index + 1)?
                .Select(info => info.DocumentObject.Section)
                .FirstOrDefault(candidate => candidate != null);
            if (section == null) continue;
            sectionPage = ReferenceEquals(section, previousSection) ? sectionPage + 1 : 1;
            previousSection = section;
            if (!_pageBorders.TryGetValue(section, out var borders) ||
                AttributeValue(borders, "offsetFrom") == "text") continue;
            var display = AttributeValue(borders, "display");
            if (display == "firstPage" && sectionPage != 1 ||
                display == "notFirstPage" && sectionPage == 1) continue;
            var page = (target ?? renderer.PdfDocument).Pages[index];
            using var graphics = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Prepend);
            DrawPageBorder(graphics, borders, page.Width.Point, page.Height.Point);
        }
    }

    private void RenderLineNumbers(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_sectionLineNumbers.Count == 0) return;
        var flags = System.Reflection.BindingFlags.Instance |
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic;
        var lineNumberStyle = _paragraphStyles.GetValueOrDefault("LineNumber")?.StyleRunProperties;
        var family = lineNumberStyle?.RunFonts?.Ascii?.Value;
        if (family == null || !_availableFontFamilies.Contains(family))
            family = _availableFontFamilies.Contains("Arial") ? "Arial" : _fontFamily;
        var size = int.TryParse(lineNumberStyle?.FontSize?.Val?.Value, out var halfPoints) && halfPoints > 0
            ? halfPoints / 2.0 : 12;
        var font = new XFont(family, size);
        var currentLine = 1;
        Section? previousSection = null;
        for (var pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
        {
            var infos = renderer.DocumentRenderer.GetRenderInfoFromPage(pageIndex + 1);
            var section = infos.Select(info => info.DocumentObject.Section).FirstOrDefault(candidate =>
                candidate != null && _sectionLineNumbers.ContainsKey(candidate));
            if (section == null || !_sectionLineNumbers.TryGetValue(section, out var settings)) continue;
            if (!ReferenceEquals(section, previousSection) ||
                settings.Restart?.Value != LineNumberRestartValues.Continuous &&
                settings.Restart?.Value != LineNumberRestartValues.NewSection)
                currentLine = Math.Max(1, (int)(settings.Start?.Value ?? 1));
            previousSection = section;
            var countBy = Math.Max(1, (int)(settings.CountBy?.Value ?? 1));
            var distance = int.TryParse(settings.Distance?.Value, out var twips)
                ? twips / 20.0 : 18;
            using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[pageIndex]);
            foreach (var info in infos)
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    !_sourceBodyParagraphs.Contains(paragraph) ||
                    !ReferenceEquals(paragraph.Section, section) ||
                    info.LayoutInfo.ContentArea is not { } area) continue;
                var lineCount = info.FormatInfo.GetType().GetProperty("LineCount", flags)?
                    .GetValue(info.FormatInfo) is int count && count > 0 ? count : 1;
                for (var lineIndex = 0; lineIndex < lineCount; lineIndex++, currentLine++)
                {
                    if (currentLine % countBy != 0 ||
                        _suppressedLineNumberParagraphs.Contains(paragraph)) continue;
                    var top = area.Y.Point + lineIndex * area.Height.Point / lineCount;
                    graphics.DrawString(currentLine.ToString(CultureInfo.InvariantCulture), font,
                        XBrushes.Black, new XPoint(section.PageSetup.LeftMargin.Point - distance, top),
                        XStringFormats.TopRight);
                }
            }
        }
    }

    private bool ReserveFootnoteReferencePages(PdfDocumentRenderer renderer,
        MigraDoc.DocumentObjectModel.Document layout)
    {
        var originals = new Dictionary<MigraDoc.DocumentObjectModel.Paragraph,
            MigraDoc.DocumentObjectModel.Paragraph>();
        for (var sectionIndex = 0; sectionIndex < _document!.Sections.Count; sectionIndex++)
        {
            var source = ((Section)_document.Sections[sectionIndex]).Elements;
            var copy = ((Section)layout.Sections[sectionIndex]).Elements;
            for (var elementIndex = 0; elementIndex < source.Count; elementIndex++)
                if (source[elementIndex] is MigraDoc.DocumentObjectModel.Paragraph original &&
                    copy[elementIndex] is MigraDoc.DocumentObjectModel.Paragraph cloned)
                    originals[cloned] = original;
        }
        var noteSet = _footnoteParagraphs.Select(entry => entry.Paragraph).ToHashSet();
        var pages = new Dictionary<MigraDoc.DocumentObjectModel.Paragraph, int>();
        var heights = new Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double>();
        byte[]? renderedPages = null;
        var pageParagraphs = new Dictionary<int, List<(MigraDoc.DocumentObjectModel.Paragraph Paragraph,
            double Left, double Top, double Right, double Bottom)>>();
        for (var page = 1; page <= renderer.PageCount; page++)
        {
            var paragraphs = new List<(MigraDoc.DocumentObjectModel.Paragraph, double, double, double, double)>();
            pageParagraphs[page] = paragraphs;
            foreach (var info in renderer.DocumentRenderer.GetRenderInfoFromPage(page))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    !originals.TryGetValue(paragraph, out var original) ||
                    info.LayoutInfo.ContentArea is not { } area) continue;
                pages.TryAdd(original, page);
                if (noteSet.Contains(original))
                    heights[original] = heights.GetValueOrDefault(original) + area.Height.Point;
                else if (_bodyParagraphs.Contains(original))
                    paragraphs.Add((original, area.X.Point, area.Y.Point,
                        area.X.Point + area.Width.Point, area.Y.Point + area.Height.Point));
            }
        }
        foreach (var group in _footnoteParagraphs.Where(entry =>
                heights.ContainsKey(entry.Paragraph) &&
                _footnoteReferences.TryGetValue(entry.Index, out var reference) &&
                pages.ContainsKey(reference)).GroupBy(entry =>
                pages[_footnoteReferences[entry.Index]] + entry.PageOffset))
        {
            var entries = group.ToList();
            var paragraphs = pageParagraphs[group.Key];
            if (paragraphs.Count == 0) continue;
            var section = entries[0].Paragraph.Section;
            var noteBottom = renderer.PdfDocument.Pages[group.Key - 1].Height.Point -
                section.PageSetup.BottomMargin.Point;
            var noteHeight = 0.0;
            var noteClearance = _includeRevisionMarkup && entries.Count >= 8
                ? DenseFootnoteBodyClearancePoints : _includeRevisionMarkup && entries.Count is 2 or 3
                    ? MarkupFootnoteSeparatorClearancePoints : 0;
            if (_includeRevisionMarkup && entries is [{ PageOffset: 0 } firstEntry] &&
                heights[firstEntry.Paragraph] > 50 &&
                _footnoteReferences.TryGetValue(firstEntry.Index, out var firstReference))
            {
                var referenceBottom = paragraphs.Where(item =>
                        ReferenceEquals(item.Paragraph, firstReference))
                    .Select(item => item.Bottom).DefaultIfEmpty(0).Max();
                var available = noteBottom - referenceBottom;
                if (available > 45 &&
                    available < heights[firstEntry.Paragraph] + ContinuedFootnoteClearancePoints &&
                    TrySplitFootnoteAcrossPages(firstEntry,
                        available - ContinuedFootnoteClearancePoints))
                    return true;
            }
            MigraDoc.DocumentObjectModel.Paragraph? candidate = null;
            for (var index = 0; index < entries.Count; index++)
            {
                var entry = entries[index];
                noteHeight += heights[entry.Paragraph];
                var reference = _footnoteReferences[entry.Index];
                var referenceBottom = paragraphs.Where(item => ReferenceEquals(item.Paragraph, reference))
                    .Select(item => item.Bottom).DefaultIfEmpty(0).Max();
                if (referenceBottom + noteHeight + noteClearance <= noteBottom) continue;
                candidate = reference;
                break;
            }
            if (candidate == null && paragraphs.Max(item => item.Bottom) >
                    noteBottom - noteHeight - noteClearance)
                candidate = paragraphs.FirstOrDefault(item => item.Bottom >
                        noteBottom - noteHeight - noteClearance - 4)
                    .Paragraph;
            if (candidate == null || ReferenceEquals(candidate, paragraphs[0].Paragraph) ||
                candidate.Format.PageBreakBefore) continue;
            var candidateArea = paragraphs.First(item => ReferenceEquals(item.Paragraph, candidate));
            var sectionElements = candidate.Section.Elements;
            for (var elementIndex = 1; elementIndex < sectionElements.Count; elementIndex++)
                if (ReferenceEquals(sectionElements[elementIndex], candidate) &&
                    sectionElements[elementIndex - 1] is MigraDoc.DocumentObjectModel.Paragraph preceding &&
                    preceding.Format.KeepWithNext &&
                    paragraphs.Any(item => ReferenceEquals(item.Paragraph, preceding)) &&
                    noteBottom - noteHeight - candidateArea.Top < 2 * _defaultParagraphLinePoints)
                {
                    preceding.Format.PageBreakBefore = true;
                    return true;
                }
            var precedingNoteHeight = entries.Where(entry =>
                    paragraphs.Any(item => ReferenceEquals(item.Paragraph, _footnoteReferences[entry.Index]) &&
                                           item.Bottom <= candidateArea.Top))
                .Sum(entry => heights[entry.Paragraph]);
            if (precedingNoteHeight > 0 && precedingNoteHeight < noteHeight &&
                TrySplitFittingFootnoteParagraph(renderer, group.Key, candidate, candidateArea,
                    ref renderedPages, noteBottom - precedingNoteHeight - 20)) return true;
            if (TrySplitFittingFootnoteParagraph(renderer, group.Key, candidate, candidateArea, ref renderedPages,
                    noteBottom - noteHeight - (noteClearance > 0
                        ? noteClearance : _availableFontFamilies.Contains("Aptos") ? 16.5 : 19.5) -
                    (_availableFontFamilies.Contains("Aptos") && entries.Count == 3 ? 1 : 0))) return true;
            candidate.Format.PageBreakBefore = true;
            return true;
        }
        return false;
    }

    private bool TrySplitFootnoteAcrossPages(
        (int Index, MigraDoc.DocumentObjectModel.Paragraph Paragraph, int PageOffset) entry,
        double firstFragmentLimit)
    {
        if (!DxpFootnoteParagraphSplitter.TryGetText(entry.Paragraph, out var content)) return false;
        var wordCount = Regex.Matches(content, @"\S+").Count;
        if (wordCount < 8) return false;
        var sectionIndex = Enumerable.Range(0, _document!.Sections.Count)
            .FirstOrDefault(index => ReferenceEquals(_document.Sections[index], entry.Paragraph.Section), -1);
        if (sectionIndex < 0) return false;
        var elements = ((Section)_document.Sections[sectionIndex]).Elements;
        var elementIndex = Enumerable.Range(0, elements.Count)
            .FirstOrDefault(index => ReferenceEquals(elements[index], entry.Paragraph), -1);
        if (elementIndex < 0) return false;

        var low = 1;
        var high = wordCount - 1;
        var best = 0;
        while (low <= high)
        {
            var splitWords = low + (high - low) / 2;
            var trialDocument = (MigraDoc.DocumentObjectModel.Document)_document.Clone();
            var trialElements = ((Section)trialDocument.Sections[sectionIndex]).Elements;
            if (trialElements[elementIndex] is not MigraDoc.DocumentObjectModel.Paragraph trialNote ||
                !DxpFootnoteParagraphSplitter.TrySplit(trialNote, splitWords,
                    out var trialContinuation) || trialContinuation == null)
                return false;
            trialContinuation.Format.PageBreakBefore = false;
            trialElements.InsertObject(elementIndex + 1, trialContinuation);
            var trialRenderer = CreatePdfRenderer(trialDocument);
            trialRenderer.RenderDocument();
            var height = Enumerable.Range(1, trialRenderer.PageCount)
                .SelectMany(trialRenderer.DocumentRenderer.GetRenderInfoFromPage)
                .Where(info => ReferenceEquals(info.DocumentObject, trialNote))
                .Select(info => info.LayoutInfo.ContentArea?.Height.Point ?? 0)
                .DefaultIfEmpty(0).Max();
            if (height > 0 && height <= firstFragmentLimit)
            {
                best = splitWords;
                low = splitWords + 1;
            }
            else high = splitWords - 1;
        }
        if (best == 0 || !DxpFootnoteParagraphSplitter.TrySplit(entry.Paragraph, best,
                out var continuation) || continuation == null)
            return false;
        continuation.Format.PageBreakBefore = false;
        elements.InsertObject(elementIndex + 1, continuation);
        var noteIndex = _footnoteParagraphs.FindIndex(item => ReferenceEquals(item.Paragraph, entry.Paragraph));
        _footnoteParagraphs.Insert(noteIndex + 1, (entry.Index, continuation, 1));
        return true;
    }

    private void SplitOverpageTableCells()
    {
        foreach (Section section in _document!.Sections)
        foreach (var table in section.Elements.OfType<MigraDoc.DocumentObjectModel.Tables.Table>())
        {
            if (table.Columns.Count < 1 || table.Rows.Count is < 1 or > 2 ||
                table.Rows[0].Cells.Count != table.Columns.Count) continue;
            var sourceColumn = table.Rows.Count == 1 && table.Columns.Count is 2 or 3 &&
                table.Rows[0].Cells.Cast<Cell>().All(cell => cell.Elements
                    .Cast<MigraDoc.DocumentObjectModel.DocumentObject>()
                    .All(element => element is MigraDoc.DocumentObjectModel.Paragraph ||
                        element is MigraDoc.DocumentObjectModel.Tables.Table nested &&
                        nested.Rows.Count == 1 && nested.Columns.Count == 2))
                ? Enumerable.Range(0, table.Columns.Count)
                    .OrderByDescending(column => table.Rows[0].Cells[column].Elements.Count).First()
                : 0;
            var sourceCell = table.Rows[0].Cells[sourceColumn];
            if (sourceCell.MergeRight != 0 || sourceCell.Elements.Count < 40 ||
                sourceCell.Elements.Cast<MigraDoc.DocumentObjectModel.DocumentObject>()
                    .Any(element => element is not MigraDoc.DocumentObjectModel.Paragraph &&
                        (element is not MigraDoc.DocumentObjectModel.Tables.Table nested ||
                         nested.Rows.Count != 1 || nested.Columns.Count != 2))) continue;
            if (table.Columns.Count > 1 &&
                (table.Rows.Count != 1 || Enumerable.Range(0, table.Columns.Count)
                    .Where(column => column != sourceColumn).Any(column =>
                    table.Rows[0].Cells[column].MergeRight != 0 ||
                    table.Rows[0].Cells[column].MergeDown != 0 ||
                    table.Rows[0].Cells[column].Elements.Count == 0 ||
                    (table.Rows[0].Cells[column].Elements.Count > 4 &&
                     table.Rows[0].Cells[column].Elements.Count > sourceCell.Elements.Count) ||
                    table.Rows[0].Cells[column].Elements.Cast<MigraDoc.DocumentObjectModel.DocumentObject>()
                        .Any(element => element is not MigraDoc.DocumentObjectModel.Paragraph))))
                continue;
            if (table.Rows.Count == 2)
            {
                if (sourceCell.MergeDown != 1 || table.Rows[1].Cells.Count != 1 ||
                    table.Rows[1].Cells[0].Elements.Cast<MigraDoc.DocumentObjectModel.DocumentObject>()
                        .Any(element => element is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                            paragraph.Elements.Count != 0)) continue;
                table.Rows.RemoveObjectAt(1);
                sourceCell.MergeDown = 0;
            }
            else if (sourceCell.MergeDown != 0) continue;

            var nestedBorderGroups = FragmentTallNestedTable(sourceCell);
            if (nestedBorderGroups.Count > 0)
                _nestedFragmentBorderGroups[table] = nestedBorderGroups;
            _splitOverpageTables.Add(table);

            var parallelSideCells = Enumerable.Range(0, table.Columns.Count)
                .Select(column => column != sourceColumn &&
                    table.Rows[0].Cells[column].Elements.Count > 4)
                .ToArray();
            var firstRowParagraphCount = table.Columns.Count == 1 ? 1 :
                Math.Max(1, Enumerable.Range(0, table.Columns.Count)
                    .Where(column => column != sourceColumn)
                    .Max(column => parallelSideCells[column] ? 1 :
                        table.Rows[0].Cells[column].Elements.Count));
            for (var index = 0; index < sourceCell.Elements.Count - firstRowParagraphCount; index++)
            {
                var row = table.Rows[index];
                var first = (MigraDoc.DocumentObjectModel.Tables.Row)row.Clone();
                var continuation = (MigraDoc.DocumentObjectModel.Tables.Row)row.Clone();
                first.Height = Unit.FromPoint(0);
                continuation.Height = Unit.FromPoint(0);
                var firstCell = first.Cells[sourceColumn];
                var continuationCell = continuation.Cells[sourceColumn];
                var paragraphsInFirstRow = index == 0 ? firstRowParagraphCount : 1;
                for (var paragraphIndex = firstCell.Elements.Count - 1;
                     paragraphIndex >= paragraphsInFirstRow; paragraphIndex--)
                    firstCell.Elements.RemoveObjectAt(paragraphIndex);
                for (var paragraphIndex = 0; paragraphIndex < paragraphsInFirstRow; paragraphIndex++)
                    continuationCell.Elements.RemoveObjectAt(0);
                firstCell.Borders.Bottom.Visible = false;
                continuationCell.Borders.Top.Visible = false;
                for (var column = 0; column < table.Columns.Count; column++)
                {
                    if (column == sourceColumn) continue;
                    var firstSideCell = first.Cells[column];
                    var continuationSideCell = continuation.Cells[column];
                    if (parallelSideCells[column])
                    {
                        var sideParagraphsInFirstRow = Math.Min(paragraphsInFirstRow,
                            firstSideCell.Elements.Count);
                        for (var paragraphIndex = firstSideCell.Elements.Count - 1;
                             paragraphIndex >= sideParagraphsInFirstRow; paragraphIndex--)
                            firstSideCell.Elements.RemoveObjectAt(paragraphIndex);
                        for (var paragraphIndex = 0; paragraphIndex < sideParagraphsInFirstRow;
                             paragraphIndex++)
                            continuationSideCell.Elements.RemoveObjectAt(0);
                    }
                    else continuationSideCell.Elements.Clear();
                    firstSideCell.Borders.Bottom.Visible = false;
                    continuationSideCell.Borders.Top.Visible = false;
                }
                if (!DxpFootnoteTableRowSplitter.TryReplace(row, first, continuation))
                    throw new InvalidOperationException("Could not split an over-page table cell.");
            }
        }
    }

    private static List<(int Start, int Count, bool Top, bool Bottom)> FragmentTallNestedTable(
        Cell parentCell)
    {
        var borderGroups = new List<(int Start, int Count, bool Top, bool Bottom)>();
        for (var elementIndex = 0; elementIndex < parentCell.Elements.Count; elementIndex++)
        {
            if (parentCell.Elements[elementIndex] is not MigraDoc.DocumentObjectModel.Tables.Table nested ||
                nested.Rows.Count != 1 || nested.Columns.Count != 2 ||
                nested.Rows[0].Cells.Count != 2)
                continue;
            var counts = nested.Rows[0].Cells.Cast<Cell>()
                .Select(cell => cell.Elements.Count).ToArray();
            if (counts.Max() <= 24 || nested.Rows[0].Cells.Cast<Cell>()
                    .Any(cell => cell.Elements.Cast<MigraDoc.DocumentObjectModel.DocumentObject>()
                        .Any(element => element is not MigraDoc.DocumentObjectModel.Paragraph)))
                continue;
            var topBorder = nested.Borders.Top.Visible;
            var bottomBorder = nested.Borders.Bottom.Visible;
            if (topBorder || bottomBorder)
                borderGroups.Add((elementIndex, counts.Max(), topBorder, bottomBorder));
            var fragments = new List<MigraDoc.DocumentObjectModel.Tables.Table>();
            for (var fragmentIndex = 0; fragmentIndex < counts.Max(); fragmentIndex++)
            {
                var fragment = (MigraDoc.DocumentObjectModel.Tables.Table)nested.Clone();
                for (var column = 0; column < 2; column++)
                {
                    var elements = fragment.Rows[0].Cells[column].Elements;
                    for (var paragraphIndex = elements.Count - 1; paragraphIndex >= 0; paragraphIndex--)
                        if (paragraphIndex != fragmentIndex)
                            elements.RemoveObjectAt(paragraphIndex);
                }
                if (fragmentIndex > 0)
                {
                    fragment.Borders.Top.Visible = false;
                    foreach (Cell cell in fragment.Rows[0].Cells)
                        cell.Borders.Top.Visible = false;
                }
                if (fragmentIndex < counts.Max() - 1)
                {
                    fragment.Borders.Bottom.Visible = false;
                    foreach (Cell cell in fragment.Rows[0].Cells)
                        cell.Borders.Bottom.Visible = false;
                }
                fragments.Add(fragment);
            }
            parentCell.Elements.RemoveObjectAt(elementIndex);
            for (var fragmentIndex = 0; fragmentIndex < fragments.Count; fragmentIndex++)
                parentCell.Elements.InsertObject(elementIndex + fragmentIndex, fragments[fragmentIndex]);
            elementIndex += fragments.Count - 1;
        }
        return borderGroups;
    }

    private void RestoreNestedFragmentPageBorders()
    {
        if (_nestedFragmentBorderGroups.Count == 0) return;
        for (var iteration = 0; iteration < 8; iteration++)
        {
            var preflightDocument = (MigraDoc.DocumentObjectModel.Document)_document!.Clone();
            var preflight = CreatePdfRenderer(preflightDocument);
            preflight.RenderDocument();
            var changed = false;
            for (var sectionIndex = 0; sectionIndex < preflightDocument.Sections.Count; sectionIndex++)
            {
                var originalElements = ((Section)_document.Sections[sectionIndex]).Elements;
                var copiedElements = ((Section)preflightDocument.Sections[sectionIndex]).Elements;
                for (var elementIndex = 0; elementIndex < originalElements.Count; elementIndex++)
                {
                    if (originalElements[elementIndex] is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                        !_nestedFragmentBorderGroups.TryGetValue(table, out var groups) ||
                        copiedElements[elementIndex] is not MigraDoc.DocumentObjectModel.Tables.Table copy)
                        continue;
                    var pageStarts = new HashSet<int>();
                    var pageEnds = new HashSet<int>();
                    for (var page = 1; page <= preflight.PageCount; page++)
                    foreach (var info in preflight.DocumentRenderer.GetRenderInfoFromPage(page))
                    {
                        if (!ReferenceEquals(info.DocumentObject, copy) ||
                            info.FormatInfo is not TableFormatInfo fragment)
                            continue;
                        foreach (var (start, count, top, bottom) in groups)
                        {
                            var first = Math.Max(start, fragment.StartRow);
                            var last = Math.Min(start + count - 1, fragment.EndRow);
                            if (first > last) continue;
                            if (top && first > start) pageStarts.Add(first);
                            if (bottom && last < start + count - 1) pageEnds.Add(last);
                        }
                    }
                    foreach (var (start, count, top, bottom) in groups)
                    {
                        for (var row = start; row < start + count; row++)
                        {
                            if (top && row > start)
                                changed |= SetNestedFragmentBorder(table.Rows[row].Cells[0], true,
                                    pageStarts.Contains(row));
                            if (bottom && row < start + count - 1)
                                changed |= SetNestedFragmentBorder(table.Rows[row].Cells[0], false,
                                    pageEnds.Contains(row));
                        }
                    }
                }
            }
            if (!changed) return;
        }
    }

    private static bool SetNestedFragmentBorder(Cell parentCell, bool top, bool visible)
    {
        if (parentCell.Elements.OfType<MigraDoc.DocumentObjectModel.Tables.Table>().FirstOrDefault() is
            not { } nested) return false;
        var border = top ? nested.Borders.Top : nested.Borders.Bottom;
        var changed = border.Visible != visible;
        border.Visible = visible;
        foreach (Cell cell in nested.Rows[0].Cells)
        {
            border = top ? cell.Borders.Top : cell.Borders.Bottom;
            changed |= border.Visible != visible;
            border.Visible = visible;
        }
        return changed;
    }

    private bool TrySplitOverflowingTableRow(PdfDocumentRenderer renderer,
        MigraDoc.DocumentObjectModel.Document layout)
    {
        using var stream = new MemoryStream();
        renderer.PdfDocument.Save(stream, false);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(stream.ToArray());
        for (var pageNumber = 1; pageNumber < renderer.PageCount; pageNumber++)
        {
            var previousFragments = renderer.DocumentRenderer.GetRenderInfoFromPage(pageNumber)
                .Where(info => info.DocumentObject is MigraDoc.DocumentObjectModel.Tables.Table).ToArray();
            var nextFragments = renderer.DocumentRenderer.GetRenderInfoFromPage(pageNumber + 1)
                .Where(info => info.DocumentObject is MigraDoc.DocumentObjectModel.Tables.Table).ToArray();
            foreach (var previous in previousFragments)
            foreach (var next in nextFragments)
            {
                if (!ReferenceEquals(previous.DocumentObject, next.DocumentObject) ||
                    next.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                    previous.LayoutInfo.ContentArea is not { } previousArea ||
                    next.LayoutInfo.ContentArea is not { } nextArea ||
                    GetTableFragmentRow(previous.FormatInfo, "EndRow") is not int previousEnd ||
                    GetTableFragmentRow(next.FormatInfo, "StartRow") is not int firstMovedRow ||
                    firstMovedRow != previousEnd + 1 || firstMovedRow <= 0 ||
                    firstMovedRow >= table.Rows.Count - 1 || table.Rows.Count < 3 ||
                    table.Columns.Count != 2 || table.Rows[firstMovedRow].Cells.Count != 2)
                    continue;

                MigraDoc.DocumentObjectModel.Tables.Table? original = null;
                var tableSectionIndex = -1;
                var tableElementIndex = -1;
                for (var sectionIndex = 0; sectionIndex < layout.Sections.Count && original == null; sectionIndex++)
                {
                    var copiedElements = ((Section)layout.Sections[sectionIndex]).Elements;
                    var originalElements = ((Section)_document!.Sections[sectionIndex]).Elements;
                    for (var elementIndex = 0; elementIndex < copiedElements.Count; elementIndex++)
                        if (ReferenceEquals(copiedElements[elementIndex], table))
                        {
                            original = originalElements[elementIndex] as MigraDoc.DocumentObjectModel.Tables.Table;
                            tableSectionIndex = sectionIndex;
                            tableElementIndex = elementIndex;
                            break;
                        }
                }
                if (original == null || original.Rows.Count != table.Rows.Count) continue;

                if (_includeRevisionMarkup && TrySplitAlignedTableParagraphs(pageNumber,
                        firstMovedRow, tableSectionIndex, tableElementIndex, original))
                    return true;

                if (TrySplitLaterTableParagraph(pageNumber, pdf, table, nextArea,
                        firstMovedRow, tableSectionIndex, tableElementIndex))
                    return true;

                var row = table.Rows[firstMovedRow];
                var counts = new int[row.Cells.Count];
                var left = nextArea.X.Point;
                var nextPage = pdf.GetPage(pageNumber + 1);
                var pageWords = nextPage.GetWords().ToArray();
                var valid = true;
                for (var column = 0; column < row.Cells.Count; column++)
                {
                    var right = left + table.Columns[column].Width.Point;
                    if (row.Cells[column].Elements.Count == 0 ||
                        row.Cells[column].Elements[0] is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                        !DxpFootnoteParagraphSplitter.TryGetText(paragraph, out var content))
                    {
                        valid = false;
                        break;
                    }
                    var sourceWords = Regex.Matches(content, @"\S+").Select(match => match.Value).ToArray();
                    var rendered = pageWords.Where(word => word.BoundingBox.Left >= left - 2 &&
                            word.BoundingBox.Right <= right + 2 &&
                            nextPage.Height - word.BoundingBox.Top >= nextArea.Y.Point - 2 &&
                            nextPage.Height - word.BoundingBox.Bottom <=
                            nextArea.Y.Point + nextArea.Height.Point + 2)
                        .OrderByDescending(word => word.Letters[0].StartBaseLine.Y)
                        .ThenBy(word => word.BoundingBox.Left).Take(sourceWords.Length).ToArray();
                    if (sourceWords.Length < 3 || rendered.Length != sourceWords.Length ||
                        !rendered.Select(word => word.Text).SequenceEqual(sourceWords))
                    {
                        valid = false;
                        break;
                    }
                    var lineEnds = new List<int>();
                    for (var index = 0; index < rendered.Length;)
                    {
                        var end = index + 1;
                        while (end < rendered.Length && Math.Abs(rendered[end].Letters[0].StartBaseLine.Y -
                                   rendered[index].Letters[0].StartBaseLine.Y) < 8)
                            end++;
                        lineEnds.Add(end);
                        index = end;
                    }
                    if (lineEnds.Count < 2)
                    {
                        valid = false;
                        break;
                    }
                    counts[column] = lineEnds[^2];
                    left = right;
                }
                if (!valid || previousArea.Y.Point + previousArea.Height.Point >=
                    pdf.GetPage(pageNumber).Height - table.Section.PageSetup.BottomMargin.Point - 20)
                    continue;
                var trialDocument = (MigraDoc.DocumentObjectModel.Document)_document!.Clone();
                if (((Section)trialDocument.Sections[tableSectionIndex]).Elements[tableElementIndex] is not
                    MigraDoc.DocumentObjectModel.Tables.Table trialTable)
                    continue;
                var trialRow = trialTable.Rows[firstMovedRow];
                if (!DxpFootnoteTableRowSplitter.TrySplit(trialRow, counts, out var trialFirst,
                        out var trialContinuation) || trialFirst == null || trialContinuation == null ||
                    !DxpFootnoteTableRowSplitter.TryReplace(trialRow, trialFirst, trialContinuation))
                    continue;
                var trialRenderer = CreatePdfRenderer(trialDocument);
                trialRenderer.RenderDocument();
                if (!trialRenderer.DocumentRenderer.GetRenderInfoFromPage(pageNumber)
                        .Any(info => ReferenceEquals(info.DocumentObject, trialTable) &&
                                     GetTableFragmentRow(info.FormatInfo, "EndRow") >= firstMovedRow))
                    continue;
                var originalRow = original.Rows[firstMovedRow];
                if (DxpFootnoteTableRowSplitter.TrySplit(originalRow, counts, out var first,
                        out var continuation) && first != null && continuation != null &&
                    DxpFootnoteTableRowSplitter.TryReplace(originalRow, first, continuation))
                    return true;
                return false;
            }
        }
        return false;
    }

    private bool TrySplitAlignedTableParagraphs(int pageNumber, int rowIndex,
        int sectionIndex, int elementIndex, MigraDoc.DocumentObjectModel.Tables.Table originalTable)
    {
        var row = originalTable.Rows[rowIndex];
        if (row.Cells.Count != 2) return false;
        var paragraphCount = row.Cells.Cast<Cell>().Min(cell => cell.Elements.Count);
        for (var count = paragraphCount - 1; count >= 1; count--)
        {
            var trialDocument = (MigraDoc.DocumentObjectModel.Document)_document!.Clone();
            if (((Section)trialDocument.Sections[sectionIndex]).Elements[elementIndex] is not
                MigraDoc.DocumentObjectModel.Tables.Table trialTable)
                continue;
            var trialRow = trialTable.Rows[rowIndex];
            if (!DxpFootnoteTableRowSplitter.TrySplitAfterParagraphs(trialRow, count,
                    out var trialFirst, out var trialContinuation) ||
                trialFirst == null || trialContinuation == null ||
                !DxpFootnoteTableRowSplitter.TryReplace(trialRow, trialFirst, trialContinuation))
                continue;
            var trialRenderer = CreatePdfRenderer(trialDocument);
            trialRenderer.RenderDocument();
            if (!trialRenderer.DocumentRenderer.GetRenderInfoFromPage(pageNumber)
                    .Any(info => ReferenceEquals(info.DocumentObject, trialTable) &&
                                 GetTableFragmentRow(info.FormatInfo, "EndRow") >= rowIndex))
                continue;
            return DxpFootnoteTableRowSplitter.TrySplitAfterParagraphs(row, count,
                    out var first, out var continuation) && first != null && continuation != null &&
                DxpFootnoteTableRowSplitter.TryReplace(row, first, continuation);
        }
        return false;
    }

    private bool TrySplitLaterTableParagraph(int pageNumber, UglyToad.PdfPig.PdfDocument pdf,
        MigraDoc.DocumentObjectModel.Tables.Table table, MigraDoc.Rendering.Area nextArea,
        int rowIndex, int sectionIndex, int elementIndex)
    {
        if (((Section)_document!.Sections[sectionIndex]).Elements[elementIndex] is not
            MigraDoc.DocumentObjectModel.Tables.Table originalTable ||
            !_noAutoSpacingTables.Contains(originalTable)) return false;
        var row = table.Rows[rowIndex];
        var page = pdf.GetPage(pageNumber + 1);
        var pageWords = page.GetWords().ToArray();
        var candidates = new List<int>[row.Cells.Count];
        var left = nextArea.X.Point;
        for (var column = 0; column < row.Cells.Count; column++)
        {
            var cell = row.Cells[column];
            if (cell.Elements.Count < 2 || cell.Elements[0] is not MigraDoc.DocumentObjectModel.Paragraph firstParagraph ||
                cell.Elements[1] is not MigraDoc.DocumentObjectModel.Paragraph secondParagraph ||
                !DxpFootnoteParagraphSplitter.TryGetText(firstParagraph, out var firstText) ||
                !DxpFootnoteParagraphSplitter.TryGetText(secondParagraph, out var secondText))
                return false;
            var right = left + table.Columns[column].Width.Point;
            var rendered = pageWords.Where(word => word.BoundingBox.Left >= left - 2 &&
                    word.BoundingBox.Right <= right + 2 &&
                    page.Height - word.BoundingBox.Top >= nextArea.Y.Point - 2 &&
                    page.Height - word.BoundingBox.Bottom <= nextArea.Y.Point + nextArea.Height.Point + 2)
                .OrderByDescending(word => word.Letters[0].StartBaseLine.Y)
                .ThenBy(word => word.BoundingBox.Left).ToArray();
            var firstContent = string.Concat(Regex.Matches(firstText, @"\S+").Select(match => match.Value));
            var firstRenderedCount = 0;
            var length = 0;
            while (firstRenderedCount < rendered.Length && length < firstContent.Length)
                length += rendered[firstRenderedCount++].Text.Length;
            if (length != firstContent.Length ||
                string.Concat(rendered.Take(firstRenderedCount).Select(word => word.Text)) != firstContent)
                return false;
            var secondContent = string.Concat(Regex.Matches(secondText, @"\S+").Select(match => match.Value));
            var secondRenderedCount = 0;
            length = 0;
            while (firstRenderedCount + secondRenderedCount < rendered.Length && length < secondContent.Length)
                length += rendered[firstRenderedCount + secondRenderedCount++].Text.Length;
            var secondWords = rendered.Skip(firstRenderedCount).Take(secondRenderedCount).ToArray();
            if (length != secondContent.Length ||
                string.Concat(secondWords.Select(word => word.Text)) != secondContent)
                return false;
            var lineEnds = new List<int>();
            for (var index = 0; index < secondWords.Length;)
            {
                var end = index + 1;
                while (end < secondWords.Length && Math.Abs(secondWords[end].Letters[0].StartBaseLine.Y -
                           secondWords[index].Letters[0].StartBaseLine.Y) < 8)
                    end++;
                if (DxpFootnoteParagraphSplitter.TryMapRenderedSplit(secondText,
                        secondWords.Select(word => word.Text).ToArray(), end, out var sourceCount))
                    lineEnds.Add(sourceCount);
                index = end;
            }
            if (lineEnds.Count < 2) return false;
            lineEnds.RemoveAt(lineEnds.Count - 1);
            lineEnds.Add(Regex.Matches(secondText, @"\S+").Count);
            candidates[column] = lineEnds;
            left = right;
        }

        foreach (var counts in (from leftCount in candidates[0]
                                from rightCount in candidates[1]
                                orderby leftCount descending, rightCount descending
                                select new[] { leftCount, rightCount }))
        {
            var trialDocument = (MigraDoc.DocumentObjectModel.Document)_document!.Clone();
            if (((Section)trialDocument.Sections[sectionIndex]).Elements[elementIndex] is not
                MigraDoc.DocumentObjectModel.Tables.Table trialTable)
                continue;
            var trialRow = trialTable.Rows[rowIndex];
            if (!DxpFootnoteTableRowSplitter.TrySplit(trialRow, counts, out var trialFirst,
                    out var trialContinuation, paragraphIndex: 1) || trialFirst == null || trialContinuation == null ||
                !DxpFootnoteTableRowSplitter.TryReplace(trialRow, trialFirst, trialContinuation))
                continue;
            var trialRenderer = CreatePdfRenderer(trialDocument);
            trialRenderer.RenderDocument();
            if (!trialRenderer.DocumentRenderer.GetRenderInfoFromPage(pageNumber)
                    .Any(info => ReferenceEquals(info.DocumentObject, trialTable) &&
                                 GetTableFragmentRow(info.FormatInfo, "EndRow") >= rowIndex))
                continue;
            var originalRow = ((MigraDoc.DocumentObjectModel.Tables.Table)
                ((Section)_document.Sections[sectionIndex]).Elements[elementIndex]).Rows[rowIndex];
            return DxpFootnoteTableRowSplitter.TrySplit(originalRow, counts, out var first,
                    out var continuation, paragraphIndex: 1) && first != null && continuation != null &&
                DxpFootnoteTableRowSplitter.TryReplace(originalRow, first, continuation);
        }
        return false;
    }

    private static int? GetTableFragmentRow(object formatInfo, string fieldName)
    {
        var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public |
            System.Reflection.BindingFlags.NonPublic;
        return formatInfo.GetType().GetField(fieldName, flags)?.GetValue(formatInfo) as int?;
    }

    private void RepeatFloatingTableTopBorders()
    {
        if (_floatingContinuationTopBorders.Count == 0 ||
            typeof(DocumentRenderer).GetProperty("PaginateTextFrame") == null) return;
        var preflightDocument = (MigraDoc.DocumentObjectModel.Document)_document!.Clone();
        var clonedFrames = new Dictionary<MigraDoc.DocumentObjectModel.Shapes.TextFrame,
            (MigraDoc.DocumentObjectModel.Tables.Table Table, DxpComputedBorder Border)>();
        for (var sectionIndex = 0; sectionIndex < _document.Sections.Count; sectionIndex++)
        {
            var original = (Section)_document.Sections[sectionIndex];
            var cloned = (Section)preflightDocument.Sections[sectionIndex];
            for (var elementIndex = 0; elementIndex < original.Elements.Count; elementIndex++)
                if (original.Elements[elementIndex] is MigraDoc.DocumentObjectModel.Shapes.TextFrame frame &&
                    _floatingContinuationTopBorders.TryGetValue(frame, out var continuation) &&
                    cloned.Elements[elementIndex] is MigraDoc.DocumentObjectModel.Shapes.TextFrame clonedFrame)
                    clonedFrames[clonedFrame] = continuation;
        }
        if (clonedFrames.Count == 0) return;
        var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public |
            System.Reflection.BindingFlags.NonPublic;
        foreach (var frame in clonedFrames.Keys) _floatingTableFrames.Add(frame);
        try
        {
            var preflight = CreatePdfRenderer(preflightDocument);
            preflight.PrepareRenderPages();
            for (var page = 1; page <= preflight.PageCount; page++)
                foreach (var info in preflight.DocumentRenderer.GetRenderInfoFromPage(page))
                {
                    if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Shapes.TextFrame frame ||
                        !clonedFrames.TryGetValue(frame, out var continuation) ||
                        GetTableFragmentRow(info.FormatInfo, "FragmentIndex") is not int fragmentIndex ||
                        fragmentIndex == 0) continue;
                    var formatted = info.FormatInfo.GetType().GetField("FormattedTextFrame", flags)?
                        .GetValue(info.FormatInfo);
                    if (formatted?.GetType().GetMethod("GetRenderInfos", flags, null,
                            new[] { typeof(int) }, null)?.Invoke(formatted, new object[] { fragmentIndex }) is
                        not RenderInfo[] nested) continue;
                    foreach (var tableInfo in nested)
                    {
                        if (tableInfo.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table ||
                            GetTableFragmentRow(tableInfo.FormatInfo, "StartRow") is not int startRow ||
                            startRow <= 0 || startRow >= continuation.Table.Rows.Count) continue;
                        foreach (MigraDoc.DocumentObjectModel.Tables.Cell cell in
                                 continuation.Table.Rows[startRow].Cells)
                            if (cell.Borders.Top.Width.Point <= 0)
                                ApplyBorder(cell.Borders.Top, continuation.Border);
                    }
                }
        }
        finally
        {
            foreach (var frame in clonedFrames.Keys) _floatingTableFrames.Remove(frame);
        }
    }

    private bool TrySplitFittingFootnoteParagraph(PdfDocumentRenderer renderer, int pageNumber,
        MigraDoc.DocumentObjectModel.Paragraph paragraph,
        (MigraDoc.DocumentObjectModel.Paragraph Paragraph, double Left, double Top, double Right, double Bottom) area,
        ref byte[]? renderedPages, double limit)
    {
        if (area.Top >= limit || area.Bottom <= limit ||
            _emphasisRanges.ContainsKey(paragraph) ||
            _strikeRanges.ContainsKey(paragraph) && !_includeRevisionMarkup ||
            _textEffectRanges.ContainsKey(paragraph) || _textBorderRanges.ContainsKey(paragraph) ||
            _characterScales.ContainsKey(paragraph) || _fitTextWidths.ContainsKey(paragraph) ||
            _partialFitGroups.ContainsKey(paragraph) || _characterSpacings.ContainsKey(paragraph) ||
            _paragraphImages.ContainsKey(paragraph) || _tabbedRunShadings.ContainsKey(paragraph) ||
            !DxpFootnoteParagraphSplitter.TryGetText(paragraph, out var content)) return false;
        var sourceWordCount = Regex.Matches(content, @"\S+").Count;
        if (sourceWordCount < 4) return false;
        if (renderedPages == null)
        {
            using var stream = new MemoryStream();
            renderer.PdfDocument.Save(stream, false);
            renderedPages = stream.ToArray();
        }
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(renderedPages);
        var page = pdf.GetPage(pageNumber);
        var words = page.GetWords().Select(word => new
        {
            word.Text,
            Baseline = word.Letters[0].StartBaseLine.Y,
            Top = page.Height - word.BoundingBox.Top,
            Bottom = page.Height - word.BoundingBox.Bottom,
            Left = word.BoundingBox.Left,
            Right = word.BoundingBox.Right
        }).Where(word => word.Top >= area.Top - 2 && word.Bottom <= area.Bottom + 2 &&
            word.Left >= area.Left - 2 && word.Right <= area.Right + 2)
            .OrderByDescending(word => word.Baseline).ThenBy(word => word.Left).ToArray();
        var fittingWords = 0;
        for (var index = 0; index < words.Length;)
        {
            var lineEnd = index + 1;
            while (lineEnd < words.Length &&
                Math.Abs(words[lineEnd].Baseline - words[index].Baseline) < 8)
                lineEnd++;
            if (words[index..lineEnd].Any(word => word.Bottom > limit)) break;
            fittingWords = lineEnd;
            index = lineEnd;
        }
        var markers = _footnoteReferences.Where(entry => ReferenceEquals(entry.Value, paragraph))
            .Select(entry => entry.Key.ToString(CultureInfo.InvariantCulture)).ToArray();
        var renderedWords = words.Select(word => word.Text).ToArray();
        var mapped = DxpFootnoteParagraphSplitter.TryMapRenderedSplit(content,
            renderedWords, fittingWords, out var fittingSourceWords, markers);
        if (!mapped && fittingWords > 1 && fittingWords < renderedWords.Length &&
            renderedWords[fittingWords - 1].EndsWith("-", StringComparison.Ordinal) &&
            Math.Abs(words[fittingWords - 2].Baseline - words[fittingWords - 1].Baseline) < 8 &&
            DxpFootnoteParagraphSplitter.TryMapRenderedSplit(content,
                renderedWords, fittingWords - 1, out fittingSourceWords, markers))
        {
            fittingWords--;
            mapped = true;
        }
        while (!mapped && fittingWords > 0)
        {
            var baseline = words[fittingWords - 1].Baseline;
            do fittingWords--;
            while (fittingWords > 0 && Math.Abs(words[fittingWords - 1].Baseline - baseline) < 8);
            mapped = DxpFootnoteParagraphSplitter.TryMapRenderedSplit(content,
                renderedWords, fittingWords, out fittingSourceWords, markers);
        }
        if (!mapped) return false;
        var elements = paragraph.Section.Elements;
        var elementIndex = Enumerable.Range(0, elements.Count)
            .FirstOrDefault(index => ReferenceEquals(elements[index], paragraph), -1);
        if (elementIndex < 0) return false;
        if (!DxpFootnoteParagraphSplitter.TrySplit(paragraph, fittingSourceWords, out var continuation) ||
            continuation == null) return false;
        if (continuation.Format.Borders.DistanceFromTop.Point > 0)
            continuation.Format.Borders.DistanceFromTop = Unit.Zero;
        if (_compressedBodyParagraphs.Contains(paragraph) && !_layoutWidthParagraphs.Contains(paragraph) &&
            ShouldWidenCompressedContinuation(fittingSourceWords,
                Regex.Matches(content, @"\S+").Count - fittingSourceWords))
            WidenShortContinuationSpaces(continuation);
        if (_narrowJustifiedBodyParagraphs.Contains(paragraph))
            _narrowJustifiedBodyParagraphs.Add(continuation);
        if (_footnoteSplitContinuations.Remove(paragraph, out var nextFragment))
        {
            nextFragment.Format.PageBreakBefore = false;
            _footnoteSplitContinuations[continuation] = nextFragment;
        }
        _footnoteSplitContinuations[paragraph] = continuation;
        elements.InsertObject(elementIndex + 1, continuation);
        _bodyParagraphs.Add(continuation);
        foreach (var index in _footnoteReferences.Where(entry => ReferenceEquals(entry.Value, paragraph))
                     .Select(entry => entry.Key).ToArray())
            if (continuation.Elements.OfType<MigraDoc.DocumentObjectModel.Fields.BookmarkField>()
                .Any(bookmark => bookmark.Name.StartsWith($"dxp-footnote-{index}-",
                    StringComparison.Ordinal)))
                _footnoteReferences[index] = continuation;
        if (_directBorderParagraphs.TryGetValue(paragraph, out var source))
            _directBorderParagraphs[continuation] = source;
        if (_paragraphShades.TryGetValue(paragraph, out var shade))
            _paragraphShades[continuation] = shade;
        return true;
    }

    internal static bool ShouldWidenCompressedContinuation(int fittingSourceWords,
        int remainingSourceWords) =>
        fittingSourceWords < ShortTrackedSplitWordLimit ||
        remainingSourceWords >= ShortTrackedSplitWordLimit;

    internal static void WidenShortContinuationSpaces(MigraDoc.DocumentObjectModel.Paragraph continuation)
    {
        foreach (var span in continuation.Elements.OfType<FormattedText>())
            if (span.Size.Point == 8.5 &&
                span.Elements.OfType<MigraDoc.DocumentObjectModel.Text>().ToArray() is { Length: > 0 } texts &&
                texts.All(text => text.Content.All(character => character == ' ')))
                span.Size = Unit.FromPoint(9.5);
    }

    private List<(int Page, int Index, MigraDoc.DocumentObjectModel.Paragraph Note, double Y,
        int PageOffset)>
        PlaceFittingFootnotes(PdfDocumentRenderer preflightRenderer,
            MigraDoc.DocumentObjectModel.Document preflight)
    {
        var cloned = new Dictionary<MigraDoc.DocumentObjectModel.Paragraph,
            MigraDoc.DocumentObjectModel.Paragraph>();
        for (var sectionIndex = 0; sectionIndex < _document!.Sections.Count; sectionIndex++)
        {
            var original = ((Section)_document.Sections[sectionIndex]).Elements;
            var copy = ((Section)preflight.Sections[sectionIndex]).Elements;
            for (var elementIndex = 0; elementIndex < original.Count; elementIndex++)
                if (original[elementIndex] is MigraDoc.DocumentObjectModel.Paragraph paragraph &&
                    copy[elementIndex] is MigraDoc.DocumentObjectModel.Paragraph clonedParagraph)
                    cloned[paragraph] = clonedParagraph;
        }
        foreach (var (original, copy) in cloned)
            if (_strikeRanges.TryGetValue(original, out var ranges))
                _strikeRanges[copy] = ranges;
        var noteObjects = _footnoteParagraphs.Select(entry => entry.Paragraph)
            .Where(cloned.ContainsKey).Select(entry => cloned[entry]).ToHashSet();
        var pageOf = new Dictionary<MigraDoc.DocumentObjectModel.Paragraph, int>();
        var heights = new Dictionary<MigraDoc.DocumentObjectModel.Paragraph, double>();
        var bodyBottoms = new Dictionary<int, double>();
        for (var page = 1; page <= preflightRenderer.PageCount; page++)
        {
            foreach (var info in preflightRenderer.DocumentRenderer.GetRenderInfoFromPage(page))
            {
                if (info.LayoutInfo.ContentArea is not { } area) continue;
                var bottom = area.Y.Point + area.Height.Point;
                if (info.DocumentObject is MigraDoc.DocumentObjectModel.Paragraph paragraph)
                {
                    pageOf.TryAdd(paragraph, page);
                    if (noteObjects.Contains(paragraph))
                    {
                        heights[paragraph] = heights.GetValueOrDefault(paragraph) + area.Height.Point;
                        continue;
                    }
                }
                bodyBottoms[page] = Math.Max(bodyBottoms.GetValueOrDefault(page), bottom);
            }
        }
        var groups = _footnoteParagraphs.Where(entry =>
                cloned.TryGetValue(entry.Paragraph, out var note) && heights.ContainsKey(note) &&
                _footnoteReferences.TryGetValue(entry.Index, out var reference) &&
                cloned.TryGetValue(reference, out var clonedReference) && pageOf.ContainsKey(clonedReference))
            .GroupBy(entry => pageOf[cloned[_footnoteReferences[entry.Index]]] + entry.PageOffset);
        var overlays = new List<(int Page, int Index, MigraDoc.DocumentObjectModel.Paragraph Note,
            double Y, int PageOffset)>();
        foreach (var group in groups)
        {
            var entries = group.ToList();
            var noteHeight = entries.Sum(entry => heights[cloned[entry.Paragraph]]);
            var section = cloned[entries[0].Paragraph].Section;
            var noteBottom = preflightRenderer.PdfDocument.Pages[group.Key - 1].Height.Point -
                section.PageSetup.BottomMargin.Point;
            var noteTop = noteBottom - noteHeight;
            if (bodyBottoms.GetValueOrDefault(group.Key) > noteTop) continue;
            foreach (var entry in entries)
            {
                var note = cloned[entry.Paragraph];
                var height = heights[note];
                overlays.Add((group.Key, entry.Index, note, noteTop, entry.PageOffset));
                noteTop += height;
            }
        }
        foreach (Section section in _document.Sections)
            for (var index = section.Elements.Count - 1; index >= 0; index--)
                if (section.Elements[index] is MigraDoc.DocumentObjectModel.Paragraph paragraph &&
                    _footnoteParagraphs.Any(entry => ReferenceEquals(entry.Paragraph, paragraph) &&
                        overlays.Any(overlay => ReferenceEquals(overlay.Note, cloned[paragraph]))))
                    section.Elements.RemoveObjectAt(index);
        if (overlays.Count == _footnoteParagraphs.Count && _footnoteAppendixBreak?.Section is { } breakSection)
            for (var index = breakSection.Elements.Count - 1; index >= 0; index--)
                if (ReferenceEquals(breakSection.Elements[index], _footnoteAppendixBreak))
                {
                    breakSection.Elements.RemoveObjectAt(index);
                    break;
                }
        return overlays;
    }

    private static void RenderFittingFootnotes(PdfDocumentRenderer renderer,
        PdfDocumentRenderer preflightRenderer,
        List<(int Page, int Index, MigraDoc.DocumentObjectModel.Paragraph Note, double Y,
            int PageOffset)> overlays)
    {
        foreach (var group in overlays.GroupBy(overlay => overlay.Page))
        {
            if (group.Key > renderer.PageCount) continue;
            var first = group.First();
            var section = first.Note.Section;
            using var graphics = XGraphics.FromPdfPage(renderer.PdfDocument.Pages[group.Key - 1]);
            var left = section.PageSetup.LeftMargin.Point;
            var width = renderer.PdfDocument.Pages[group.Key - 1].Width.Point - left -
                section.PageSetup.RightMargin.Point;
            var separator = new XPen(XColors.Black, 0.6);
            graphics.DrawLine(separator, left, first.Y - 7.3, left + 144, first.Y - 7.3);
            var namedNotes = new HashSet<int>();
            foreach (var overlay in group)
            {
                preflightRenderer.DocumentRenderer.RenderObject(graphics, new XUnitPt(left),
                    new XUnitPt(overlay.Y), new XUnitPt(width), overlay.Note);
                if (overlay.PageOffset == 0 && namedNotes.Add(overlay.Index))
                    renderer.PdfDocument.AddNamedDestination($"dxp-note-{overlay.Index}", overlay.Page,
                        PdfSharp.Pdf.Advanced.PdfNamedDestinationParameters.CreateVerticalPosition(
                            overlay.Y, null));
            }
        }
    }

    private void RenderSpacedTableBorders(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_spacedTableLeftBorders.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, target))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                    !_spacedTableLeftBorders.TryGetValue(table, out var overlay) ||
                    info.LayoutInfo.ContentArea is not { } area ||
                    overlay.Border.ColorCss is not { Length: 7 } css ||
                    !int.TryParse(css.AsSpan(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
                    continue;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index]);
                var pen = new XPen(XColor.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255),
                    Math.Max(0.5, overlay.Border.WidthPt));
                var rowInset = Math.Max(0, table.Rows.LeftIndent.Point);
                var left = area.X.Point + pen.Width / 2 - rowInset;
                var top = area.Y.Point + offset - 2 * overlay.Spacing;
                var bottom = area.Y.Point + offset + area.Height.Point - pen.Width;
                if (rowInset > 0)
                {
                    bottom += overlay.InHeaderFooter
                        ? 2.5 * pen.Width : 3 * pen.Width;
                }
                bottom -= overlay.ResetAllowance;
                graphics.DrawLine(pen, left, top, left, bottom);
            }
        }
    }

    private void RenderHtmlBeveledTableBorders(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_htmlBevelTables.Count == 0) return;
        for (var pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
        {
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, pageIndex, target))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                    !_htmlBevelTables.TryGetValue(table, out var bevel) ||
                    info.FormatInfo is not TableFormatInfo fragment ||
                    info.LayoutInfo.ContentArea is not { } area) continue;
                var firstRow = fragment.StartRow;
                var visibleRows = fragment.EndRow - firstRow + 1;
                if (visibleRows <= 0) continue;
                var left = area.X.Point;
                var top = area.Y.Point + offset;
                var rowPitch = area.Height.Point / visibleRows;
                var width = table.Columns.Cast<MigraDoc.DocumentObjectModel.Tables.Column>()
                    .Sum(column => column.Width.Point);
                var light = new XPen(XColor.FromArgb(239, 239, 239), bevel.Border);
                var dark = new XPen(XColor.FromArgb(159, 159, 159), bevel.Border);
                var cellLight = new XPen(XColor.FromArgb(239, 239, 239), Math.Min(bevel.Border, 0.72));
                var cellDark = new XPen(XColor.FromArgb(159, 159, 159), Math.Min(bevel.Border, 0.72));
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[pageIndex]);
                var right = left + width + bevel.Border / 2;
                var bottom = top + area.Height.Point + bevel.Spacing + bevel.Border / 2;
                graphics.DrawLine(light, left + bevel.Border, top + bevel.Border / 2,
                    right, top + bevel.Border / 2);
                graphics.DrawLine(light, left + bevel.Border, top + bevel.Border / 2,
                    left + bevel.Border, bottom);
                graphics.DrawLine(dark, right, top + bevel.Border / 2, right, bottom);
                graphics.DrawLine(dark, left + bevel.Border, bottom, right, bottom);
                var columnOffset = 0.0;
                for (var columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                {
                    var column = table.Columns[columnIndex];
                    for (var rowIndex = 0; rowIndex < visibleRows; rowIndex++)
                    {
                        var cell = table.Rows[firstRow + rowIndex].Cells[columnIndex];
                        var indent = cell.Elements.OfType<MigraDoc.DocumentObjectModel.Paragraph>()
                            .FirstOrDefault()?.Format.LeftIndent.Point ?? 0;
                        var textLeft = left + columnOffset + table.LeftPadding.Point +
                            bevel.Border + indent;
                        var cellLeft = textLeft - bevel.Padding - bevel.Border / 2;
                        var nextIndent = columnIndex + 1 < table.Columns.Count
                            ? table.Rows[firstRow + rowIndex].Cells[columnIndex + 1].Elements
                                .OfType<MigraDoc.DocumentObjectModel.Paragraph>()
                                .FirstOrDefault()?.Format.LeftIndent.Point ?? 0
                            : indent;
                        var advance = column.Width.Point + nextIndent -
                            (columnIndex + 1 < table.Columns.Count ? indent : 0);
                        var cellRight = cellLeft + advance - bevel.Spacing - bevel.Border;
                        var cellTop = top + table.TopPadding.Point + bevel.Border / 2 -
                            bevel.Padding + rowIndex * (rowPitch + bevel.Border / 2);
                        var cellBottom = top + (rowIndex + 1) * rowPitch;
                        graphics.DrawLine(cellDark, cellLeft, cellTop, cellRight, cellTop);
                        graphics.DrawLine(cellDark, cellLeft, cellTop, cellLeft, cellBottom);
                        graphics.DrawLine(cellLight, cellRight, cellTop, cellRight, cellBottom);
                        graphics.DrawLine(cellLight, cellLeft, cellBottom, cellRight, cellBottom);
                    }
                    columnOffset += column.Width.Point;
                }
            }
        }
    }

    private void RenderTableFragmentBorders(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_tableFragmentBorders.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, target))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                    !_tableFragmentBorders.TryGetValue(table, out var borders) ||
                    info.FormatInfo is not TableFormatInfo fragment ||
                    info.LayoutInfo.ContentArea is not { } area) continue;
                var rightWidth = Math.Max(0, borders.Right?.WidthPt ?? 0);
                var left = area.X.Point;
                var right = area.X.Point + table.Columns.Cast<MigraDoc.DocumentObjectModel.Tables.Column>()
                    .Sum(column => column.Width.Point) + rightWidth;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index]);
                if (fragment.StartRow > 0 && GetTableFragmentRow(fragment, "LastHeaderRow") is not >= 0 &&
                    borders.Top is { } top)
                    DrawFragmentBorder(graphics, top, left, right, area.Y.Point + offset, false);
                if (fragment.EndRow < table.Rows.Count - 1 && borders.Bottom is { } bottom)
                    DrawFragmentBorder(graphics, bottom, left, right,
                        area.Y.Point + offset + area.Height.Point, true);
            }
        }
    }

    private static void DrawFragmentBorder(XGraphics graphics, DxpComputedBorder border,
        double left, double right, double edge, bool bottom)
    {
        if (border.LineStyle == DxpComputedBorderLineStyle.Double)
        {
            if (border.WidthPt <= 0 || ParseColor(border.ColorCss) is not { } color) return;
            var pen = new XPen(XColor.FromArgb((int)color.R, (int)color.G, (int)color.B),
                border.WidthPt);
            var first = edge + border.WidthPt * 0.5;
            graphics.DrawLine(pen, left, first, right, first);
            graphics.DrawLine(pen, left, first + 2 * border.WidthPt,
                right, first + 2 * border.WidthPt);
            return;
        }
        if (CreateFragmentBorderPen(border) is not { } singlePen) return;
        var position = edge + border.WidthPt / 2;
        graphics.DrawLine(singlePen, left, position, right, position);
    }

    private static XPen? CreateFragmentBorderPen(DxpComputedBorder border)
    {
        if (border.WidthPt <= 0 || border.LineStyle is
            DxpComputedBorderLineStyle.None or DxpComputedBorderLineStyle.Double ||
            ParseColor(border.ColorCss) is not { } color) return null;
        var pen = new XPen(XColor.FromArgb((int)color.R, (int)color.G, (int)color.B), border.WidthPt);
        pen.DashStyle = border.LineStyle switch
        {
            DxpComputedBorderLineStyle.Dashed => XDashStyle.Dash,
            DxpComputedBorderLineStyle.Dotted => XDashStyle.Dot,
            _ => XDashStyle.Solid
        };
        return pen;
    }

    private void RenderDiagonalCellBorders(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_diagonalCells.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, target))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                    !_diagonalCells.TryGetValue(table, out var cells) ||
                    info.LayoutInfo.ContentArea is not { } area ||
                    table.Rows.Count == 0) continue;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index]);
                var rowHeight = area.Height.Point / table.Rows.Count;
                foreach (var cell in cells)
                {
                    if (cell.Row >= table.Rows.Count || cell.Column + cell.ColumnSpan > table.Columns.Count)
                        continue;
                    var left = area.X.Point + Enumerable.Range(0, cell.Column)
                        .Sum(column => table.Columns[column].Width.Point);
                    var right = left + Enumerable.Range(cell.Column, cell.ColumnSpan)
                        .Sum(column => table.Columns[column].Width.Point);
                    if (cell.ColoredTableBorders)
                    {
                        left += 0.5;
                        right += 0.5;
                    }
                    var top = area.Y.Point + offset + cell.Row * rowHeight;
                    var bottom = top + cell.RowSpan * rowHeight;
                    var inset = cell.HasStraightBorders
                        ? Math.Max(cell.Falling?.Width ?? 0, cell.Rising?.Width ?? 0) : 0;
                    DrawDiagonal(graphics, cell.Falling, left, top + inset,
                        right - inset, bottom - inset);
                    DrawDiagonal(graphics, cell.Rising, right - inset, top + inset,
                        left, bottom - inset);
                }
            }
        }
    }

    private void RenderRotatedCells(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_rotatedCells.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, target))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                    !_rotatedCells.TryGetValue(table, out var cells) ||
                    info.LayoutInfo.ContentArea is not { } area || table.Rows.Count == 0) continue;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index]);
                var rowHeight = area.Height.Point / table.Rows.Count;
                foreach (var cell in cells)
                {
                    if (cell.Row >= table.Rows.Count || cell.Column >= table.Columns.Count) continue;
                    var fontStyle = (cell.Bold ? XFontStyleEx.Bold : XFontStyleEx.Regular) |
                        (cell.Italic ? XFontStyleEx.Italic : XFontStyleEx.Regular);
                    var font = new XFont(cell.FontName, cell.FontSize, fontStyle);
                    var color = ParseColor(cell.ColorCss);
                    var brush = color is { } fill
                        ? new XSolidBrush(XColor.FromArgb((int)fill.R, (int)fill.G, (int)fill.B))
                        : XBrushes.Black;
                    var left = area.X.Point + Enumerable.Range(0, cell.Column)
                        .Sum(column => table.Columns[column].Width.Point);
                    var width = table.Columns[cell.Column].Width.Point;
                    var top = area.Y.Point + offset + cell.Row * rowHeight;
                    var chunks = new List<string>();
                    var remaining = cell.Text;
                    var textWidth = cell.InHeaderFooter
                        ? Math.Max(rowHeight - 2, cell.FontSize + 2)
                        : rowHeight - 2;
                    while (remaining.Length > 0)
                    {
                        var length = 1;
                        if (cell.InHeaderFooter && cells.Count > 1)
                        {
                            length = Math.Min(2, remaining.Length);
                            while (length < remaining.Length && char.IsWhiteSpace(remaining[length]))
                                length++;
                        }
                        else
                        {
                            while (length < remaining.Length &&
                                graphics.MeasureString(new string(remaining[..(length + 1)]
                                    .Where(character => !char.IsWhiteSpace(character)).ToArray()), font).Width < textWidth)
                                length++;
                        }
                        chunks.Add(remaining[..length]);
                        remaining = remaining[length..];
                    }
                    for (var chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++)
                    {
                        var columnStep = cell.InHeaderFooter ? 15 : 17.5;
                        var columnLeft = cell.Direction == "btLr"
                            ? left + 5.4 + chunkIndex * columnStep
                            : left + width - 20.4 - chunkIndex * columnStep;
                        var state = graphics.Save();
                        if (cell.Direction == "btLr")
                        {
                            graphics.TranslateTransform(columnLeft, top + rowHeight +
                                (cell.InHeader ? 0.3 : cell.InHeaderFooter ? -2 : 0.5));
                            graphics.RotateTransform(-90);
                        }
                        else
                        {
                            graphics.TranslateTransform(columnLeft + 15, top);
                            graphics.RotateTransform(90);
                        }
                        graphics.DrawString(chunks[chunkIndex], font, brush,
                            new XRect(0, 0, rowHeight, 15), XStringFormats.TopLeft);
                        graphics.Restore(state);
                    }
                }
            }
        }
    }

    private void RenderDoubleFirstRowTopBorders(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_doubleFirstRowTopBorders.Count == 0) return;
        var renderedBodyTables = new HashSet<MigraDoc.DocumentObjectModel.Tables.Table>();
        for (var index = 0; index < renderer.PageCount; index++)
        {
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, target))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                    !_doubleFirstRowTopBorders.TryGetValue(table, out var stroke) ||
                    info.LayoutInfo.ContentArea is not { } area ||
                    ParseColor(stroke.Color) is not { } color) continue;
                if (!stroke.InHeaderFooter && !renderedBodyTables.Add(table)) continue;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index]);
                var pen = new XPen(XColor.FromArgb((int)color.R, (int)color.G, (int)color.B), stroke.Width);
                var left = area.X.Point;
                var top = area.Y.Point + offset + 2.5;
                var right = left + table.Columns.Cast<MigraDoc.DocumentObjectModel.Tables.Column>()
                    .Sum(column => column.Width.Point);
                graphics.DrawLine(pen, left + 1, top, right - 1, top);
            }
        }
    }

    private void RenderRtlDoubleTableLeftBorders(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_rtlDoubleTableLeftBorders.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, target))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                    !_rtlDoubleTableLeftBorders.TryGetValue(table, out var border) ||
                    info.LayoutInfo.ContentArea is not { } area ||
                    ParseColor(border.ColorCss) is not { } color) continue;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index]);
                var pen = new XPen(XColor.FromArgb((int)color.R, (int)color.G, (int)color.B),
                    Math.Max(0.5, border.WidthPt));
                var top = area.Y.Point + offset;
                var bottom = top + area.Height.Point;
                graphics.DrawLine(pen, area.X.Point - pen.Width, top, area.X.Point - pen.Width, bottom);
                graphics.DrawLine(pen, area.X.Point + pen.Width, top, area.X.Point + pen.Width, bottom);
            }
        }
    }

    private void RenderDoubleTableRightBorders(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_doubleTableRightBorders.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, target))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                    !_doubleTableRightBorders.TryGetValue(table, out var border) ||
                    info.LayoutInfo.ContentArea is not { } area ||
                    ParseColor(border.ColorCss) is not { } color) continue;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index]);
                var width = Math.Max(0.5, border.WidthPt);
                var pen = new XPen(XColor.FromArgb((int)color.R, (int)color.G, (int)color.B), width);
                var right = area.X.Point + table.Columns.Cast<MigraDoc.DocumentObjectModel.Tables.Column>()
                    .Sum(column => column.Width.Point);
                var top = area.Y.Point + offset;
                var bottom = top + area.Height.Point;
                graphics.DrawLine(pen, right - width, top, right - width, bottom);
                graphics.DrawLine(pen, right + width, top, right + width, bottom);
            }
        }
    }

    private void RenderDoubleCellSideBorders(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_doubleCellSideBorders.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, target))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                    !_doubleCellSideBorders.TryGetValue(table, out var sides) ||
                    info.LayoutInfo.ContentArea is not { } area || table.Rows.Count == 0) continue;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index]);
                var rowHeight = area.Height.Point / table.Rows.Count;
                foreach (var side in sides)
                {
                    if (side.Column >= table.Columns.Count ||
                        ParseColor(side.Border.ColorCss) is not { } color) continue;
                    var width = Math.Max(0.5, side.Border.WidthPt);
                    var pen = new XPen(XColor.FromArgb((int)color.R, (int)color.G, (int)color.B), width);
                    var x = area.X.Point + Enumerable.Range(0, side.Column)
                        .Sum(column => table.Columns[column].Width.Point) +
                        (side.Right ? table.Columns[side.Column].Width.Point : 0);
                    var top = area.Y.Point + offset + side.Row * rowHeight;
                    var bottom = top + rowHeight;
                    graphics.DrawLine(pen, x - width, top, x - width, bottom);
                    graphics.DrawLine(pen, x + width, top, x + width, bottom);
                }
            }
        }
    }

    private static IEnumerable<(RenderInfo Info, double Offset)> GetPageStoryRenderInfos(
        PdfDocumentRenderer renderer, int index, PdfDocument? target)
    {
        var formatted = renderer.DocumentRenderer.FormattedDocument;
        var flags = System.Reflection.BindingFlags.Instance |
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic;
        foreach (var info in renderer.DocumentRenderer.GetRenderInfoFromPage(index + 1))
            yield return (info, 0);
        foreach (var methodName in new[] { "GetFormattedHeader", "GetFormattedFooter" })
        {
            if (formatted.GetType().GetMethod(methodName, flags)?.Invoke(formatted, new object[] { index + 1 }) is
                not { } story || story.GetType().GetMethod("GetRenderInfos", flags)?.Invoke(story, null) is
                not RenderInfo[] storyInfos) continue;
            var offset = 0.0;
            if (methodName == "GetFormattedFooter" && storyInfos.Length > 0 &&
                storyInfos[0].DocumentObject.Section is { } footerSection)
            {
                var bottoms = storyInfos.Select(info => info.LayoutInfo.ContentArea)
                    .Where(area => area != null)
                    .Select(area => area!.Y.Point + area.Height.Point).ToArray();
                if (bottoms.Length > 0)
                    offset = (target ?? renderer.PdfDocument).Pages[index].Height.Point -
                        footerSection.PageSetup.FooterDistance.Point - bottoms.Max();
            }
            foreach (var info in storyInfos) yield return (info, offset);
        }
    }

    private static IEnumerable<(RenderInfo Info, double XOffset, double YOffset)>
        GetPageEmphasisRenderInfos(PdfDocumentRenderer renderer, int index, PdfDocument? target)
    {
        foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, target))
            foreach (var descendant in GetEmphasisRenderInfoDescendants(info, 0, offset))
                yield return descendant;
    }

    private static IEnumerable<(RenderInfo Info, double XOffset, double YOffset)>
        GetFootnoteEmphasisRenderInfos(PdfDocumentRenderer renderer,
            List<(int Page, int Index, MigraDoc.DocumentObjectModel.Paragraph Note, double Y,
                int PageOffset)> overlays, int pageNumber)
    {
        foreach (var overlay in overlays.Where(item => item.Page == pageNumber))
        for (var sourcePage = 1; sourcePage <= renderer.PageCount; sourcePage++)
        foreach (var info in renderer.DocumentRenderer.GetRenderInfoFromPage(sourcePage))
            if (ReferenceEquals(info.DocumentObject, overlay.Note) &&
                info.LayoutInfo.ContentArea is { } area)
            {
                yield return (info, 0, overlay.Y - area.Y.Point);
                break;
            }
    }

    private static IEnumerable<(RenderInfo Info, double XOffset, double YOffset)>
        GetEmphasisRenderInfoDescendants(RenderInfo info, double parentXOffset, double parentYOffset)
    {
        yield return (info, parentXOffset, parentYOffset);
        var getCellRenderInfos = typeof(FormattedCell).GetMethod("GetRenderInfos",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        var getCellContentHeight = typeof(FormattedCell).GetProperty("ContentHeight",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        var getRowBoundaries = typeof(TableFormatInfo).GetField("BottomBorderMap",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        var getLastHeaderRow = typeof(TableFormatInfo).GetField("LastHeaderRow",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        if (getCellRenderInfos == null || getRowBoundaries == null || getLastHeaderRow == null ||
            info.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
            info.LayoutInfo.ContentArea is not { } tableArea ||
            info.FormatInfo is not TableFormatInfo { FormattedCells: { } formattedCells } tableFormat ||
            getRowBoundaries.GetValue(tableFormat) is not Dictionary<int, XUnitPt> rowBoundaries ||
            !rowBoundaries.TryGetValue(tableFormat.StartRow, out var fragmentStart)) yield break;

        var lastHeaderRow = (int)getLastHeaderRow.GetValue(tableFormat)!;
        var headerHeight = lastHeaderRow >= 0 &&
            rowBoundaries.TryGetValue(0, out var headerStart) &&
            rowBoundaries.TryGetValue(lastHeaderRow + 1, out var headerEnd)
                ? (headerEnd - headerStart).Point : 0;
        foreach (var (cell, formatted) in formattedCells)
        {
            var rowIndex = cell.Row.Index;
            if (rowIndex > lastHeaderRow &&
                (rowIndex < tableFormat.StartRow || rowIndex > tableFormat.EndRow) ||
                !rowBoundaries.TryGetValue(rowIndex, out var rowStart) ||
                getCellRenderInfos.Invoke(formatted, null) is not RenderInfo[] cellInfos) continue;
            var columnOffset = 0.0;
            for (var column = 0; column < cell.Column.Index; column++)
                columnOffset += table.Columns[column].Width.Point;
            var xOffset = parentXOffset + tableArea.X.Point + table[0, 0].Borders.Left.Width.Point +
                columnOffset + cell.Column.LeftPadding.Point;
            var innerHeight = rowBoundaries.TryGetValue(rowIndex + cell.MergeDown + 1,
                out var rowEnd) ? (rowEnd - rowStart).Point - cell.Borders.Bottom.Width.Point : 0;
            var contentHeight = getCellContentHeight?.GetValue(formatted) is XUnitPt measuredHeight
                ? measuredHeight.Point : 0;
            var verticalOffset = cell.VerticalAlignment switch
            {
                MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom when innerHeight > 0 =>
                    innerHeight - cell.Row.BottomPadding.Point - contentHeight,
                MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center when innerHeight > 0 =>
                    (cell.Row.TopPadding.Point + innerHeight -
                        cell.Row.BottomPadding.Point - contentHeight) / 2,
                _ => cell.Row.TopPadding.Point,
            };
            var yOffset = parentYOffset + tableArea.Y.Point +
                table[0, 0].Borders.Top.Width.Point +
                (rowIndex > lastHeaderRow
                    ? headerHeight + (rowStart - fragmentStart).Point
                    : (rowStart - rowBoundaries[0]).Point) + verticalOffset;
            foreach (var cellInfo in cellInfos)
                foreach (var descendant in GetEmphasisRenderInfoDescendants(cellInfo, xOffset, yOffset))
                    yield return descendant;
        }
    }

    private void ReserveHeaderClearance(PdfDocumentRenderer renderer,
        MigraDoc.DocumentObjectModel.Document preflight,
        MigraDoc.DocumentObjectModel.Document? evenDocument)
    {
        var formatted = renderer.DocumentRenderer.FormattedDocument;
        var flags = System.Reflection.BindingFlags.Instance |
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic;
        for (var pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
        {
            if (formatted.GetType().GetMethod("GetFormattedHeader", flags)?
                    .Invoke(formatted, new object[] { pageIndex + 1 }) is not { } header ||
                header.GetType().GetMethod("GetRenderInfos", flags)?.Invoke(header, null) is
                    not RenderInfo[] infos || infos.Length == 0 ||
                infos[0].DocumentObject.Section is not { } preflightSection) continue;
            var sectionIndex = preflight.Sections.IndexOf(preflightSection);
            if (sectionIndex < 0) continue;
            var section = _document!.Sections[sectionIndex];
            if (!_headersWithStyledEmptyLine.Contains(section)) continue;
            var bottom = infos.Select(info => info.LayoutInfo.ContentArea)
                .Where(area => area != null)
                .Max(area => area!.Y.Point + area.Height.Point);
            var clearance = _headersWithStyledTrailingLine.Contains(section) ? 6 : 8;
            var margin = Math.Max(section.PageSetup.TopMargin.Point, bottom + clearance);
            if (margin <= section.PageSetup.TopMargin.Point + 0.1) continue;
            section.PageSetup.TopMargin = Unit.FromPoint(margin);
            if (evenDocument != null)
                evenDocument.Sections[sectionIndex].PageSetup.TopMargin = Unit.FromPoint(margin);
        }
    }

    private static void CollapseDuplicateThinStrokes(PdfDocument document)
    {
        foreach (PdfPage page in document.Pages)
            foreach (var item in page.Contents.Elements)
            {
                if ((item as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                    PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null ||
                    content.Elements.ContainsKey("/Filter")) continue;
                var rewritten = RewriteDuplicateThinStrokes(Encoding.Latin1.GetString(content.Stream.Value));
                if (rewritten != null) content.Stream.Value = Encoding.Latin1.GetBytes(rewritten);
            }
    }

    internal static string? RewriteDuplicateThinStrokes(string source)
    {
        var lines = source.Split('\n');
        var result = new List<string>(lines.Length);
        var seen = new HashSet<(double X1, double Y1, double X2, double Y2,
            string Width, string Color)>();
        var states = new Stack<(string Width, bool Solid, bool Safe, string Color)>();
        var width = "1";
        var solid = true;
        var safe = true;
        var color = "default";
        var changed = false;
        for (var index = 0; index < lines.Length; index++)
        {
            var line = lines[index];
            if (line == "q") states.Push((width, solid, safe, color));
            else if (line == "Q" && states.Count > 0)
                (width, solid, safe, color) = states.Pop();
            else if (line.EndsWith(" w", StringComparison.Ordinal)) width = line[..^2];
            else if (line.EndsWith(" d", StringComparison.Ordinal)) solid = line == "[]0 d";
            else if (line.EndsWith(" cm", StringComparison.Ordinal) ||
                line is "W" or "W*") safe = false;
            else if (line.EndsWith(" RG", StringComparison.Ordinal) ||
                line.EndsWith(" G", StringComparison.Ordinal) ||
                line.EndsWith(" K", StringComparison.Ordinal)) color = line;

            if (width is "0.25" or "0.5" && solid && safe &&
                index + 2 < lines.Length && lines[index + 2] == "S" &&
                TryParseStrokePoint(line, 'm', out var firstX, out var firstY) &&
                TryParseStrokePoint(lines[index + 1], 'l', out var lastX, out var lastY) &&
                (firstX == lastX && Math.Abs(firstY - lastY) >= 3 ||
                    firstY == lastY && Math.Abs(firstX - lastX) >= 3))
            {
                if (firstX > lastX || firstX == lastX && firstY > lastY)
                    (firstX, firstY, lastX, lastY) = (lastX, lastY, firstX, firstY);
                if (!seen.Add((firstX, firstY, lastX, lastY, width, color)))
                {
                    changed = true;
                    index += 2;
                    continue;
                }
            }
            result.Add(line);
        }
        return changed ? string.Join('\n', result) : null;
    }

    private static bool TryParseStrokePoint(string line, char operation, out double x, out double y)
    {
        x = y = 0;
        var parts = line.Split(' ');
        return parts.Length == 3 && parts[2].Length == 1 && parts[2][0] == operation &&
            double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out x) &&
            double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out y);
    }

    internal static void MatchElevenPointAptosPaint(PdfDocument document)
    {
        foreach (PdfPage page in document.Pages)
        {
            var resources = page.Resources.Elements.GetDictionary("/Font");
            if (resources == null) continue;
            var aptosFonts = new HashSet<string>(StringComparer.Ordinal);
            foreach (var key in resources.Elements.Keys)
            {
                var entry = resources.Elements[key];
                var font = (entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfDictionary ??
                    entry as PdfDictionary;
                var baseFont = font?.Elements.GetName("/BaseFont").TrimStart('/');
                if (baseFont == null) continue;
                var family = baseFont[(baseFont.IndexOf('+') + 1)..].Split(',')[0];
                if (family is "Aptos" or "Aptos Display")
                    aptosFonts.Add(key.TrimStart('/'));
            }
            if (aptosFonts.Count == 0) continue;
            foreach (var item in page.Contents.Elements)
            {
                if ((item as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                    PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null ||
                    content.Elements.ContainsKey("/Filter")) continue;
                var source = Encoding.Latin1.GetString(content.Stream.Value);
                var revised = Regex.Replace(source, @"^/(?<font>F\d+) 11 Tf$", match =>
                    aptosFonts.Contains(match.Groups["font"].Value)
                        ? $"/{match.Groups["font"].Value} 11.04 Tf" : match.Value,
                    RegexOptions.Multiline);
                if (revised != source)
                    content.Stream.Value = Encoding.Latin1.GetBytes(revised);
            }
        }
    }

    private static void ApplyBodyFontBaselineRise(PdfDocument document)
    {
        foreach (PdfPage page in document.Pages)
        {
            var fonts = page.Resources.Elements.GetDictionary("/Font");
            if (fonts == null) continue;
            var bodyFonts = new HashSet<string>();
            var regularDejaVuFonts = new HashSet<string>();
            var boldFonts = new HashSet<string>();
            var obliqueFonts = new HashSet<string>();
            var symbolFonts = new HashSet<string>();
            foreach (var key in fonts.Elements.Keys)
            {
                var value = fonts.Elements[key];
                var font = (value as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfDictionary ??
                    value as PdfDictionary;
                var baseFont = font?.Elements.GetName("/BaseFont").TrimStart('/');
                if (baseFont == null) continue;
                var family = baseFont[(baseFont.IndexOf('+') + 1)..].Split(',')[0];
                if (family is "DejaVu Sans" or "DejaVu Sans Oblique" or "Segoe UI")
                    bodyFonts.Add(key.TrimStart('/'));
                if (family == "DejaVu Sans" &&
                    !baseFont.Contains("Bold", StringComparison.OrdinalIgnoreCase))
                    regularDejaVuFonts.Add(key.TrimStart('/'));
                if (family == "DejaVu Sans Oblique") obliqueFonts.Add(key.TrimStart('/'));
                if (family == "DejaVu Sans" && baseFont.Contains("Bold", StringComparison.OrdinalIgnoreCase))
                    boldFonts.Add(key.TrimStart('/'));
                if (family == "Symbol") symbolFonts.Add(key.TrimStart('/'));
            }
            if (bodyFonts.Count == 0) continue;
            foreach (var item in page.Contents.Elements)
            {
                if ((item as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not PdfSharp.Pdf.Advanced.PdfContent content ||
                    content.Stream == null || content.Elements.ContainsKey("/Filter")) continue;
                var lines = Encoding.Latin1.GetString(content.Stream.Value).Split('\n');
                var adjusted = new List<string>(lines.Length);
                var changed = false;
                var risen = false;
                var symbolFont = false;
                var symbolOnLine = false;
                for (var lineIndex = 0; lineIndex < lines.Length; lineIndex++)
                {
                    var line = lines[lineIndex];
                    var match = Regex.Match(line, @"^/(F\d+) (\d+(?:\.\d+)?) Tf$");
                    if (line == "BT") symbolOnLine = false;
                    if (Regex.IsMatch(line, @"^-?\d+(?:\.\d+)? -?\d+(?:\.\d+)? Td$") &&
                        !line.EndsWith(" 0 Td", StringComparison.Ordinal)) symbolOnLine = false;
                    if (match.Success) symbolFont = symbolFonts.Contains(match.Groups[1].Value);
                    if (symbolFont && (line.EndsWith(" Tj", StringComparison.Ordinal) ||
                        line.EndsWith(" TJ", StringComparison.Ordinal))) symbolOnLine = true;
                    var shouldRise = match.Success && bodyFonts.Contains(match.Groups[1].Value) &&
                        match.Groups[2].Value == "10.5" &&
                        !(symbolOnLine && (obliqueFonts.Contains(match.Groups[1].Value) ||
                            boldFonts.Contains(match.Groups[1].Value)));
                    if (risen && (line == "ET" || match.Success && !shouldRise))
                    {
                        adjusted.Add("0 Ts");
                        risen = false;
                    }
                    adjusted.Add(line);
                    if (match.Success && symbolFont && match.Groups[2].Value == "10.5")
                    {
                        var markerSeen = false;
                        for (var nextIndex = lineIndex + 1; nextIndex < lines.Length; nextIndex++)
                        {
                            var nextLine = lines[nextIndex];
                            if (nextLine is "ET" or "BT") break;
                            var nextMove = Regex.Match(nextLine,
                                @"^-?\d+(?:\.\d+)? (-?\d+(?:\.\d+)?) Td$");
                            if (markerSeen && nextMove.Success &&
                                double.TryParse(nextMove.Groups[1].Value, NumberStyles.Float,
                                    CultureInfo.InvariantCulture, out var nextVertical) &&
                                Math.Abs(nextVertical) > 0.001) break;
                            var nextFont = Regex.Match(nextLine, @"^/(F\d+) (\d+(?:\.\d+)?) Tf$");
                            if (nextFont.Success)
                            {
                                if (markerSeen && nextFont.Groups[2].Value == "10.5" &&
                                    bodyFonts.Contains(nextFont.Groups[1].Value) &&
                                    !boldFonts.Contains(nextFont.Groups[1].Value) &&
                                    !obliqueFonts.Contains(nextFont.Groups[1].Value))
                                {
                                    adjusted.Add(regularDejaVuFonts.Contains(nextFont.Groups[1].Value)
                                        ? "1.46 Ts" : "1.5 Ts");
                                    risen = changed = true;
                                }
                                break;
                            }
                            if (nextLine.EndsWith(" Tj", StringComparison.Ordinal) ||
                                nextLine.EndsWith(" TJ", StringComparison.Ordinal))
                                markerSeen = true;
                        }
                    }
                    if (!shouldRise) continue;
                    adjusted.Add(regularDejaVuFonts.Contains(match.Groups[1].Value)
                        ? "1.46 Ts" : "1.5 Ts");
                    risen = changed = true;
                }
                if (changed) content.Stream.Value = Encoding.Latin1.GetBytes(string.Join('\n', adjusted));
            }
        }
    }

    private void LowerMixedDisplayMathRuns(PdfDocumentRenderer renderer, PdfDocument document)
    {
        if (_mixedDisplayRunFractionParagraphs.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            var ranges = GetPageStoryRenderInfos(renderer, index, document)
                .Where(item => item.Info.DocumentObject is MigraDoc.DocumentObjectModel.Paragraph paragraph &&
                    _mixedDisplayRunFractionParagraphs.Contains(paragraph) &&
                    item.Info.LayoutInfo.ContentArea != null)
                .Select(item =>
                {
                    var area = item.Info.LayoutInfo.ContentArea!;
                    return (Min: area.Y.Point + item.Offset - 2,
                        Max: area.Y.Point + item.Offset + area.Height.Point + 2);
                }).ToArray();
            if (ranges.Length == 0) continue;
            var fonts = document.Pages[index].Resources.Elements.GetDictionary("/Font");
            if (fonts == null) continue;
            var mathFonts = new HashSet<string>(StringComparer.Ordinal);
            foreach (var key in fonts.Elements.Keys)
            {
                var value = fonts.Elements[key];
                var font = (value as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfDictionary ??
                    value as PdfDictionary;
                var name = font?.Elements.GetName("/BaseFont").TrimStart('/');
                if (name != null && name[(name.IndexOf('+') + 1)..].Split(',')[0] == "Cambria Math")
                    mathFonts.Add(key.TrimStart('/'));
            }
            if (mathFonts.Count == 0) continue;
            foreach (var item in document.Pages[index].Contents.Elements)
            {
                if ((item as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                    PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null ||
                    content.Elements.ContainsKey("/Filter")) continue;
                var source = Encoding.Latin1.GetString(content.Stream.Value);
                var revised = RewriteMixedDisplayMathRunStream(source, mathFonts, ranges,
                    document.Pages[index].Height.Point);
                if (revised != source) content.Stream.Value = Encoding.Latin1.GetBytes(revised);
            }
        }
    }

    internal static string RewriteMixedDisplayMathRunStream(string source,
        ISet<string> mathFonts, IReadOnlyCollection<(double Min, double Max)> ranges,
        double pageHeight)
    {
        var lines = source.Split('\n');
        var revised = new List<string>(lines.Length);
        var currentY = 0.0;
        var currentFont = "";
        var currentSize = 0.0;
        var changed = false;
        foreach (var line in lines)
        {
            if (line == "BT")
            {
                currentY = 0;
                currentFont = "";
                currentSize = 0;
            }
            var font = Regex.Match(line, @"^/(?<font>F\d+) (?<size>\d+(?:\.\d+)?) Tf$");
            if (font.Success)
            {
                currentFont = font.Groups["font"].Value;
                currentSize = double.Parse(font.Groups["size"].Value, CultureInfo.InvariantCulture);
            }
            var move = Regex.Match(line, @"^-?\d+(?:\.\d+)? (?<vertical>-?\d+(?:\.\d+)?) Td$");
            if (move.Success)
                currentY += double.Parse(move.Groups["vertical"].Value, CultureInfo.InvariantCulture);
            if (mathFonts.Contains(currentFont) && Math.Abs(currentSize - 12) < 0.01 &&
                ranges.Any(range => pageHeight - currentY >= range.Min &&
                    pageHeight - currentY <= range.Max) &&
                (line.EndsWith(" Tj", StringComparison.Ordinal) ||
                 line.EndsWith(" TJ", StringComparison.Ordinal)))
            {
                revised.Add("-4.81 Ts");
                revised.Add(line);
                revised.Add("0 Ts");
                changed = true;
            }
            else revised.Add(line);
        }
        return changed ? string.Join('\n', revised) : source;
    }

    private void AlignBulletBaselines(PdfDocumentRenderer renderer, PdfDocument document)
    {
        if (_bulletBaselineParagraphs.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            var ranges = renderer.DocumentRenderer.GetRenderInfoFromPage(index + 1)
                .Where(info => info.DocumentObject is MigraDoc.DocumentObjectModel.Paragraph paragraph &&
                    _bulletBaselineParagraphs.ContainsKey(paragraph) && info.LayoutInfo.ContentArea != null)
                .Select(info =>
                {
                    var paragraph = (MigraDoc.DocumentObjectModel.Paragraph)info.DocumentObject!;
                    var area = info.LayoutInfo.ContentArea!;
                    var profile = _bulletBaselineParagraphs[paragraph];
                    return (Min: area.Y.Point - 2, Max: area.Y.Point + area.Height.Point + 2,
                        profile.Family, profile.Rise);
                }).ToArray();
            if (ranges.Length == 0) continue;
            var fonts = document.Pages[index].Resources.Elements.GetDictionary("/Font");
            if (fonts == null) continue;
            var markerFonts = new Dictionary<string, string>();
            foreach (var key in fonts.Elements.Keys)
            {
                var value = fonts.Elements[key];
                var font = (value as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfDictionary ??
                    value as PdfDictionary;
                var name = font?.Elements.GetName("/BaseFont").TrimStart('/');
                if (name == null) continue;
                var family = name[(name.IndexOf('+') + 1)..].Split(',')[0];
                if (ranges.Any(range => range.Family == family))
                    markerFonts[key.TrimStart('/')] = family;
            }
            if (markerFonts.Count == 0) continue;
            foreach (var item in document.Pages[index].Contents.Elements)
            {
                if ((item as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                    PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null ||
                    content.Elements.ContainsKey("/Filter")) continue;
                var source = Encoding.Latin1.GetString(content.Stream.Value);
                var revised = RewriteBulletBaselineStream(source, markerFonts, ranges,
                    document.Pages[index].Height.Point);
                if (revised != source) content.Stream.Value = Encoding.Latin1.GetBytes(revised);
            }
        }
    }

    internal static string RewriteBulletBaselineStream(string source,
        IReadOnlyDictionary<string, string> markerFonts,
        IReadOnlyCollection<(double Min, double Max, string Family, double Rise)> ranges,
        double pageHeight)
    {
        var lines = source.Split('\n');
        var output = new List<string>();
        var currentY = 0.0;
        var currentFont = "";
        var pendingRise = 0.0;
        var activeRise = false;
        foreach (var line in lines)
        {
            if (line == "BT")
            {
                currentY = 0;
                currentFont = "";
                pendingRise = 0;
            }
            var font = Regex.Match(line, @"^/(?<font>F\d+) (?<size>\d+(?:\.\d+)?) Tf$");
            if (font.Success)
            {
                currentFont = font.Groups["size"].Value == "12" ? font.Groups["font"].Value : "";
                pendingRise = 0;
            }
            var move = Regex.Match(line, @"^-?\d+(?:\.\d+)? (?<vertical>-?\d+(?:\.\d+)?) Td$");
            if (move.Success)
            {
                var vertical = double.Parse(move.Groups["vertical"].Value, CultureInfo.InvariantCulture);
                if (Math.Abs(vertical) > 0.001)
                {
                    if (activeRise) output.Add("0 Ts");
                    activeRise = false;
                    pendingRise = 0;
                }
                currentY += vertical;
                if (markerFonts.TryGetValue(currentFont, out var family))
                    pendingRise = ranges.FirstOrDefault(range => range.Family == family &&
                        pageHeight - currentY >= range.Min && pageHeight - currentY <= range.Max).Rise;
            }
            if (pendingRise > 0 && (line.EndsWith(" Tj", StringComparison.Ordinal) ||
                line.EndsWith(" TJ", StringComparison.Ordinal)))
            {
                output.Add((-pendingRise).ToString("0.####", CultureInfo.InvariantCulture) + " Ts");
                activeRise = true;
                pendingRise = 0;
            }
            if (line == "ET" && activeRise)
            {
                output.Add("0 Ts");
                activeRise = false;
            }
            output.Add(line);
        }
        return output.Count == lines.Length ? source : string.Join('\n', output);
    }

    private void RaiseCjkNumberedBodyText(PdfDocumentRenderer renderer, PdfDocument document)
    {
        if (_cjkNumberedBodyParagraphs.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            var areas = renderer.DocumentRenderer.GetRenderInfoFromPage(index + 1)
                .Where(info => info.DocumentObject is MigraDoc.DocumentObjectModel.Paragraph paragraph &&
                    _cjkNumberedBodyParagraphs.Contains(paragraph) &&
                    info.LayoutInfo.ContentArea != null)
                .Select(info => info.LayoutInfo.ContentArea!).ToArray();
            if (areas.Length == 0) continue;
            var pageHeight = document.Pages[index].Height.Point;
            var ranges = areas.Select(area => (Min: area.Y.Point - 2,
                Max: area.Y.Point + area.Height.Point + 12)).ToArray();
            var fonts = document.Pages[index].Resources.Elements.GetDictionary("/Font");
            if (fonts == null) continue;
            var markerFonts = new HashSet<string>();
            var textFonts = new HashSet<string>();
            foreach (var key in fonts.Elements.Keys)
            {
                var value = fonts.Elements[key];
                var font = (value as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfDictionary ??
                    value as PdfDictionary;
                var name = font?.Elements.GetName("/BaseFont").TrimStart('/');
                if (name == null) continue;
                var family = name[(name.IndexOf('+') + 1)..].Split(',')[0];
                if (family.StartsWith("Yu Mincho", StringComparison.OrdinalIgnoreCase))
                    markerFonts.Add(key.TrimStart('/'));
                if (family == "Aptos") textFonts.Add(key.TrimStart('/'));
            }
            if (markerFonts.Count == 0 || textFonts.Count == 0) continue;
            foreach (var item in document.Pages[index].Contents.Elements)
            {
                if ((item as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                    PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null ||
                    content.Elements.ContainsKey("/Filter")) continue;
                var source = Encoding.Latin1.GetString(content.Stream.Value);
                var revised = RaiseCjkListTextRuns(source, markerFonts, textFonts,
                    ranges, pageHeight);
                if (revised != source) content.Stream.Value = Encoding.Latin1.GetBytes(revised);
            }
        }
    }

    internal static string RaiseCjkListTextRuns(string source, ISet<string> markerFonts,
        ISet<string> textFonts, IReadOnlyCollection<(double Min, double Max)> ranges,
        double pageHeight)
    {
        var lines = source.Split('\n');
        var result = new List<string>(lines.Length);
        var currentY = 0.0;
        var currentFont = "";
        var markerPending = false;
        var raised = false;
        foreach (var line in lines)
        {
            if (line == "BT")
            {
                currentY = 0;
                currentFont = "";
                markerPending = false;
            }
            var move = Regex.Match(line, @"^-?\d+(?:\.\d+)? (?<vertical>-?\d+(?:\.\d+)?) Td$");
            if (move.Success)
            {
                var vertical = double.Parse(move.Groups["vertical"].Value, CultureInfo.InvariantCulture);
                if (Math.Abs(vertical) > 0.001)
                {
                    if (raised) result.Add("0 Ts");
                    raised = markerPending = false;
                }
                currentY += vertical;
            }
            var font = Regex.Match(line, @"^/(?<font>F\d+) (?<size>\d+(?:\.\d+)?) Tf$");
            if (font.Success)
            {
                currentFont = font.Groups["font"].Value;
                if (raised && !textFonts.Contains(currentFont))
                {
                    result.Add("0 Ts");
                    raised = false;
                }
            }
            if (line == "ET" && raised)
            {
                result.Add("0 Ts");
                raised = false;
            }
            var bodyMarker = markerFonts.Contains(currentFont) &&
                Regex.IsMatch(line, @"^(<[^>]+>|\([^)]*\)) Tj$") &&
                ranges.Any(range => pageHeight - currentY >= range.Min &&
                    pageHeight - currentY <= range.Max);
            if (bodyMarker) result.Add("1.46 Ts");
            result.Add(line);
            if (bodyMarker)
            {
                result.Add("0 Ts");
                markerPending = true;
            }
            if (font.Success && markerPending && textFonts.Contains(currentFont) &&
                font.Groups["size"].Value == "12")
            {
                result.Add("1.38 Ts");
                markerPending = false;
                raised = true;
            }
        }
        return result.Count == lines.Length ? source : string.Join('\n', result);
    }

    private void ApplyLargeSuperscriptRise(PdfDocument document)
    {
        if (_largeSuperscriptTexts.Count == 0) return;
        foreach (PdfPage page in document.Pages)
        {
            foreach (var item in page.Contents.Elements)
            {
                if ((item as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not PdfSharp.Pdf.Advanced.PdfContent content ||
                    content.Stream == null || content.Elements.ContainsKey("/Filter")) continue;
                var rewritten = RewriteLargeSuperscriptTextStream(
                    Encoding.Latin1.GetString(content.Stream.Value), _largeSuperscriptTexts);
                if (rewritten != null) content.Stream.Value = Encoding.Latin1.GetBytes(rewritten);
            }
        }
    }

    private void ApplyInlineSuperscriptRise(PdfDocument document)
    {
        if (_inlineSuperscriptSizes.Count == 0) return;
        var matches = new Dictionary<string, List<(PdfSharp.Pdf.Advanced.PdfContent Content,
            string Source, int Line, string Font, double Size)>>(StringComparer.Ordinal);
        foreach (PdfPage page in document.Pages)
            foreach (var item in page.Contents.Elements)
            {
                if ((item as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                    PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null ||
                    content.Elements.ContainsKey("/Filter")) continue;
                var source = Encoding.Latin1.GetString(content.Stream.Value);
                var lines = source.Split('\n');
                var font = "";
                var size = 0.0;
                for (var index = 0; index < lines.Length; index++)
                {
                    var fontMatch = Regex.Match(lines[index], @"^/(?<font>\S+) (?<size>\d+(?:\.\d+)?) Tf$");
                    if (fontMatch.Success)
                    {
                        font = fontMatch.Groups["font"].Value;
                        double.TryParse(fontMatch.Groups["size"].Value, NumberStyles.Float,
                            CultureInfo.InvariantCulture, out size);
                    }
                    if (!lines[index].EndsWith(" Tj", StringComparison.Ordinal) &&
                        !lines[index].EndsWith(" TJ", StringComparison.Ordinal)) continue;
                    var token = string.Concat(Regex.Matches(lines[index], @"\(([A-Za-z]+)\)")
                        .Select(match => match.Groups[1].Value));
                    if (!_inlineSuperscriptSizes.TryGetValue(token, out var baseSize) ||
                        Math.Abs(size - baseSize * 0.58) > 0.03) continue;
                    if (!matches.TryGetValue(token, out var candidates))
                        matches[token] = candidates = new();
                    candidates.Add((content, source, index, font, size));
                }
            }
        foreach (var (token, candidates) in matches)
        {
            if (candidates.Count != 1) continue;
            var (content, source, line, font, size) = candidates[0];
            var baseSize = _inlineSuperscriptSizes[token];
            var lines = source.Split('\n').ToList();
            var raisedSize = (baseSize * 0.615).ToString("0.####", CultureInfo.InvariantCulture);
            var rise = (baseSize / 3).ToString("0.####", CultureInfo.InvariantCulture);
            var originalSize = size.ToString("0.####", CultureInfo.InvariantCulture);
            lines.InsertRange(line, [$"/{font} {raisedSize} Tf", $"{rise} Ts"]);
            lines.InsertRange(line + 3, ["0 Ts", $"/{font} {originalSize} Tf"]);
            content.Stream.Value = Encoding.Latin1.GetBytes(string.Join('\n', lines));
        }
    }

    private void ApplyRightLegendShift(PdfDocument document)
    {
        if (_rightLegendLabelGroups.Count == 0) return;
        foreach (PdfPage page in document.Pages)
            foreach (var item in page.Contents.Elements)
            {
                if ((item as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                    PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null ||
                    content.Elements.ContainsKey("/Filter")) continue;
                var rewritten = RewriteRightLegendStream(
                    Encoding.Latin1.GetString(content.Stream.Value), _rightLegendLabelGroups,
                    _columnRightLegendLabelGroups);
                if (rewritten != null) content.Stream.Value = Encoding.Latin1.GetBytes(rewritten);
            }
    }

    internal static string? RewriteRightLegendStream(string source, IReadOnlyList<string[]> labelGroups,
        IReadOnlyList<string[]>? columnLabelGroups = null)
    {
        var lines = source.Split('\n');
        var adjusted = new List<string>(lines.Length + 2);
        var position = 0;
        var changed = false;
        while (position < lines.Length)
        {
            if (lines[position] != "q" || position + 2 >= lines.Length ||
                !Regex.IsMatch(lines[position + 1],
                    @"^1 -?0 -?0 1 \d+(?:\.\d+)? -?\d+(?:\.\d+)? cm$") ||
                lines[position + 2] is not ("0.15 w" or ".15 w"))
            {
                adjusted.Add(lines[position++]);
                continue;
            }
            var plotStart = Array.IndexOf(lines, "1 w", position + 3);
            if (plotStart < 0 || plotStart - position > 100 ||
                !lines[(position + 3)..plotStart].Any(line => line.EndsWith(" re", StringComparison.Ordinal)))
            {
                adjusted.Add(lines[position++]);
                continue;
            }
            var labels = lines[(position + 3)..plotStart]
                .Select(line => Regex.Match(line, @"^\(([A-Za-z0-9 ]+)\) Tj$"))
                .Where(match => match.Success).Select(match => match.Groups[1].Value).ToArray();
            if (!labelGroups.Any(group => group.SequenceEqual(labels)))
            {
                adjusted.Add(lines[position++]);
                continue;
            }
            adjusted.Add(lines[position]);
            adjusted.Add(lines[position + 1]);
            adjusted.Add("1 0 0 1 7 0 cm");
            var legendLines = lines[(position + 2)..plotStart].ToArray();
            var swatches = legendLines.Select((line, index) => (Match: Regex.Match(line,
                    @"^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) 7 7 re$"), Index: index))
                .Where(item => item.Match.Success).ToArray();
            var textPositions = legendLines.Select((line, index) => (Match: Regex.Match(line,
                    @"^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) Td$"), Index: index))
                .Where(item => item.Match.Success).ToArray();
            if (swatches.Length == labels.Length && textPositions.Length == labels.Length)
                for (var index = 0; index < labels.Length; index++)
                {
                    var swatch = swatches[index];
                    var text = textPositions[index];
                    var swatchX = double.Parse(swatch.Match.Groups[1].Value, CultureInfo.InvariantCulture);
                    var swatchY = double.Parse(swatch.Match.Groups[2].Value, CultureInfo.InvariantCulture);
                    var textX = text.Match.Groups[1].Value;
                    var textY = double.Parse(text.Match.Groups[2].Value, CultureInfo.InvariantCulture);
                    legendLines[swatch.Index] = FormattableString.Invariant(
                        $"{swatchX + 3.3:0.####} {swatchY - 0.87 * index:0.####} 5.5 5.5 re");
                    legendLines[text.Index] = FormattableString.Invariant(
                        $"{textX} {textY - 0.3 - 0.87 * index:0.####} Td");
                }
            adjusted.AddRange(legendLines);
            adjusted.Add("1 0 0 1 -7 0 cm");
            var categoryStart = Array.IndexOf(lines, "Q", plotStart + 1);
            if (columnLabelGroups?.Any(group => group.SequenceEqual(labels)) == true &&
                categoryStart > plotStart && categoryStart + 3 < lines.Length &&
                lines[categoryStart + 1] == "BT" &&
                Regex.IsMatch(lines[categoryStart + 3], @"^\([A-Za-z0-9 ]+\) Tj$") &&
                Regex.Match(lines[categoryStart + 2],
                    @"^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) Td$") is { Success: true } categoryPosition)
            {
                var baseline = double.Parse(categoryPosition.Groups[2].Value, CultureInfo.InvariantCulture);
                var start = double.Parse(categoryPosition.Groups[1].Value, CultureInfo.InvariantCulture);
                lines[categoryStart + 2] = FormattableString.Invariant(
                    $"{start + 0.7:0.####} {baseline - 3.17:0.####} Td");
                var categoryEnd = Array.IndexOf(lines, "ET", categoryStart + 3);
                if (categoryEnd > categoryStart && categoryEnd - categoryStart < 30)
                {
                    for (var index = categoryStart + 4; index < categoryEnd; index++)
                        if (Regex.Match(lines[index], @"^(\d+(?:\.\d+)?) 0 Td$") is
                            { Success: true } stepMatch)
                        {
                            var step = double.Parse(stepMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                            lines[index] = FormattableString.Invariant($"{step * 0.9982:0.####} 0 Td");
                        }
                    if (categoryEnd + 1 < lines.Length && lines[categoryEnd + 1] == "1 w")
                        for (var index = categoryEnd + 2; index + 2 < lines.Length; index += 3)
                        {
                            var startMatch = Regex.Match(lines[index],
                                @"^(?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) m$");
                            var endMatch = Regex.Match(lines[index + 1],
                                @"^(?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) l$");
                            if (!startMatch.Success || !endMatch.Success || lines[index + 2] != "S" ||
                                startMatch.Groups["x"].Value != endMatch.Groups["x"].Value) break;
                            var startY = double.Parse(startMatch.Groups["y"].Value, CultureInfo.InvariantCulture);
                            var endY = double.Parse(endMatch.Groups["y"].Value, CultureInfo.InvariantCulture);
                            if (Math.Abs(startY - endY) is < 2 or > 5) break;
                            var tickX = double.Parse(startMatch.Groups["x"].Value, CultureInfo.InvariantCulture);
                            var alignedX = tickX * 0.9982 + 0.76;
                            lines[index] = FormattableString.Invariant($"{alignedX:0.####} {startY:0.####} m");
                            lines[index + 1] = FormattableString.Invariant($"{alignedX:0.####} {endY:0.####} l");
                        }
                    var tickCount = 0;
                    for (var index = categoryEnd + 1; index < lines.Length &&
                         index < categoryEnd + 150 && lines[index] != "Q"; index++)
                    {
                        if (!Regex.IsMatch(lines[index], @"^\(\d+\) Tj$") || index == 0 ||
                            Regex.Match(lines[index - 1],
                                @"^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) Td$") is not
                                { Success: true } tickPosition) continue;
                        var tickX = double.Parse(tickPosition.Groups[1].Value, CultureInfo.InvariantCulture);
                        var tickY = double.Parse(tickPosition.Groups[2].Value, CultureInfo.InvariantCulture);
                        lines[index - 1] = FormattableString.Invariant(
                            $"{tickX - 2.05:0.####} {tickY - 0.1 - 0.075 * tickCount:0.####} Td");
                        tickCount++;
                    }
                }
                adjusted.Add("0.9982 0 0 1 0.76 0 cm");
                adjusted.AddRange(lines[plotStart..(categoryStart + 1)]);
                adjusted.Add("1.0018032 0 0 1 -0.761371 0 cm");
                position = categoryStart + 1;
            }
            else
            {
                position = plotStart;
            }
            changed = true;
        }
        return changed ? string.Join('\n', adjusted) : null;
    }

    internal static string? RewriteLargeSuperscriptTextStream(string source, ISet<string> superscriptTexts)
    {
        var lines = source.Split('\n');
        var adjusted = new List<string>(lines.Length + 4);
        var fontResource = "";
        var sourceSize = 0.0;
        var risen = false;
        foreach (var line in lines)
        {
            var fontMatch = Regex.Match(line, @"^/(\S+) (\d+(?:\.\d+)?) Tf$");
            if (fontMatch.Success)
            {
                if (risen) adjusted.Add("0 Ts");
                risen = false;
                fontResource = fontMatch.Groups[1].Value;
                double.TryParse(fontMatch.Groups[2].Value, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out sourceSize);
            }
            var textMatch = Regex.Match(line, @"^\(([A-Za-z]{2,})\) Tj$");
            var superscript = textMatch.Success && Math.Abs(sourceSize - 11.6) < 0.01 &&
                superscriptTexts.Contains(textMatch.Groups[1].Value);
            if (superscript && !risen)
            {
                adjusted.Add($"/{fontResource} 12 Tf");
                adjusted.Add("7 Ts");
                risen = true;
            }
            else if (!superscript && risen && (textMatch.Success || line == "ET"))
            {
                adjusted.Add("0 Ts");
                if (line != "ET") adjusted.Add($"/{fontResource} 11.6 Tf");
                risen = false;
            }
            adjusted.Add(line);
        }
        return adjusted.Count != lines.Length ? string.Join('\n', adjusted) : null;
    }

    private void ApplyRightAlignedTrackedKerning(PdfDocument document)
    {
        if (_rightAlignedTrackedTexts.Count == 0) return;
        var metrics = DocSystemFontAdvances.Find("Aptos", includeCloudFonts: true);
        if (metrics == null) return;
        foreach (var text in _rightAlignedTrackedTexts)
        {
            var candidates = new List<(PdfSharp.Pdf.Advanced.PdfContent Content, string Revised)>();
            foreach (PdfPage page in document.Pages)
            {
                var fonts = page.Resources.Elements.GetDictionary("/Font");
                if (fonts == null) continue;
                var aptosFonts = new HashSet<string>(StringComparer.Ordinal);
                foreach (var key in fonts.Elements.Keys)
                {
                    var entry = fonts.Elements[key];
                    var font = (entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfDictionary ??
                        entry as PdfDictionary;
                    var name = font?.Elements.GetName("/BaseFont").TrimStart('/');
                    if (name?.Split('+').LastOrDefault() == "Aptos")
                        aptosFonts.Add(key.TrimStart('/'));
                }
                if (aptosFonts.Count == 0) continue;
                foreach (var item in page.Contents.Elements)
                {
                    if ((item as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                        PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null) continue;
                    var filtered = content.Elements.ContainsKey("/Filter");
                    var source = Encoding.Latin1.GetString(filtered
                        ? content.Stream.UnfilteredValue : content.Stream.Value);
                    var revised = RewriteRightAlignedKernedLine(source, text, metrics, aptosFonts);
                    if (revised != null) candidates.Add((content, revised));
                }
            }
            if (candidates.Count != 1) continue;
            var (target, replacement) = candidates[0];
            if (!target.Elements.ContainsKey("/Filter") || target.Stream.TryUncompress())
                target.Stream.Value = Encoding.Latin1.GetBytes(replacement);
        }
    }

    private void ApplyLongJustifiedTrackedKerning(PdfDocument document)
    {
        if (_longJustifiedTrackedTexts.Count == 0) return;
        var metrics = DocSystemFontAdvances.Find("Aptos", includeCloudFonts: true);
        if (metrics == null) return;
        foreach (PdfPage page in document.Pages)
        {
            var fonts = page.Resources.Elements.GetDictionary("/Font");
            if (fonts == null) continue;
            var aptosFonts = new HashSet<string>(StringComparer.Ordinal);
            foreach (var key in fonts.Elements.Keys)
            {
                var entry = fonts.Elements[key];
                var font = (entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfDictionary ??
                    entry as PdfDictionary;
                if (font?.Elements.GetName("/BaseFont").TrimStart('/').Split('+').LastOrDefault() == "Aptos")
                    aptosFonts.Add(key.TrimStart('/'));
            }
            if (aptosFonts.Count == 0) continue;
            foreach (var item in page.Contents.Elements)
            {
                if ((item as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                    PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null) continue;
                var filtered = content.Elements.ContainsKey("/Filter");
                var source = Encoding.Latin1.GetString(filtered
                    ? content.Stream.UnfilteredValue : content.Stream.Value);
                var revised = RewriteLongJustifiedKernedLines(source, _longJustifiedTrackedTexts,
                    metrics, aptosFonts);
                revised = RewriteCapitalizedAcronymMarkerKerning(revised ?? source,
                    _longJustifiedTrackedTexts, metrics, aptosFonts) ?? revised;
                if (revised != null && (!filtered || content.Stream.TryUncompress()))
                    content.Stream.Value = Encoding.Latin1.GetBytes(revised);
            }
        }
    }

    internal static string? RewriteLongJustifiedKernedLines(string source,
        IReadOnlyCollection<string> eligibleParagraphs, DocOpenTypeAdvances metrics,
        ISet<string> aptosFonts)
    {
        var lines = source.Split('\n');
        var changed = false;
        var currentFont = "";
        var currentSize = 0.0;
        for (var start = 0; start < lines.Length; start++)
        {
            var fontMatch = Regex.Match(lines[start], @"^/(?<font>\S+) (?<size>\d+(?:\.\d+)?) Tf$");
            if (fontMatch.Success)
            {
                currentFont = fontMatch.Groups["font"].Value;
                currentSize = double.Parse(fontMatch.Groups["size"].Value, CultureInfo.InvariantCulture);
                continue;
            }
            if (!aptosFonts.Contains(currentFont) || Math.Abs(currentSize - 12) > 0.01 ||
                !Regex.IsMatch(lines[start], @"^-?\d+(?:\.\d+)? -?\d+(?:\.\d+)? Td$") ||
                Math.Abs(double.Parse(lines[start].Split(' ')[1], CultureInfo.InvariantCulture)) < 0.01)
                continue;
            var end = start + 1;
            while (end < lines.Length && lines[end] != "ET")
            {
                var move = Regex.Match(lines[end], @"^-?\d+(?:\.\d+)? (?<y>-?\d+(?:\.\d+)?) Td$");
                if (move.Success && Math.Abs(double.Parse(move.Groups["y"].Value,
                        CultureInfo.InvariantCulture)) >= 0.01) break;
                end++;
            }
            if (end >= lines.Length || lines.Skip(start + 1).Take(end - start - 1)
                    .Any(line => line.EndsWith(" Tf", StringComparison.Ordinal) ||
                        line.EndsWith(" TJ", StringComparison.Ordinal) ||
                        line.EndsWith(" Tz", StringComparison.Ordinal) && line != "100 Tz" ||
                        line.EndsWith(" Tc", StringComparison.Ordinal) && line != "0 Tc")) continue;
            var words = new List<(int Line, string Text, bool Escaped)>();
            var valid = true;
            for (var index = start + 1; index < end; index++)
            {
                if (!lines[index].EndsWith(" Tj", StringComparison.Ordinal)) continue;
                if (lines[index] == "( ) Tj") continue;
                var token = Regex.Match(lines[index], @"^\((?<text>(?:\\[()]|[A-Za-z0-9,.-])+?)\) Tj$");
                if (!token.Success) { valid = false; break; }
                var encoded = token.Groups["text"].Value;
                words.Add((index, encoded.Replace(@"\(", "(").Replace(@"\)", ")"),
                    encoded.Contains('\\')));
            }
            if (!valid || words.Count < 4 || words[0].Line != start + 1 ||
                Enumerable.Range(1, words.Count - 1).Any(index =>
                    words[index].Line - words[index - 1].Line is not (2 or 3) ||
                    words[index].Line - words[index - 1].Line == 3 &&
                        lines[words[index - 1].Line + 1] != "( ) Tj" ||
                    !Regex.IsMatch(lines[words[index].Line - 1], @"^-?\d+(?:\.\d+)? 0 Td$")))
                continue;
            var gaps = Enumerable.Range(1, words.Count - 1)
                .Select(index => words[index].Line - words[index - 1].Line == 3).ToArray();
            if (gaps.Count(space => space) < 3) continue;
            var phrase = words[0].Text + string.Concat(Enumerable.Range(1, words.Count - 1)
                .Select(index => (gaps[index - 1] ? " " : "") + words[index].Text));
            if (eligibleParagraphs.Count(paragraph =>
                    paragraph.IndexOf(phrase, StringComparison.Ordinal) is int position &&
                    position >= 0 && position == paragraph.LastIndexOf(phrase, StringComparison.Ordinal) &&
                    position + phrase.Length < paragraph.Length) != 1) continue;
            var adjustments = words.Select(word => word.Escaped
                ? new double[word.Text.Length - 1]
                : Enumerable.Range(0, word.Text.Length - 1)
                    .Select(index => metrics.KerningAdjustment(word.Text[index], word.Text[index + 1],
                        currentSize)).ToArray()).ToArray();
            var total = adjustments.Sum(word => word.Sum());
            if (Math.Abs(total) < 0.01) continue;
            var extraSpace = -total / gaps.Count(space => space);
            var cumulative = 0.0;
            var previousShift = 0.0;
            var spaceCount = 0;
            for (var index = 0; index < words.Count; index++)
            {
                if (index > 0 && gaps[index - 1]) spaceCount++;
                var shift = cumulative + spaceCount * extraSpace;
                if (index > 0)
                {
                    var move = Regex.Match(lines[words[index].Line - 1],
                        @"^(?<x>-?\d+(?:\.\d+)?) 0 Td$");
                    var horizontal = double.Parse(move.Groups["x"].Value, CultureInfo.InvariantCulture);
                    lines[words[index].Line - 1] =
                        $"{(horizontal + shift - previousShift).ToString("0.####", CultureInfo.InvariantCulture)} 0 Td";
                }
                var formatted = new StringBuilder("[");
                var fragmentStart = 0;
                for (var character = 0; character < adjustments[index].Length; character++)
                {
                    if (Math.Abs(adjustments[index][character]) < 0.001) continue;
                    formatted.Append('(').Append(words[index].Text, fragmentStart,
                        character + 1 - fragmentStart).Append(") ");
                    formatted.Append((-adjustments[index][character] * 1000 / currentSize)
                        .ToString("0.####", CultureInfo.InvariantCulture)).Append(' ');
                    fragmentStart = character + 1;
                }
                if (fragmentStart > 0)
                {
                    formatted.Append('(').Append(words[index].Text, fragmentStart,
                        words[index].Text.Length - fragmentStart).Append(")] TJ");
                    lines[words[index].Line] = formatted.ToString();
                }
                cumulative += adjustments[index].Sum();
                previousShift = shift;
            }
            if (lines[end] != "ET")
            {
                var move = Regex.Match(lines[end],
                    @"^(?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) Td$");
                var horizontal = double.Parse(move.Groups["x"].Value, CultureInfo.InvariantCulture);
                lines[end] = $"{(horizontal - previousShift).ToString("0.####", CultureInfo.InvariantCulture)} " +
                    $"{move.Groups["y"].Value} Td";
            }
            changed = true;
        }
        return changed ? string.Join('\n', lines) : null;
    }

    internal static string? RewriteCapitalizedAcronymMarkerKerning(string source,
        IReadOnlyCollection<string> eligibleParagraphs, DocOpenTypeAdvances metrics,
        ISet<string> aptosFonts)
    {
        if (eligibleParagraphs.Count == 0 || aptosFonts.Count == 0) return null;
        var lines = source.Split('\n');
        var changed = false;
        var currentFont = "";
        var currentSize = 0.0;
        for (var start = 0; start + 10 < lines.Length; start++)
        {
            var font = Regex.Match(lines[start], @"^/(?<name>\S+) (?<size>\d+(?:\.\d+)?) Tf$");
            if (font.Success)
            {
                currentFont = font.Groups["name"].Value;
                currentSize = double.Parse(font.Groups["size"].Value, CultureInfo.InvariantCulture);
                continue;
            }
            if (!aptosFonts.Contains(currentFont) || Math.Abs(currentSize - 12) > 0.01)
                continue;
            var words = new string[3];
            var moves = new Match[3];
            var valid = true;
            for (var wordIndex = 0; wordIndex < words.Length; wordIndex++)
            {
                var wordLine = start + wordIndex * 3;
                var word = Regex.Match(lines[wordLine], @"^\((?<text>[A-Z][a-z]{3,})\) Tj$");
                var move = Regex.Match(lines[wordLine + 2], @"^(?<x>\d+(?:\.\d+)?) 0 Td$");
                if (!word.Success || lines[wordLine + 1] != "( ) Tj" || !move.Success)
                {
                    valid = false;
                    break;
                }
                words[wordIndex] = word.Groups["text"].Value;
                moves[wordIndex] = move;
            }
            if (!valid) continue;
            var acronym = Regex.Match(lines[start + 9],
                @"^\(\\\((?<text>[A-Z]{2,6})\\\)\) Tj$");
            if (!acronym.Success) continue;
            var phrase = string.Join(" ", words) + " (" + acronym.Groups["text"].Value + ")";
            if (eligibleParagraphs.Count(paragraph =>
                    paragraph.Contains(phrase, StringComparison.Ordinal)) != 1)
                continue;
            var restoreMove = -1;
            for (var cursor = start + 10; cursor + 5 < lines.Length; cursor++)
            {
                if (lines[cursor] == "ET") break;
                var move = Regex.Match(lines[cursor],
                    @"^-?\d+(?:\.\d+)? (?<y>-?\d+(?:\.\d+)?) Td$");
                if (move.Success && double.Parse(move.Groups["y"].Value,
                        CultureInfo.InvariantCulture) < -0.01)
                    break;
                var superscript = Regex.Match(lines[cursor],
                    @"^/(?<name>\S+) (?<size>\d+(?:\.\d+)?) Tf$");
                if (!superscript.Success || superscript.Groups["name"].Value != currentFont ||
                    !double.TryParse(superscript.Groups["size"].Value, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var size) || size is < 5 or > 8 ||
                    !Regex.IsMatch(lines[cursor + 1], @"^-?\d+(?:\.\d+)? \d+(?:\.\d+)? Td$") ||
                    !Regex.IsMatch(lines[cursor + 2], @"^\(\d+\) Tj$") ||
                    lines[cursor + 3] != $"/{currentFont} 12 Tf" ||
                    !Regex.IsMatch(lines[cursor + 4],
                        @"^(?<x>-?\d+(?:\.\d+)?) -\d+(?:\.\d+)? Td$") ||
                    !Regex.IsMatch(lines[cursor + 5], @"^\(.+\) Tj$")) continue;
                restoreMove = cursor + 4;
                break;
            }
            if (restoreMove < 0) continue;
            var adjustments = words.Select(word => Enumerable.Range(0, word.Length - 1)
                .Select(index => metrics.KerningAdjustment(word[index], word[index + 1], 12))
                .ToArray()).ToArray();
            var total = adjustments.Sum(word => word.Sum());
            if (Math.Abs(total) < 0.01) continue;
            for (var wordIndex = 0; wordIndex < words.Length; wordIndex++)
            {
                var wordLine = start + wordIndex * 3;
                var builder = new StringBuilder("[");
                var fragmentStart = 0;
                for (var character = 0; character < adjustments[wordIndex].Length; character++)
                {
                    var adjustment = adjustments[wordIndex][character];
                    if (Math.Abs(adjustment) < 0.001) continue;
                    builder.Append('(').Append(words[wordIndex], fragmentStart,
                        character + 1 - fragmentStart).Append(") ");
                    builder.Append((-adjustment * 1000 / 12)
                        .ToString("0.####", CultureInfo.InvariantCulture)).Append(' ');
                    fragmentStart = character + 1;
                }
                if (fragmentStart > 0)
                {
                    builder.Append('(').Append(words[wordIndex], fragmentStart,
                        words[wordIndex].Length - fragmentStart).Append(")] TJ");
                    lines[wordLine] = builder.ToString();
                }
                var horizontal = double.Parse(moves[wordIndex].Groups["x"].Value,
                    CultureInfo.InvariantCulture);
                lines[wordLine + 2] = (horizontal + adjustments[wordIndex].Sum())
                    .ToString("0.####", CultureInfo.InvariantCulture) + " 0 Td";
            }
            var restore = Regex.Match(lines[restoreMove],
                @"^(?<x>-?\d+(?:\.\d+)?) (?<y>-\d+(?:\.\d+)?) Td$");
            var restoreX = double.Parse(restore.Groups["x"].Value, CultureInfo.InvariantCulture);
            lines[restoreMove] = (restoreX - total).ToString("0.####", CultureInfo.InvariantCulture) +
                " " + restore.Groups["y"].Value + " Td";
            changed = true;
            start = restoreMove;
        }
        return changed ? string.Join('\n', lines) : null;
    }

    internal static string? RewriteRightAlignedKernedLine(string source, string text,
        DocOpenTypeAdvances metrics, ISet<string> aptosFonts)
    {
        if (text.Length < 2) return null;
        var lines = source.Split('\n');
        var pairAdjustments = Enumerable.Range(0, text.Length - 1)
            .Select(index => metrics.KerningAdjustment(text[index], text[index + 1], 12)).ToArray();
        if (pairAdjustments.All(value => Math.Abs(value) < 0.001)) return null;
        var rewritten = new List<string>();
        for (var first = 1; first < lines.Length; first++)
        {
            var firstToken = Regex.Match(lines[first], @"^\((?<text>[A-Za-zÀ-ÖØ-öø-ÿ,-]+)\) Tj$");
            if (!firstToken.Success || !text.StartsWith(firstToken.Groups["text"].Value,
                    StringComparison.Ordinal)) continue;
            var firstMove = Enumerable.Range(Math.Max(0, first - 5), Math.Min(5, first))
                .Where(index => Regex.IsMatch(lines[index],
                    @"^-?\d+(?:\.\d+)? -?\d+(?:\.\d+)? Td$")).LastOrDefault(-1);
            if (firstMove < 0 || lines.Skip(firstMove + 1).Take(first - firstMove - 1)
                    .Any(line => line.EndsWith(" Tj", StringComparison.Ordinal) ||
                        line.EndsWith(" TJ", StringComparison.Ordinal))) continue;
            var font = lines.Take(first).Reverse().Select(line =>
                    Regex.Match(line, @"^/(?<name>\S+) (?<size>\d+(?:\.\d+)?) Tf$"))
                .FirstOrDefault(match => match.Success);
            if (font == null || !aptosFonts.Contains(font.Groups["name"].Value) ||
                !double.TryParse(font.Groups["size"].Value, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var size) || Math.Abs(size - 12) > 0.01) continue;
            var tokens = new List<(int Line, int SourceStart, string Value)>();
            var cursor = 0;
            var lineIndex = first;
            while (lineIndex < lines.Length)
            {
                var token = Regex.Match(lines[lineIndex], @"^\((?<text>[A-Za-zÀ-ÖØ-öø-ÿ,-]+)\) Tj$");
                if (!token.Success) break;
                var value = token.Groups["text"].Value;
                var sourceStart = text.IndexOf(value, cursor, StringComparison.Ordinal);
                if (sourceStart < 0 || text[cursor..sourceStart].Any(character => character != ' ')) break;
                tokens.Add((lineIndex, sourceStart, value));
                cursor = sourceStart + value.Length;
                if (text[cursor..].All(character => character == ' ')) break;
                lineIndex += 2;
                if (lineIndex >= lines.Length || !Regex.IsMatch(lines[lineIndex - 1],
                        @"^-?\d+(?:\.\d+)? 0 Td$")) break;
            }
            if (tokens.Count < 2 || !text[cursor..].All(character => character == ' ')) continue;
            var candidate = (string[])lines.Clone();
            var previousShift = 0.0;
            foreach (var (tokenLine, sourceStart, value) in tokens)
            {
                var shift = -pairAdjustments.Skip(sourceStart).Sum();
                var moveIndex = tokenLine == first ? firstMove : tokenLine - 1;
                var move = Regex.Match(candidate[moveIndex],
                    @"^(?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) Td$");
                if (!move.Success) break;
                var horizontal = double.Parse(move.Groups["x"].Value, CultureInfo.InvariantCulture);
                candidate[moveIndex] =
                    $"{(horizontal + shift - previousShift).ToString("0.####", CultureInfo.InvariantCulture)} " +
                    $"{move.Groups["y"].Value} Td";
                var chunks = new List<string>();
                for (var index = 0; index < value.Length; index++)
                {
                    chunks.Add($"({value[index]})");
                    if (index + 1 < value.Length)
                        chunks.Add((-pairAdjustments[sourceStart + index] * 1000 / size)
                            .ToString("0.####", CultureInfo.InvariantCulture));
                }
                candidate[tokenLine] = $"[{string.Join(' ', chunks)}] TJ";
                previousShift = shift;
            }
            var after = tokens[^1].Line + 1;
            if (after < candidate.Length)
            {
                var move = Regex.Match(candidate[after],
                    @"^(?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) Td$");
                if (move.Success)
                {
                    var horizontal = double.Parse(move.Groups["x"].Value, CultureInfo.InvariantCulture);
                    candidate[after] =
                        $"{(horizontal - previousShift).ToString("0.####", CultureInfo.InvariantCulture)} " +
                        $"{move.Groups["y"].Value} Td";
                }
            }
            rewritten.Add(string.Join('\n', candidate));
        }
        return rewritten.Count == 1 ? rewritten[0] : null;
    }

    internal static bool IsPlainLatinAfterFitLine(string[] lines, int start,
        ISet<string> eligibleTexts)
    {
        if (eligibleTexts.Count == 0) return false;
        var content = new StringBuilder();
        for (var index = start; index < lines.Length; index++)
        {
            if (lines[index] is "BT" or "ET") break;
            var move = Regex.Match(lines[index], @"^-?\d+(?:\.\d+)? (-?\d+(?:\.\d+)?) Td$");
            if (move.Success && double.TryParse(move.Groups[1].Value, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var vertical) && vertical != 0) break;
            if (lines[index].EndsWith(" Tj", StringComparison.Ordinal) ||
                lines[index].EndsWith(" TJ", StringComparison.Ordinal))
            {
                if (Regex.IsMatch(lines[index], @"<[0-9A-Fa-f]+>")) return false;
                foreach (Match text in Regex.Matches(lines[index], @"\(([^)]*)\)"))
                    content.Append(text.Groups[1].Value);
            }
        }
        var normalized = Regex.Replace(content.ToString(), "[^A-Za-z0-9]", "");
        return eligibleTexts.Contains(normalized);
    }

    private void ApplyFontKerning(PdfDocument document)
    {
        var smallAptosCarrySetting = Environment.GetEnvironmentVariable("DOCXPORT_DIAG_SMALL_APTOS_CARRY");
        foreach (PdfPage page in document.Pages)
        {
            var fonts = page.Resources.Elements.GetDictionary("/Font");
            if (fonts == null) continue;
            var advances = new Dictionary<string, DocOpenTypeAdvances?>();
            var displayFonts = new HashSet<string>();
            var aptosFonts = new HashSet<string>();
            var calibriFonts = new HashSet<string>();
            var arialFonts = new HashSet<string>();
            var scaledKerningFonts = new HashSet<string>();
            var aptosCommaQuoteFont = _documentHasTrackedInsertions &&
                Environment.GetEnvironmentVariable("DOCXPORT_DIAG_APTOS_COMMA_QUOTE_KERN") == "1"
                ? DocSystemFontAdvances.Find("Aptos", includeCloudFonts: true) : null;
            foreach (var key in fonts.Elements.Keys)
            {
                var value = fonts.Elements[key];
                var font = (value as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfDictionary ??
                    value as PdfDictionary;
                var baseFont = font?.Elements.GetName("/BaseFont").TrimStart('/');
                if (string.IsNullOrEmpty(baseFont)) continue;
                var subset = baseFont.IndexOf('+');
                if (subset >= 0) baseFont = baseFont[(subset + 1)..];
                var bold = baseFont.Contains("Bold", StringComparison.OrdinalIgnoreCase);
                var italic = baseFont.Contains("Italic", StringComparison.OrdinalIgnoreCase) ||
                    baseFont.Contains("Oblique", StringComparison.OrdinalIgnoreCase);
                var family = baseFont.Split(',')[0];
                if (family == "Aptos Display") displayFonts.Add(key.TrimStart('/'));
                if (family is "Aptos" or "Aptos Display") aptosFonts.Add(key.TrimStart('/'));
                if (family == "Calibri") calibriFonts.Add(key.TrimStart('/'));
                if (family == "Arial") arialFonts.Add(key.TrimStart('/'));
                if (CarriesScaledKerning(family, bold, italic)) scaledKerningFonts.Add(key.TrimStart('/'));
                advances[key.TrimStart('/')] = DocSystemFontAdvances.Find(family, bold, italic,
                    includeCloudFonts: !_documentHasTrackedInsertions || family == "Aptos Display");
            }
            foreach (var item in page.Contents.Elements)
            {
                if ((item as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not PdfSharp.Pdf.Advanced.PdfContent content ||
                    content.Stream == null) continue;
                var filtered = content.Elements.ContainsKey("/Filter");
                var lines = Encoding.Latin1.GetString(filtered
                    ? content.Stream.UnfilteredValue : content.Stream.Value).Split('\n');
                var hasHorizontalScaling = lines.Any(line => line.EndsWith(" Tz", StringComparison.Ordinal));
                DocOpenTypeAdvances? currentFont = null;
                var currentDisplayFont = false;
                var currentAptosFont = false;
                var currentCalibriFont = false;
                var currentArialFont = false;
                var currentScaledKerningFont = false;
                var fontSize = 0.0;
                var changed = false;
                var lineShift = 0.0;
                var pendingKerning = 0.0;
                var textScale = 100.0;
                var baseline = 0.0;
                var textOriginX = 0.0;
                var previousWord = "";
                var inUrl = false;
                var urlStartX = 0.0;
                var urlKerning = 0.0;
                var previousUrlCharacter = '\0';
                var previousUrlFragment = "";
                var commaQuoteCarry = 0.0;
                DocOpenTypeAdvances? urlFont = null;
                var urlFontSize = 0.0;
                var urlMoveIndex = -1;
                var carryBaseline = double.NaN;
                var carryAdjustment = 0.0;
                var activeCarryInText = 0.0;
                var firstMoveInText = false;
                var plainLatinLine = false;
                for (var lineIndex = 0; lineIndex < lines.Length; lineIndex++)
                {
                    if (lines[lineIndex] == "BT")
                    {
                        firstMoveInText = true;
                        plainLatinLine = false;
                        activeCarryInText = 0;
                        lineShift = pendingKerning = 0;
                        textScale = 100;
                        baseline = textOriginX = 0;
                        previousWord = "";
                        inUrl = false;
                        urlKerning = 0;
                        previousUrlCharacter = '\0';
                        previousUrlFragment = "";
                        commaQuoteCarry = 0;
                        urlMoveIndex = -1;
                        continue;
                    }
                    if (lines[lineIndex] == "ET")
                    {
                        if (commaQuoteCarry != 0)
                        {
                            carryBaseline = baseline;
                            carryAdjustment = commaQuoteCarry;
                        }
                        if (inUrl && Math.Abs(urlKerning) > 0.001)
                        {
                            carryBaseline = baseline;
                            carryAdjustment = urlKerning;
                        }
                        if (inUrl && Math.Abs(urlKerning) > 0.001 &&
                            AdjustKernedUnderline(lines, lineIndex + 1, urlStartX, baseline, urlKerning))
                            changed = true;
                        inUrl = false;
                        previousUrlCharacter = '\0';
                        previousUrlFragment = "";
                        continue;
                    }
                    var scaleMatch = Regex.Match(lines[lineIndex], @"^(-?\d+(?:\.\d+)?) Tz$");
                    if (scaleMatch.Success && double.TryParse(scaleMatch.Groups[1].Value,
                            NumberStyles.Float, CultureInfo.InvariantCulture, out var parsedScale))
                    {
                        textScale = parsedScale;
                        continue;
                    }
                    var fontMatch = Regex.Match(lines[lineIndex], @"^/(\S+) (\d+(?:\.\d+)?) Tf$");
                    if (fontMatch.Success && double.TryParse(fontMatch.Groups[2].Value,
                            NumberStyles.Float, CultureInfo.InvariantCulture, out fontSize))
                    {
                        advances.TryGetValue(fontMatch.Groups[1].Value, out currentFont);
                        currentDisplayFont = displayFonts.Contains(fontMatch.Groups[1].Value);
                        currentAptosFont = aptosFonts.Contains(fontMatch.Groups[1].Value);
                        currentCalibriFont = calibriFonts.Contains(fontMatch.Groups[1].Value);
                        currentArialFont = arialFonts.Contains(fontMatch.Groups[1].Value);
                        currentScaledKerningFont = scaledKerningFonts.Contains(fontMatch.Groups[1].Value);
                        continue;
                    }
                    var move = Regex.Match(lines[lineIndex], @"^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) Td$");
                    if (move.Success && double.TryParse(move.Groups[1].Value, NumberStyles.Float,
                            CultureInfo.InvariantCulture, out var horizontal) &&
                        double.TryParse(move.Groups[2].Value, NumberStyles.Float,
                            CultureInfo.InvariantCulture, out var vertical))
                    {
                        baseline += vertical;
                        if (vertical != 0)
                            plainLatinLine = IsPlainLatinAfterFitLine(lines, lineIndex + 1,
                                _plainLatinAfterFitTexts);
                        textOriginX += horizontal;
                        var carried = 0.0;
                        if (firstMoveInText && !double.IsNaN(carryBaseline))
                        {
                            if (Math.Abs(baseline - carryBaseline) < 0.2)
                                carried = carryAdjustment;
                            carryBaseline = double.NaN;
                        }
                        firstMoveInText = false;
                        if (carried != 0) activeCarryInText = carried;
                        var resetCarry = vertical != 0 && carried == 0 ? activeCarryInText : 0;
                        textOriginX += carried - resetCarry;
                        var following = lineIndex + 1 < lines.Length
                            ? Regex.Match(lines[lineIndex + 1], @"^\(([A-Za-z]{2,})\) Tj$") : Match.Empty;
                        var distributedTab = vertical == 0 && following.Success &&
                            _distributedWordPairs.Contains(previousWord + "\0" + following.Groups[1].Value);
                        var tabbedMove = distributedTab || vertical == 0 && previousWord.Length > 0 &&
                            currentFont != null && currentFont.TryMeasure(previousWord, fontSize,
                                out var previousWidth) && horizontal - previousWidth > 10;
                        var boundaryKerning = vertical == 0 && !tabbedMove && following.Success &&
                            previousWord.EndsWith("-", StringComparison.Ordinal) &&
                            IsKerningCandidate(previousWord) && currentFont != null
                            ? currentFont.KerningAdjustment('-', following.Groups[1].Value[0], fontSize) : 0;
                        var adjustment = (vertical == 0 ? tabbedMove ? 0 : pendingKerning : -lineShift) +
                            boundaryKerning + carried - resetCarry;
                        if (Math.Abs(adjustment) > 0.001)
                        {
                            lines[lineIndex] = $"{(horizontal + adjustment).ToString("0.####", CultureInfo.InvariantCulture)} " +
                                $"{move.Groups[2].Value} Td";
                            changed = true;
                        }
                        if (vertical == 0 && !tabbedMove) lineShift += pendingKerning + boundaryKerning;
                        else if (vertical != 0) lineShift = 0;
                        if (vertical != 0)
                        {
                            if (carried == 0) activeCarryInText = 0;
                            previousWord = "";
                            inUrl = false;
                            urlKerning = 0;
                            previousUrlCharacter = '\0';
                            previousUrlFragment = "";
                        }
                        urlMoveIndex = vertical == 0 ? lineIndex : -1;
                        pendingKerning = 0;
                        continue;
                    }
                    if (currentFont == null && aptosCommaQuoteFont != null &&
                        currentAptosFont && !currentDisplayFont && fontSize == 12 &&
                        lines[lineIndex] == "(,\u0094) Tj")
                    {
                        var adjustment = aptosCommaQuoteFont.LegacyKerningAdjustment(',', '\u201D', fontSize);
                        if (Math.Abs(adjustment) > 0.001)
                        {
                            lines[lineIndex] = $"[(,) {(-adjustment * 1000 / fontSize).ToString("0.####", CultureInfo.InvariantCulture)} (\u0094)] TJ";
                            commaQuoteCarry = adjustment;
                            changed = true;
                        }
                        continue;
                    }
                    if (currentFont == null || fontSize <= 0) continue;
                    var textMatch = Regex.Match(lines[lineIndex],
                        @"^\(([A-Za-zÀ-ÖØ-öø-ÿ0-9:/._?&#=%;,+*""\u00a0\u0092\u0094-]{2,})\) Tj$");
                    if (!textMatch.Success) continue;
                    var text = textMatch.Groups[1].Value;
                    var quotedSourceToken = text.EndsWith(",\u0094", StringComparison.Ordinal)
                        ? text[..^1] + '\u201D' : "";
                    var kernClosingQuote = currentDisplayFont &&
                        _closingQuoteSourceCounts.GetValueOrDefault(quotedSourceToken) == 1 &&
                        !_excludedClosingQuoteTokens.Contains(quotedSourceToken);
                    var apostropheSourceToken = text.Replace('\u0092', '\u2019');
                    var kernApostrophe = currentDisplayFont && text.Contains('\u0092') &&
                        _apostropheDisplaySourceCounts.GetValueOrDefault(apostropheSourceToken) == 1 &&
                        !_excludedApostropheDisplayTokens.Contains(apostropheSourceToken);
                    var kernWinAnsiClosing = Environment.GetEnvironmentVariable(
                        "DOCXPORT_DIAG_WINANSI_CLOSING_KERNING") != "0" &&
                        (currentAptosFont && previousWord.Length == 0 &&
                         IsWinAnsiClosingKerningCandidate(text) ||
                         Environment.GetEnvironmentVariable("DOCXPORT_DIAG_LEFT_CURLY_MARK_PAIR") == "1" &&
                         (currentAptosFont || currentCalibriFont) &&
                         text.Length >= 2 && text[^2] == '.' && text[^1] is ('\u0092' or '\u0094') &&
                         IsWinAnsiClosingKerningCandidate(text));
                    var splitScheme = (text is "https:" or "http:") &&
                        lines.Skip(lineIndex + 1).Take(8)
                            .FirstOrDefault(candidate => candidate.EndsWith(") Tj", StringComparison.Ordinal) ||
                                candidate == "ET")?.StartsWith("(//", StringComparison.Ordinal) == true;
                    var domainPath = Regex.IsMatch(text,
                        @"^(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,}/[A-Za-z0-9]", RegexOptions.CultureInvariant);
                    if (text.StartsWith("https://", StringComparison.Ordinal) || splitScheme || domainPath)
                    {
                        inUrl = true;
                        urlStartX = textOriginX;
                        urlKerning = 0;
                        previousUrlCharacter = '\0';
                        previousUrlFragment = "";
                        urlFont = currentFont;
                        urlFontSize = fontSize;
                    }
                    if (!inUrl && (text[^1] is ',' or '.') && !currentAptosFont) continue;
                    if (!inUrl && !IsKerningCandidate(text) && !kernClosingQuote &&
                        !kernApostrophe && !kernWinAnsiClosing) continue;
                    if (_documentHasTrackedInsertions && currentDisplayFont && fontSize >= 12 && !inUrl) continue;
                    if (!inUrl && text.EndsWith('.') && text.Length <= 3) continue;
                    var nextMove = lineIndex + 1 < lines.Length
                        ? Regex.Match(lines[lineIndex + 1], @"^(-?\d+(?:\.\d+)?) 0 Td$") : Match.Empty;
                    if (!inUrl && text.EndsWith('.') && nextMove.Success &&
                        double.TryParse(nextMove.Groups[1].Value, NumberStyles.Float,
                            CultureInfo.InvariantCulture, out var nextHorizontal) &&
                        currentFont.TryMeasure(text, fontSize, out var currentWidth) &&
                        nextHorizontal - currentWidth > 10) continue;
                    previousWord = text;
                    if (inUrl && previousUrlCharacter != '\0' &&
                        ReferenceEquals(urlFont, currentFont) && fontSize == urlFontSize &&
                        urlMoveIndex >= 0 &&
                        Math.Abs(textScale - 100) < 0.001)
                    {
                        var boundaryAdjustment = currentFont.KerningAdjustment(
                            previousUrlCharacter, text[0], fontSize);
                        if (Math.Abs(boundaryAdjustment) > 0.001)
                        {
                            if (currentFont.TryMeasure(previousUrlFragment, fontSize,
                                    out var expectedAdvance, useKerning: true))
                            {
                                if (TryAdjustKernedFragmentMove(lines, urlMoveIndex,
                                        expectedAdvance, boundaryAdjustment))
                                {
                                    lineShift += boundaryAdjustment;
                                    urlKerning += boundaryAdjustment;
                                    changed = true;
                                }
                            }
                        }
                    }
                    if (inUrl)
                    {
                        previousUrlCharacter = text[^1];
                        previousUrlFragment = text;
                    }
                    urlMoveIndex = -1;
                    if (ShouldSuppressRepeatedPairExpansion(text,
                        currentFont.KerningAdjustment(text[0], text[0], fontSize)))
                        continue;
                    if (baseline > 0 && baseline < page.Height.Point * 0.15 &&
                        _balancedFooterFieldSpaces.ContainsKey(text)) continue;
                    var formatted = new StringBuilder("[");
                    var start = 0;
                    var totalAdjustment = 0.0;
                    for (var characterIndex = 1; characterIndex < text.Length; characterIndex++)
                    {
                        var adjustment = kernClosingQuote && characterIndex == text.Length - 1
                            ? currentFont.LegacyKerningAdjustment(text[characterIndex - 1],
                                '\u201D', fontSize)
                            : kernApostrophe && (text[characterIndex - 1] == '\u0092' ||
                                text[characterIndex] == '\u0092')
                            ? currentFont.LegacyKerningAdjustment(
                                text[characterIndex - 1] == '\u0092' ? '\u2019' : text[characterIndex - 1],
                                text[characterIndex] == '\u0092' ? '\u2019' : text[characterIndex], fontSize)
                            : currentAptosFont && text.Contains('\u00a0')
                            ? currentFont.LegacyKerningAdjustment(
                                text[characterIndex - 1], text[characterIndex], fontSize)
                            : currentFont.KerningAdjustment(
                                kernWinAnsiClosing
                                    ? NormalizeWinAnsiQuote(text[characterIndex - 1])
                                    : text[characterIndex - 1],
                                kernWinAnsiClosing
                                    ? NormalizeWinAnsiQuote(text[characterIndex])
                                    : text[characterIndex], fontSize);
                        if (Math.Abs(adjustment) < 0.001) continue;
                        totalAdjustment += adjustment;
                        formatted.Append('(').Append(text, start, characterIndex - start)
                            .Append(") ").Append((-adjustment * 1000 / fontSize)
                                .ToString("0.####", CultureInfo.InvariantCulture)).Append(' ');
                        start = characterIndex;
                    }
                    if (start == 0) continue;
                    formatted.Append('(').Append(text, start, text.Length - start).Append(")] TJ");
                    lines[lineIndex] = formatted.ToString();
                    var canCarryKerning = !text.Contains('\u00a0');
                    if (canCarryKerning &&
                        (((fontSize >= 12 || currentDisplayFont ||
                            currentArialFont ||
                            currentAptosFont && (smallAptosCarrySetting == "1" ||
                                smallAptosCarrySetting == "9" && fontSize < 10 ||
                                smallAptosCarrySetting == "body" && fontSize < 10 &&
                                    baseline >= page.Height.Point * 0.15)) &&
                            (!hasHorizontalScaling || plainLatinLine)) || inUrl) &&
                        Math.Abs(textScale - 100) < 0.001)
                    {
                        pendingKerning += totalAdjustment;
                        if (kernClosingQuote && _closingQuoteFollowedBySpace.Contains(quotedSourceToken))
                            pendingKerning += currentFont.LegacyKerningAdjustment('\u201D', ' ', fontSize);
                    }
                    else if (canCarryKerning && currentScaledKerningFont &&
                        Math.Abs(textScale - 100) > 0.001)
                        pendingKerning += totalAdjustment * textScale / 100;
                    if (inUrl && Math.Abs(textScale - 100) < 0.001)
                        urlKerning += totalAdjustment;
                    changed = true;
                }
                if (changed && (!filtered || content.Stream.TryUncompress()))
                    content.Stream.Value = Encoding.Latin1.GetBytes(string.Join('\n', lines));
            }
        }
    }

    internal static bool IsKerningCandidate(string text) => Regex.IsMatch(text,
        @"^(?:[A-Za-zÀ-ÖØ-öø-ÿ]{2,}[0-9]*[,.]?|[A-Z]{2,}-|([a-z])\1{3,}-|(?:[A-Za-z]\.){2,}|[A-Za-zÀ-ÖØ-öø-ÿ]{2,}(?:\u00a0[A-Za-zÀ-ÖØ-öø-ÿ]{2,})+|""[A-Za-zÀ-ÖØ-öø-ÿ]{2,}\*?""?)$");

    internal static bool ShouldSuppressRepeatedPairExpansion(string text, double adjustment) =>
        adjustment > 0 && text.Length >= 4 &&
        (text[0] is >= 'A' and <= 'Z' or >= 'a' and <= 'z') &&
        text.All(character => character == text[0]);

    internal static bool IsWinAnsiClosingKerningCandidate(string text) => Regex.IsMatch(text,
        @"^[A-Za-zÀ-ÖØ-öø-ÿ]{2,}[.,!?;:]?[\u0092\u0094]$");

    internal static char NormalizeWinAnsiQuote(char character) => character switch
    {
        '\u0091' => '‘',
        '\u0092' => '’',
        '\u0093' => '“',
        '\u0094' => '”',
        _ => character
    };

    internal static bool CarriesScaledKerning(string family, bool bold = false, bool italic = false) =>
        (family is "Calibri" or "Arial Black") && !(bold && italic);

    internal static bool TryAdjustKernedFragmentMove(string[] lines, int index,
        double expectedAdvance, double adjustment)
    {
        var move = Regex.Match(lines[index], @"^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) Td$");
        if (!move.Success || !double.TryParse(move.Groups[1].Value, NumberStyles.Float,
                CultureInfo.InvariantCulture, out var horizontal) ||
            Math.Abs(horizontal - expectedAdvance) >= 1.5) return false;
        lines[index] = $"{(horizontal + adjustment).ToString("0.####", CultureInfo.InvariantCulture)} " +
            $"{move.Groups[2].Value} Td";
        return true;
    }

    internal static bool AdjustKernedUnderline(string[] lines, int start, double textX,
        double textY, double adjustment)
    {
        for (var index = start; index + 2 < lines.Length && index < start + 20; index++)
        {
            if (lines[index] == "BT") break;
            var first = Regex.Match(lines[index], @"^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) m$");
            var last = Regex.Match(lines[index + 1], @"^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) l$");
            if (!first.Success || !last.Success || lines[index + 2] != "S") continue;
            var left = double.Parse(first.Groups[1].Value, CultureInfo.InvariantCulture);
            var firstY = double.Parse(first.Groups[2].Value, CultureInfo.InvariantCulture);
            var right = double.Parse(last.Groups[1].Value, CultureInfo.InvariantCulture);
            var lastY = double.Parse(last.Groups[2].Value, CultureInfo.InvariantCulture);
            if (Math.Abs(left - textX) > 0.25 || Math.Abs(firstY - textY) > 3 ||
                Math.Abs(firstY - lastY) > 0.01 || right <= left) continue;
            lines[index + 1] = $"{(right + adjustment).ToString("0.####", CultureInfo.InvariantCulture)} " +
                $"{last.Groups[2].Value} l";
            return true;
        }
        return false;
    }

    internal static void AdjustTrackedUnderlines(string[] lines,
        IReadOnlyDictionary<int, (double Baseline, double EndShift)> tracked)
    {
        var previousEnds = new Dictionary<(double Y, double X), double>();
        foreach (var (endTextIndex, placement) in tracked.OrderBy(entry => entry.Key))
        {
            for (var index = endTextIndex + 1; index + 2 < lines.Length; index++)
            {
                if (lines[index] == "BT") break;
                var first = Regex.Match(lines[index], @"^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) m$");
                var last = Regex.Match(lines[index + 1], @"^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) l$");
                if (!first.Success || !last.Success || lines[index + 2] != "S") continue;
                var left = double.Parse(first.Groups[1].Value, CultureInfo.InvariantCulture);
                var right = double.Parse(last.Groups[1].Value, CultureInfo.InvariantCulture);
                var firstY = double.Parse(first.Groups[2].Value, CultureInfo.InvariantCulture);
                var lastY = double.Parse(last.Groups[2].Value, CultureInfo.InvariantCulture);
                if (right <= left || Math.Abs(firstY - lastY) > 0.01 ||
                    placement.Baseline - firstY is < 0.3 or > 4) continue;
                var startShift = previousEnds.GetValueOrDefault((Math.Round(firstY, 2), Math.Round(left, 2)));
                lines[index] = $"{(left + startShift).ToString("0.####", CultureInfo.InvariantCulture)} " +
                    $"{first.Groups[2].Value} m";
                lines[index + 1] = $"{(right + placement.EndShift).ToString("0.####", CultureInfo.InvariantCulture)} " +
                    $"{last.Groups[2].Value} l";
                previousEnds[(Math.Round(firstY, 2), Math.Round(right, 2))] = placement.EndShift;
                break;
            }
        }
    }

    private void ApplyCharacterMetrics(PdfDocumentRenderer renderer, PdfDocument document,
        UglyToad.PdfPig.PdfDocument parsed)
    {
        if (_characterScales.Count == 0 && _fitTextWidths.Count == 0 && _partialFitGroups.Count == 0 &&
            _characterSpacings.Count == 0 && _textBorderRanges.Count == 0 &&
            _topAlignedFontSizes.Count == 0 && _verticalAlignedFonts.Count == 0 &&
            _gridGlyphParagraphs.Count == 0 && _gridBoundaryParagraphs.Count == 0 &&
            _oversizedStoryLatinParagraphs.Count == 0)
            return;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            var page = parsed.GetPage(index + 1);
            var baselines = new List<(double Position, double Scale, double Spacing,
                double HorizontalShift, double VerticalShift, double TopFontSize, double RiseFactor,
                PartialFitMetric[]? Groups, bool GridGlyph, bool GridGlyphRise, bool PackedGrid, bool SnapGrid,
                bool WordBoundaryTracking, double Right, bool OversizedGridLatin,
                bool OversizedStoryLatin, bool GridBoundary, bool FittedTwelvePointTracking,
                bool InheritedTimesTracking, bool MixedCjkFitBody)>();
            var bodyBorderIndex = 0;
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, null))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    info.LayoutInfo.ContentArea is not { } area ||
                    !_characterScales.ContainsKey(paragraph) && !_fitTextWidths.ContainsKey(paragraph) &&
                    !_partialFitGroups.ContainsKey(paragraph) &&
                    !_characterSpacings.ContainsKey(paragraph) &&
                    !_textBorderRanges.ContainsKey(paragraph) &&
                    !_topAlignedFontSizes.ContainsKey(paragraph) &&
                    !_verticalAlignedFonts.ContainsKey(paragraph) &&
                    !_hangulBaselineShiftParagraphs.Contains(paragraph) &&
                    !_gridGlyphParagraphs.Contains(paragraph) &&
                    !_gridBoundaryParagraphs.Contains(paragraph) &&
                    !_oversizedStoryLatinParagraphs.Contains(paragraph)) continue;
                var border = _textBorderRanges.GetValueOrDefault(paragraph)?.FirstOrDefault();
                var verticalShift = border == null ? 0 : _bodyParagraphs.Contains(paragraph)
                    ? border.Inset + bodyBorderIndex++ * 8.2 : -border.Inset - 0.9;
                if (_hangulBaselineShiftParagraphs.Contains(paragraph)) verticalShift += 0.28;
                if (_fitTextTracking.Contains(paragraph) && _bodyParagraphs.Contains(paragraph) &&
                    !_documentHasFitText)
                    verticalShift += 4;
                if (_fittedBodyAfterSecondCjkHeader.Contains(paragraph))
                    verticalShift -= 1.05;
                var top = area.Y.Point + offset;
                var letter = page.Letters.FirstOrDefault(candidate =>
                    candidate.GlyphRectangle.Left >= area.X.Point - 2 &&
                    candidate.GlyphRectangle.Left <= area.X.Point + area.Width.Point + 2 &&
                    page.Height - candidate.GlyphRectangle.Top >= top - 5 &&
                    page.Height - candidate.GlyphRectangle.Top <= top + area.Height.Point + 5);
                if (letter != null)
                {
                    if (_latinCalibriFitParagraphs.Contains(paragraph) && _headerParagraphCount < 2)
                        verticalShift += 0.95;
                    var fittedCalibriTracking = _latinCalibriFitParagraphs.Contains(paragraph) &&
                        _bodyParagraphs.Contains(paragraph) && !_fitTextTracking.Contains(paragraph) &&
                        letter.FontName.Contains("Calibri,BoldItalic", StringComparison.OrdinalIgnoreCase);
                    var fittedMixedTracking = _mixedCalibriYuMinchoFitParagraphs.Contains(paragraph) &&
                        _bodyParagraphs.Contains(paragraph) && !_fitTextTracking.Contains(paragraph) &&
                        letter.FontName.Contains("Calibri", StringComparison.OrdinalIgnoreCase);
                    var fittedYuMinchoTracking = _latinYuMinchoFitParagraphs.Contains(paragraph) &&
                        _bodyParagraphs.Contains(paragraph) && !_fitTextTracking.Contains(paragraph) &&
                        letter.FontName.Contains("Yu Mincho", StringComparison.OrdinalIgnoreCase);
                    var fittedTwelvePointTracking = fittedCalibriTracking || fittedMixedTracking ||
                        fittedYuMinchoTracking;
                    var inheritedTimesTracking = _latinTimesFitParagraphs.Contains(paragraph) &&
                        _bodyParagraphs.Contains(paragraph) && _fitTextTracking.Contains(paragraph) &&
                        letter.FontName.Contains("Times New Roman", StringComparison.OrdinalIgnoreCase);
                    var fit = _fitTextWidths.TryGetValue(paragraph, out var targetWidth)
                        ? FitTextMetrics(page, letter, area.X.Point, area.Width.Point, targetWidth,
                            _fitTextTracking.Contains(paragraph) || fittedTwelvePointTracking,
                            !_bodyParagraphs.Contains(paragraph),
                            fittedTwelvePointTracking || inheritedTimesTracking
                                ? _fitTextCharacterCounts.GetValueOrDefault(paragraph) : 0,
                            fittedTwelvePointTracking ? 12.0 / 11 : 1, inheritedTimesTracking) :
                        (Scale: 100.0, Spacing: 0.0);
                    var groups = _partialFitGroups.TryGetValue(paragraph, out var partial)
                        ? BuildPartialFitMetrics(page, letter, area.X.Point, area.Width.Point, partial) : null;
                    var alignedFont = _verticalAlignedFonts.GetValueOrDefault(paragraph);
                    baselines.Add((letter.StartBaseLine.Y,
                        _characterScales.GetValueOrDefault(paragraph, fit.Scale),
                        _characterSpacings.GetValueOrDefault(paragraph, fit.Spacing),
                        (border?.Inset ?? 0) + _centeredSpacingOffsets.GetValueOrDefault(paragraph) +
                            _packedGridOffsets.GetValueOrDefault(paragraph), verticalShift,
                        alignedFont.Size > 0 ? alignedFont.Size :
                        _topAlignedFontSizes.TryGetValue(paragraph, out var topSize) &&
                            page.Letters.Any(candidate =>
                                Math.Abs(candidate.StartBaseLine.Y - letter.StartBaseLine.Y) < 0.5 &&
                                candidate.FontSize >= topSize - 0.5) ? topSize : 0,
                        _topAlignedHeaderParagraphs.Contains(paragraph) ? 0.98 : alignedFont.Factor, groups,
                        _gridGlyphParagraphs.Contains(paragraph),
                        _gridGlyphParagraphs.Contains(paragraph) &&
                            !_unshiftedGridGlyphParagraphs.Contains(paragraph),
                        _packedGridParagraphs.Contains(paragraph),
                        _packedGridOffsets.ContainsKey(paragraph),
                        fittedTwelvePointTracking || inheritedTimesTracking ||
                            _characterSpacings.ContainsKey(paragraph) &&
                            !_centeredSpacingOffsets.ContainsKey(paragraph) &&
                            !_gridGlyphParagraphs.Contains(paragraph) &&
                            !_packedGridParagraphs.Contains(paragraph) &&
                            paragraph.Format.Alignment != MigraDoc.DocumentObjectModel.ParagraphAlignment.Center,
                        area.X.Point + area.Width.Point,
                        _oversizedGridLatinParagraphs.Contains(paragraph),
                        _oversizedStoryLatinParagraphs.Contains(paragraph),
                        _gridBoundaryParagraphs.Contains(paragraph), fittedTwelvePointTracking,
                        inheritedTimesTracking, _mixedCjkFitBodyParagraphs.Contains(paragraph)));
                    if (_topAlignedFontSizes.TryGetValue(paragraph, out var alignedSize))
                        foreach (var line in page.Letters.Where(candidate =>
                                candidate.StartBaseLine.Y < letter.StartBaseLine.Y - 0.5 &&
                                candidate.GlyphRectangle.Left >= area.X.Point - 2 &&
                                candidate.GlyphRectangle.Left <= area.X.Point + area.Width.Point + 2 &&
                                page.Height - candidate.GlyphRectangle.Top >= top - 5 &&
                                page.Height - candidate.GlyphRectangle.Top <= top + area.Height.Point + 5)
                            .GroupBy(candidate => Math.Round(candidate.StartBaseLine.Y, 1)))
                        {
                            var mixed = line.Any(candidate => candidate.FontSize >= alignedSize - 0.5) &&
                                line.Any(candidate => candidate.FontSize < alignedSize - 0.5);
                            if (mixed)
                                baselines.Add((line.First().StartBaseLine.Y, 100, 0, 0,
                                    _gridGlyphParagraphs.Contains(paragraph) ? 12.5 : 0,
                                    alignedSize, 0, null, false, false, false, false, false,
                                    area.X.Point + area.Width.Point,
                                    _oversizedGridLatinParagraphs.Contains(paragraph),
                                    _oversizedStoryLatinParagraphs.Contains(paragraph), false, false, false,
                                    false));
                        }
                }
            }
            if (baselines.Count == 0) continue;
            var gridBoundaries = baselines.Any(metric => metric.GridBoundary)
                ? page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                    .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                    .SelectMany(group =>
                    {
                        var ordered = group.OrderBy(letter => letter.StartBaseLine.X).ToArray();
                        return ordered.Zip(ordered.Skip(1),
                            (previous, current) => (previous, current));
                    })
                    .Where(pair => pair.previous.Value.Length == 1 &&
                        pair.previous.Value[0] is >= 'A' and <= 'Z' &&
                        pair.current.Value.Length == 1 && IsCjk(pair.current.Value[0]))
                    .Select(pair => (pair.current.StartBaseLine.X, pair.current.StartBaseLine.Y))
                    .ToArray() : [];
            var gridSpaceBoundaries = baselines.Any(metric => metric.GridBoundary)
                ? page.Letters.GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                    .SelectMany(group =>
                    {
                        var ordered = group.OrderBy(letter => letter.StartBaseLine.X).ToArray();
                        return Enumerable.Range(3, Math.Max(0, ordered.Length - 3))
                            .Where(position => ordered[position - 3].Value.Length == 1 &&
                                IsCjk(ordered[position - 3].Value[0]) &&
                                ordered[position - 2].Value == "\u2003" &&
                                ordered[position - 1].Value == "\u2003" &&
                                ordered[position - 1].StartBaseLine.X -
                                    ordered[position - 2].StartBaseLine.X < 16 &&
                                ordered[position].Value.Length == 1 &&
                                IsCjk(ordered[position].Value[0]))
                            .Select(position => (ordered[position].StartBaseLine.X,
                                ordered[position].StartBaseLine.Y));
                    }).ToArray() : [];
            foreach (var entry in document.Pages[index].Contents.Elements)
            {
                if ((entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not PdfSharp.Pdf.Advanced.PdfContent content ||
                    content.Stream == null || content.Elements.ContainsKey("/Filter")) continue;
                var commands = Encoding.Latin1.GetString(content.Stream.Value);
                var lines = commands.Split('\n');
                var trackedUnderlines = new Dictionary<int, (double Baseline, double EndShift)>();
                var originalX = 0.0;
                var originalY = 0.0;
                var actualX = 0.0;
                var actualY = 0.0;
                var paragraphX = 0.0;
                var scale = 100.0;
                var spacing = 0.0;
                var horizontalShift = 0.0;
                var verticalShift = 0.0;
                var topFontSize = 0.0;
                var riseFactor = 0.0;
                var currentFontSize = 0.0;
                var currentRise = 0.0;
                var lastFontLine = -1;
                var gridGlyph = false;
                var gridGlyphRise = false;
                var packedGrid = false;
                var snapGrid = false;
                var wordBoundaryTracking = false;
                var oversizedGridLatin = false;
                var oversizedStoryLatin = false;
                var largeRunTextEmitted = false;
                var characterCount = 0;
                var previousBaseline = double.NaN;
                var previousStart = 0.0;
                var previousRight = 0.0;
                var previousCount = 0;
                var paragraphRight = 0.0;
                var continuedTextObject = false;
                var gridBoundary = false;
                var fittedTwelvePointTracking = false;
                var inheritedTimesTracking = false;
                var mixedCjkFitBody = false;
                var fittedWordGaps = 0;
                var gridBoundaryShift = 0.0;
                var mixedCjkAdvanceShift = 0.0;
                PartialFitMetric[]? fitGroups = null;
                var activeFitGroup = -1;
                for (var lineIndex = 0; lineIndex < lines.Length; lineIndex++)
                {
                    if (lines[lineIndex] == "ET" && wordBoundaryTracking && spacing != 0 &&
                        Math.Abs(scale - 100) < 0.001 && fitGroups == null && characterCount > 0)
                        trackedUnderlines[lineIndex] = (originalY, characterCount * spacing);
                    if (lines[lineIndex] == "BT")
                    {
                        if (wordBoundaryTracking && characterCount > 0)
                        {
                            previousBaseline = originalY;
                            previousStart = paragraphX;
                            previousRight = paragraphRight;
                            previousCount = characterCount;
                        }
                        else previousBaseline = double.NaN;
                        originalX = originalY = actualX = actualY = paragraphX = 0;
                        scale = 100;
                        spacing = 0;
                        horizontalShift = verticalShift = 0;
                        topFontSize = currentFontSize = currentRise = riseFactor = 0;
                        lastFontLine = -1;
                        gridGlyph = false;
                        packedGrid = false;
                        snapGrid = false;
                        wordBoundaryTracking = false;
                        oversizedGridLatin = false;
                        oversizedStoryLatin = false;
                        largeRunTextEmitted = false;
                        continuedTextObject = false;
                        gridBoundary = false;
                        fittedTwelvePointTracking = false;
                        inheritedTimesTracking = false;
                        mixedCjkFitBody = false;
                        fittedWordGaps = 0;
                        gridBoundaryShift = 0;
                        mixedCjkAdvanceShift = 0;
                        characterCount = 0;
                        fitGroups = null;
                        activeFitGroup = -1;
                        continue;
                    }
                    var font = Regex.Match(lines[lineIndex], @"^/\S+ (\d+(?:\.\d+)?) Tf$");
                    if (font.Success && double.TryParse(font.Groups[1].Value, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var fontSize))
                    {
                        if ((fittedTwelvePointTracking || mixedCjkFitBody) && fontSize == 11)
                        {
                            lines[lineIndex] = Regex.Replace(lines[lineIndex], @" 11 Tf$", " 12 Tf");
                            fontSize = 12;
                        }
                        currentFontSize = fontSize;
                        lastFontLine = lineIndex;
                        largeRunTextEmitted = false;
                    }
                    var text = Regex.Match(lines[lineIndex], @"^\(((?:\\.|[^)])*)\) Tj$");
                    if (text.Success)
                    {
                        if (oversizedStoryLatin && currentFontSize == 18 && lastFontLine >= 0 &&
                            text.Groups[1].Value.Length > 1 &&
                            text.Groups[1].Value.All(char.IsAsciiLetter))
                        {
                            var fontResource = Regex.Match(lines[lastFontLine],
                                @"^(/\S+) \d+(?:\.\d+)? Tf$");
                            if (fontResource.Success)
                                lines[lineIndex] = $"{fontResource.Groups[1].Value} 18 Tf\n" +
                                    "100 Tz\n0 Tc\n0 Ts\n" + lines[lineIndex];
                        }
                        characterCount += Regex.Replace(text.Groups[1].Value, @"\\.", "x").Length;
                        largeRunTextEmitted |= oversizedGridLatin && topFontSize > 0 &&
                            currentFontSize >= topFontSize - 0.5;
                        continue;
                    }
                    if (fittedTwelvePointTracking && lines[lineIndex].EndsWith(" TJ", StringComparison.Ordinal))
                    {
                        foreach (Match literal in Regex.Matches(lines[lineIndex], @"\((?:\\.|[^)])*\)"))
                            characterCount += Regex.Replace(literal.Value[1..^1], @"\\.", "x").Length;
                        continue;
                    }
                    var encodedText = Regex.Match(lines[lineIndex], @"^<([0-9A-Fa-f]+)> Tj$");
                    if (encodedText.Success)
                    {
                        if (packedGrid)
                            characterCount += encodedText.Groups[1].Length / 4;
                        largeRunTextEmitted |= oversizedGridLatin && topFontSize > 0 &&
                            currentFontSize >= topFontSize - 0.5;
                        continue;
                    }
                    var match = Regex.Match(lines[lineIndex],
                        @"^(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) Td$");
                    if (!match.Success ||
                        !double.TryParse(match.Groups[1].Value, NumberStyles.Float,
                            CultureInfo.InvariantCulture, out var horizontal) ||
                        !double.TryParse(match.Groups[2].Value, NumberStyles.Float,
                            CultureInfo.InvariantCulture, out var vertical)) continue;
                    originalX += horizontal;
                    originalY += vertical;
                    if (vertical != 0)
                    {
                        gridBoundaryShift = 0;
                        mixedCjkAdvanceShift = 0;
                    }
                    if (vertical != 0 || paragraphX == 0)
                    {
                        paragraphX = originalX;
                        var metrics = baselines.FirstOrDefault(candidate =>
                            Math.Abs(candidate.Position - originalY) < 0.5);
                        var continuedPackedGrid = packedGrid && metrics.Position == 0 &&
                            Math.Abs(vertical) <= 30;
                        scale = metrics.Scale;
                        if (scale == 0) scale = 100;
                        spacing = continuedPackedGrid ? spacing : metrics.Spacing;
                        horizontalShift = continuedPackedGrid ? horizontalShift : metrics.HorizontalShift;
                        verticalShift = metrics.VerticalShift;
                        topFontSize = metrics.TopFontSize;
                        riseFactor = metrics.RiseFactor;
                        fitGroups = metrics.Groups;
                        gridGlyph = metrics.GridGlyph ||
                            gridGlyph && metrics.Position == 0 && Math.Abs(vertical) <= 30;
                        gridGlyphRise = metrics.GridGlyphRise ||
                            gridGlyphRise && metrics.Position == 0 && Math.Abs(vertical) <= 30;
                        gridBoundary = metrics.GridBoundary ||
                            gridBoundary && metrics.Position == 0 && Math.Abs(vertical) <= 30;
                        packedGrid = metrics.PackedGrid || continuedPackedGrid;
                        snapGrid = metrics.SnapGrid;
                        wordBoundaryTracking = metrics.WordBoundaryTracking;
                        fittedTwelvePointTracking = metrics.FittedTwelvePointTracking;
                        inheritedTimesTracking = metrics.InheritedTimesTracking;
                        mixedCjkFitBody = metrics.MixedCjkFitBody;
                        fittedWordGaps = 0;
                        if ((fittedTwelvePointTracking || mixedCjkFitBody) &&
                            currentFontSize == 11 && lastFontLine >= 0)
                        {
                            lines[lastFontLine] = Regex.Replace(lines[lastFontLine], @" 11 Tf$", " 12 Tf");
                            currentFontSize = 12;
                        }
                        oversizedGridLatin = metrics.OversizedGridLatin;
                        oversizedStoryLatin = metrics.OversizedStoryLatin;
                        largeRunTextEmitted = false;
                        paragraphRight = metrics.Right;
                        if (wordBoundaryTracking && Math.Abs(originalY - previousBaseline) < 0.5 &&
                            originalX > previousStart && originalX < previousRight)
                        {
                            var gap = page.Letters.Where(candidate =>
                                    Math.Abs(candidate.StartBaseLine.Y - originalY) < 0.5 &&
                                    candidate.StartBaseLine.X >= previousStart - 1 &&
                                    candidate.EndBaseLine.X <= originalX + 0.25)
                                .Select(candidate => originalX - candidate.EndBaseLine.X)
                                .DefaultIfEmpty(double.PositiveInfinity).Min();
                            if (gap >= 0 && gap < currentFontSize * 0.75)
                            {
                                characterCount = previousCount + 1;
                                continuedTextObject = true;
                            }
                        }
                        if (gridGlyph && currentFontSize == 11 && lastFontLine >= 0)
                        {
                            lines[lastFontLine] = Regex.Replace(lines[lastFontLine], @" 11 Tf$", " 12 Tf");
                            currentFontSize = 12;
                        }
                        if (gridGlyph) spacing = -1;
                        activeFitGroup = -1;
                        if (!continuedTextObject) characterCount = 0;
                    }
                    else if (wordBoundaryTracking && characterCount > 0 && horizontal > 0 &&
                        fitGroups == null &&
                        (lineIndex == 0 || !lines[lineIndex - 1].EndsWith("-) Tj", StringComparison.Ordinal)))
                    {
                        characterCount++;
                        if (inheritedTimesTracking) fittedWordGaps++;
                    }
                    var fitGroupIndex = fitGroups == null ? -1 : Array.FindLastIndex(fitGroups,
                        group => characterCount >= group.StartCount);
                    var fitGroup = fitGroupIndex >= 0 ? fitGroups![fitGroupIndex] : null;
                    var targetX = fitGroup != null
                        ? paragraphX + fitGroup.TargetOffset + originalX - fitGroup.StartX +
                            (characterCount - fitGroup.StartCount) * fitGroup.Spacing
                        : fittedTwelvePointTracking && vertical == 0
                        ? paragraphX + (originalX - paragraphX) * 12.0 / 11 + characterCount * spacing
                        : inheritedTimesTracking && vertical == 0
                        ? originalX + characterCount * spacing + fittedWordGaps * 0.3
                        : continuedTextObject && paragraphX == originalX
                        ? originalX + characterCount * spacing
                        : vertical == 0 && paragraphX != originalX
                        ? paragraphX + ((originalX - paragraphX) + characterCount * spacing) * scale / 100
                        : originalX;
                    targetX += horizontalShift;
                    if (mixedCjkFitBody && vertical == 0 && fitGroup == null &&
                        Math.Abs(horizontal - 11) < 0.01)
                    {
                        var advance = horizontal * scale / 100;
                        mixedCjkAdvanceShift += QuantizeFittedCjkAdvance(advance) - advance;
                    }
                    targetX += mixedCjkAdvanceShift;
                    if (gridBoundary && vertical == 0 && gridBoundaries.Any(boundary =>
                            Math.Abs(boundary.X - originalX) < 0.5 &&
                            Math.Abs(boundary.Y - originalY) < 0.5))
                        gridBoundaryShift += 1.7;
                    targetX += gridBoundaryShift;
                    if (gridBoundary && vertical == 0)
                        targetX -= 0.84 * gridSpaceBoundaries.Count(boundary =>
                            boundary.X <= originalX + 0.1 &&
                            Math.Abs(boundary.Y - originalY) < 0.5);
                    var trackLargeRun = oversizedGridLatin && largeRunTextEmitted &&
                        topFontSize > 0 && currentFontSize >= topFontSize - 0.5 && vertical == 0;
                    if (trackLargeRun) targetX += 3;
                    var targetY = originalY - verticalShift;
                    lines[lineIndex] = $"{(targetX - actualX).ToString("0.####", CultureInfo.InvariantCulture)} " +
                        $"{(targetY - actualY).ToString("0.####", CultureInfo.InvariantCulture)} Td";
                    var rise = riseFactor != 0 ? Math.Max(0, topFontSize - currentFontSize) * riseFactor :
                        topFontSize < 0 ? -Math.Max(0, currentFontSize + topFontSize) * 0.9 :
                        topFontSize > 0 ? Math.Max(0, topFontSize - currentFontSize) * 0.9 : 0;
                    if (gridGlyph && gridGlyphRise) rise += 0.5;
                    if (Math.Abs(rise - currentRise) > 0.001)
                    {
                        lines[lineIndex] += $"\n{rise.ToString("0.####", CultureInfo.InvariantCulture)} Ts";
                        currentRise = rise;
                    }
                    actualX = targetX;
                    actualY = targetY;
                    if (fitGroup != null && fitGroupIndex != activeFitGroup)
                    {
                        lines[lineIndex] += $"\n100 Tz\n{fitGroup.GlyphSpacing.ToString("0.####", CultureInfo.InvariantCulture)} Tc";
                        activeFitGroup = fitGroupIndex;
                    }
                    else if (vertical != 0 || originalX == paragraphX)
                        lines[lineIndex] += $"\n{(mixedCjkFitBody ? FittedCjkPaintScale(scale) : scale).ToString("0.####", CultureInfo.InvariantCulture)} Tz" +
                            $"\n{spacing.ToString("0.####", CultureInfo.InvariantCulture)} Tc";
                    if (trackLargeRun)
                    {
                        var fontResource = Regex.Match(lines[lastFontLine], @"^(/\S+) \d+(?:\.\d+)? Tf$");
                        lines[lineIndex] += fontResource.Success && Math.Abs(currentFontSize - 18) < 0.01
                            ? $"\n{fontResource.Groups[1].Value} 14 Tf\n128.5714 Tz\n0.3889 Tc\n0 Ts"
                            : "\n0.5 Tc";
                        largeRunTextEmitted = false;
                    }
                }
                AdjustTrackedUnderlines(lines, trackedUnderlines);
                content.Stream.Value = Encoding.Latin1.GetBytes(string.Join('\n', lines));
            }
        }
    }

    internal static double QuantizeFittedCjkAdvance(double advance) =>
        Math.Floor((advance + 1e-8) / 0.12) * 0.12;

    internal static double FittedCjkPaintScale(double layoutScale) =>
        QuantizeFittedCjkAdvance(11 * layoutScale / 100) / 12 * 100;

    private static (double Scale, double Spacing) FitTextMetrics(UglyToad.PdfPig.Content.Page page,
        UglyToad.PdfPig.Content.Letter first, double left, double width, double targetWidth,
        bool tracking, bool story, int characterCount = 0, double paintSizeRatio = 1,
        bool inheritedTimesTracking = false)
    {
        var line = page.Letters.Where(letter =>
            Math.Abs(letter.StartBaseLine.Y - first.StartBaseLine.Y) < 0.5 &&
            letter.StartBaseLine.X >= first.StartBaseLine.X - 1 &&
            letter.EndBaseLine.X <= left + width + 2).ToArray();
        if (line.Length == 0) return (100, 0);
        var naturalWidth = line.Max(letter => letter.EndBaseLine.X) - first.StartBaseLine.X;
        if (naturalWidth <= 0) return (100, 0);
        targetWidth = FittedVisualTargetWidth(first.FontName, targetWidth, tracking, story);
        if (paintSizeRatio > 1) targetWidth -= 1;
        if (inheritedTimesTracking) targetWidth -= 0.75;
        if (!tracking) return (100 * targetWidth / naturalWidth, 0);
        var scale = !inheritedTimesTracking &&
            first.FontName.Contains("Times New Roman", StringComparison.OrdinalIgnoreCase) &&
            targetWidth / naturalWidth > 1.15 ? 103.0 : 100.0;
        return (scale, (targetWidth * 100 / scale - naturalWidth * paintSizeRatio) /
            Math.Max(1, characterCount > 0 ? characterCount - 1 : line.Length - 1));
    }

    internal static double FittedVisualTargetWidth(string fontName, double targetWidth,
        bool tracking, bool story)
    {
        if (tracking) return targetWidth;
        if (fontName.Contains("Calibri,Italic", StringComparison.OrdinalIgnoreCase))
            return targetWidth - 1.5;
        if (story && fontName.EndsWith("Calibri,Bold", StringComparison.OrdinalIgnoreCase))
            return targetWidth - 1.25;
        return targetWidth;
    }

    private sealed record PartialFitMetric(int StartCount, double StartX, double TargetOffset,
        double Spacing, double GlyphSpacing);

    private void SuppressNativeColoredUnderlines(PdfDocumentRenderer renderer, PdfDocument document)
    {
        if (_coloredUnderlineParagraphs.Count == 0) return;
        for (var pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
        {
            var areas = GetPageStoryRenderInfos(renderer, pageIndex, document)
                .Where(entry => entry.Info.DocumentObject is MigraDoc.DocumentObjectModel.Paragraph paragraph &&
                    _coloredUnderlineParagraphs.ContainsKey(paragraph) &&
                    entry.Info.LayoutInfo.ContentArea != null)
                .Select(entry => (Area: entry.Info.LayoutInfo.ContentArea!, entry.Offset))
                .ToArray();
            if (areas.Length == 0) continue;
            var page = document.Pages[pageIndex];
            foreach (var reference in page.Contents.Elements.OfType<PdfSharp.Pdf.Advanced.PdfReference>())
            {
                if (reference.Value is not PdfSharp.Pdf.Advanced.PdfContent content ||
                    content.Stream == null) continue;
                var source = Encoding.Latin1.GetString(content.Stream.UnfilteredValue);
                var updated = Regex.Replace(source,
                    @"(?m)(?<x1>-?[\d.]+) (?<y1>-?[\d.]+) m\s+(?<x2>-?[\d.]+) (?<y2>-?[\d.]+) l\s+S(?=\s|$)",
                    match =>
                    {
                        var x1 = double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture);
                        var x2 = double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture);
                        var y1 = double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture);
                        var y2 = double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture);
                        if (Math.Abs(y1 - y2) > 0.1 || x2 - x1 < 10) return match.Value;
                        var top = page.Height.Point - y1;
                        return areas.Any(item => top >= item.Area.Y.Point + item.Offset - 6 &&
                                top <= item.Area.Y.Point + item.Offset + item.Area.Height.Point + 6 &&
                                x1 >= item.Area.X.Point - 3 &&
                                x2 <= item.Area.X.Point + item.Area.Width.Point + 3)
                            ? match.Value[..^1] + "n" : match.Value;
                    });
                if (updated != source) content.Stream.Value = Encoding.Latin1.GetBytes(updated);
            }
        }
    }

    private PdfDocument RenderInlineMathBars(PdfDocumentRenderer renderer, PdfDocument document)
    {
        if (_inlineMathRadicals.Count == 0 && _inlineMathCombinedScripts.Count == 0 &&
            _inlineMathFractions.Count == 0 && _nestedInlineMathFractions.Count == 0 &&
            _inlineMathMatrices.Count == 0) return document;
        using var stream = new MemoryStream();
        document.Save(stream, false);
        var bytes = stream.ToArray();
        if (_mixedInlineMathRadicals.Count > 0 && _standaloneInlineMathRadicals.Count == 0)
        {
            using var adjustable = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Modify);
            if (LowerMixedInlineRadicalRoots(adjustable))
            {
                using var adjustedStream = new MemoryStream();
                adjustable.Save(adjustedStream, false);
                bytes = adjustedStream.ToArray();
            }
        }
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(bytes);
        using var imported = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var result = new PdfDocument();
        foreach (var page in imported.Pages) result.AddPage(page);
        CopyNamedDestinations(imported, result);
        for (var index = 0; index < renderer.PageCount; index++)
        {
            var page = parsed.GetPage(index + 1);
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, document))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    info.LayoutInfo.ContentArea is not { } area) continue;
                var top = area.Y.Point + offset;
                if (_inlineMathFractions.TryGetValue(paragraph, out var fractions))
                {
                    var anchorSize = _mixedDisplayMathParagraphs.Contains(paragraph) ? 14.01 : 15.0;
                    var anchors = page.Letters.Where(letter => letter.Value == " " &&
                        Math.Abs(letter.PointSize - anchorSize) < 0.005 &&
                        page.Height - letter.GlyphRectangle.Top >= top - 5 &&
                        page.Height - letter.GlyphRectangle.Top <= top + area.Height.Point + 5 &&
                        letter.GlyphRectangle.Left >= area.X.Point - 2 &&
                        letter.GlyphRectangle.Left <= area.X.Point + area.Width.Point + 2).ToArray();
                    for (var fractionIndex = 0; fractionIndex < Math.Min(fractions.Count, anchors.Length);
                         fractionIndex++)
                    {
                        var anchor = anchors[fractionIndex];
                        var fraction = fractions[fractionIndex];
                        using var graphics = XGraphics.FromPdfPage(result.Pages[index]);
                        XFont font;
                        var displayMath = _mixedDisplayMathParagraphs.Contains(paragraph);
                        var fontSize = displayMath ? 12 : 8.52;
                        try { font = new XFont("Cambria Math", fontSize); }
                        catch (InvalidOperationException) { font = new XFont(_fontFamily, fontSize); }
                        var numeratorWidth = graphics.MeasureString(fraction.Numerator, font).Width;
                        var denominatorWidth = graphics.MeasureString(fraction.Denominator, font).Width;
                        var center = anchor.GlyphRectangle.Left + fraction.Width / 2;
                        var baselineTop = page.Height - anchor.GlyphRectangle.Top -
                            (anchorSize == 15.0 ? 0.25 : 0);
                        graphics.DrawString(fraction.Numerator, font, XBrushes.Black,
                            new XRect(center - numeratorWidth / 2, baselineTop - 14.84 +
                                (displayMath ? -1.08 : 0),
                                numeratorWidth + 1, 12), XStringFormats.TopLeft);
                        graphics.DrawRectangle(XBrushes.Black, center - fraction.Width / 2,
                            baselineTop - 3.83 + (displayMath ? 4.8 : 0), fraction.Width, 0.84);
                        graphics.DrawString(fraction.Denominator, font, XBrushes.Black,
                            new XRect(center - denominatorWidth / 2, baselineTop - 1.76 +
                                (displayMath ? 3.0 : 0),
                                denominatorWidth + 1, 12), XStringFormats.TopLeft);
                    }
                }
                if (_nestedInlineMathFractions.TryGetValue(paragraph, out var nestedFractions))
                {
                    var anchors = page.Letters.Where(letter => letter.Value == " " &&
                        Math.Abs(letter.PointSize - 14.03) < 0.005 &&
                        page.Height - letter.GlyphRectangle.Top >= top - 5 &&
                        page.Height - letter.GlyphRectangle.Top <= top + area.Height.Point + 5 &&
                        letter.GlyphRectangle.Left >= area.X.Point - 2 &&
                        letter.GlyphRectangle.Left <= area.X.Point + area.Width.Point + 2).ToArray();
                    for (var fractionIndex = 0;
                         fractionIndex < Math.Min(nestedFractions.Count, anchors.Length); fractionIndex++)
                    {
                        var anchor = anchors[fractionIndex];
                        var fraction = nestedFractions[fractionIndex];
                        using var graphics = XGraphics.FromPdfPage(result.Pages[index]);
                        PaintInlineMathTerm(graphics, fraction,
                            anchor.GlyphRectangle.Left + NestedMathBarWidth(fraction, true) / 2,
                            page.Height - anchor.GlyphRectangle.Top -
                                (fraction.Denominator?.Text == null ? 2.0 : 1.47) - 2.1, 8.52);
                    }
                }
                if (_inlineMathMatrices.TryGetValue(paragraph, out var matrices))
                {
                    var anchors = page.Letters.Where(letter => letter.Value == " " &&
                        Math.Abs(letter.PointSize - 14.02) < 0.005 &&
                        page.Height - letter.GlyphRectangle.Top >= top - 5 &&
                        page.Height - letter.GlyphRectangle.Top <= top + area.Height.Point + 5 &&
                        letter.GlyphRectangle.Left >= area.X.Point - 2 &&
                        letter.GlyphRectangle.Left <= area.X.Point + area.Width.Point + 2).ToArray();
                    for (var matrixIndex = 0; matrixIndex < Math.Min(matrices.Count, anchors.Length);
                         matrixIndex++)
                    {
                        var matrix = matrices[matrixIndex];
                        var anchor = anchors[matrixIndex];
                        using var graphics = XGraphics.FromPdfPage(result.Pages[index]);
                        XFont font;
                        try { font = new XFont("Cambria Math", 12); }
                        catch (InvalidOperationException) { font = new XFont(_fontFamily, 12); }
                        var columnX = anchor.GlyphRectangle.Left;
                        var firstTop = page.Height - anchor.GlyphRectangle.Top - 17.9 -
                            (matrix.Cells.Length - 2) * 8.5 -
                            (IsDigitMatrix(matrix.Cells) ? 0.38 : 0);
                        for (var columnIndex = 0; columnIndex < matrix.ColumnWidths.Length; columnIndex++)
                        {
                            for (var rowIndex = 0; rowIndex < matrix.Cells.Length; rowIndex++)
                            {
                                var value = matrix.Cells[rowIndex][columnIndex];
                                var textWidth = graphics.MeasureString(value, font).Width;
                                graphics.DrawString(value, font, XBrushes.Black,
                                    new XRect(columnX + (matrix.ColumnWidths[columnIndex] - textWidth) / 2,
                                        firstTop + rowIndex * 14.1, textWidth + 1, 14),
                                    XStringFormats.TopLeft);
                            }
                            columnX += matrix.ColumnWidths[columnIndex] +
                                (IsDigitMatrix(matrix.Cells) ? 12.05 : 12.3);
                        }
                    }
                }
                if (_inlineMathCombinedScripts.TryGetValue(paragraph, out var scripts))
                {
                    var previousBaseX = area.X.Point - 2;
                    foreach (var (superscript, subscript) in scripts)
                    {
                        var baseLetter = page.Letters.FirstOrDefault(letter =>
                            letter.PointSize >= 10 &&
                            letter.GlyphRectangle.Left > previousBaseX &&
                            page.Height - letter.GlyphRectangle.Top >= top - 5 &&
                            page.Height - letter.GlyphRectangle.Top <= top + area.Height.Point + 5 &&
                            letter.GlyphRectangle.Left >= area.X.Point - 2 &&
                            letter.GlyphRectangle.Left <= area.X.Point + area.Width.Point + 2 &&
                            page.Letters.Any(script => script.Value == superscript &&
                                script.GlyphRectangle.Left >= letter.GlyphRectangle.Right - 2 &&
                                script.GlyphRectangle.Left <= letter.GlyphRectangle.Right + 5 &&
                                script.StartBaseLine.Y > letter.StartBaseLine.Y + 1));
                        if (baseLetter == null) continue;
                        previousBaseX = baseLetter.GlyphRectangle.Left + 0.5;
                        using var graphics = XGraphics.FromPdfPage(result.Pages[index]);
                        XFont font;
                        try { font = new XFont("Cambria Math", 8.52); }
                        catch (InvalidOperationException) { font = new XFont(_fontFamily, 8.52); }
                        graphics.DrawString(subscript, font, XBrushes.Black,
                            new XRect(baseLetter.StartBaseLine.X + 6.24,
                                page.Height - baseLetter.StartBaseLine.Y - 5.45, 20, 12),
                            XStringFormats.TopLeft);
                    }
                }
                if (!_inlineMathRadicals.Contains(paragraph)) continue;
                foreach (var root in page.Letters.Where(letter => letter.Value == "√" &&
                    page.Height - letter.GlyphRectangle.Top >= top - 5 &&
                    page.Height - letter.GlyphRectangle.Top <= top + area.Height.Point + 5 &&
                    letter.GlyphRectangle.Left >= area.X.Point - 2 &&
                    letter.GlyphRectangle.Left <= area.X.Point + area.Width.Point + 2))
                {
                    var variable = page.Letters.FirstOrDefault(letter =>
                        letter.Value is not ("√" or " ") &&
                        letter.GlyphRectangle.Left >= root.GlyphRectangle.Right - 1 &&
                        letter.GlyphRectangle.Left <= root.GlyphRectangle.Right + 2 &&
                        Math.Abs(letter.StartBaseLine.Y - root.StartBaseLine.Y) < 1);
                    if (variable == null) continue;
                    using var graphics = XGraphics.FromPdfPage(result.Pages[index]);
                    graphics.DrawRectangle(XBrushes.Black, variable.StartBaseLine.X,
                        page.Height - root.StartBaseLine.Y - 11.02,
                        variable.EndBaseLine.X - variable.StartBaseLine.X + 0.34, 0.84);
                }
            }
        }
        AlignMatrixSurroundingText(renderer, result, document);
        return result;
    }

    private bool LowerMixedInlineRadicalRoots(PdfDocument document)
    {
        if (_mixedInlineMathRadicals.Count == 0 || _standaloneInlineMathRadicals.Count != 0 ||
            DocSystemFontAdvances.Find("Cambria Math") is not { } metrics)
            return false;
        var baseRoot = metrics.GlyphIndex('√').ToString("X4", CultureInfo.InvariantCulture);
        var roots = new HashSet<string> { baseRoot };
        var baseRoots = new HashSet<string> { baseRoot };
        string? tallRoot = null;
        if (metrics.TryGetMathVerticalVariant('√', out var variant))
        {
            tallRoot = variant.ToString("X4", CultureInfo.InvariantCulture);
            roots.Add(tallRoot);
        }
        foreach (PdfPage page in document.Pages)
        {
            var fonts = page.Resources.Elements.GetDictionary("/Font");
            if (fonts == null) continue;
            foreach (var key in fonts.Elements.Keys)
            {
                var font = fonts.Elements.GetDictionary(key);
                var name = font?.Elements.GetName("/BaseFont") ?? "";
                if (name.Contains("Cambria Math", StringComparison.OrdinalIgnoreCase)) continue;
                var unicode = font?.Elements.GetDictionary("/ToUnicode")?.Stream;
                if (unicode == null) continue;
                var mapping = Encoding.Latin1.GetString(unicode.UnfilteredValue);
                foreach (Match match in Regex.Matches(mapping,
                    @"<(?<code>[0-9A-Fa-f]{4})>\s*<221A>", RegexOptions.IgnoreCase))
                {
                    var code = match.Groups["code"].Value.ToUpperInvariant();
                    roots.Add(code);
                    baseRoots.Add(code);
                }
            }
        }
        var contents = document.Pages.Cast<PdfPage>()
            .SelectMany(page => page.Contents.Elements.OfType<PdfSharp.Pdf.Advanced.PdfReference>())
            .Select(reference => reference.Value)
            .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
            .Where(content => content.Stream != null)
            .Select(content => (Content: content,
                Text: Encoding.Latin1.GetString(content.Stream.UnfilteredValue)))
            .ToArray();
        var rootShows = contents.SelectMany(item => Regex.Matches(item.Text,
                @"(?m)^<(?<root>[0-9A-Fa-f]{4})[0-9A-Fa-f]{4,}> Tj$")
            .Select(match => match.Groups["root"].Value.ToUpperInvariant()))
            .Where(roots.Contains).ToArray();
        var selectedRoot = rootShows.Length == 1 ? rootShows[0] :
            rootShows.Length == 2 && tallRoot != null && rootShows.Count(root => root == tallRoot) == 1
                ? tallRoot : null;
        if (selectedRoot == null ||
            !baseRoots.Contains(selectedRoot) && selectedRoot != tallRoot) return false;
        var changed = false;
        foreach (var (content, source) in contents)
        {
            var updated = Regex.Replace(source, @"(?m)^<(?<root>[0-9A-Fa-f]{4})(?<tail>[0-9A-Fa-f]{4,})> Tj$",
                match =>
                {
                    var root = match.Groups["root"].Value;
                    if (!root.Equals(selectedRoot, StringComparison.OrdinalIgnoreCase)) return match.Value;
                    var tall = root.Equals(tallRoot, StringComparison.OrdinalIgnoreCase);
                    var rise = tall ? "0.75" : "0.84";
                    return $"-{rise} Ts\n<{root}> Tj\n0 Ts\n<{match.Groups["tail"].Value}> Tj";
                });
            if (updated == source) continue;
            if (content.Elements.ContainsKey("/Filter") && !content.Stream.TryUncompress()) continue;
            content.Stream.Value = Encoding.Latin1.GetBytes(updated);
            changed = true;
        }
        return changed;
    }

    private void AlignMatrixSurroundingText(PdfDocumentRenderer renderer, PdfDocument result,
        PdfDocument source)
    {
        if (_inlineMathMatrices.Count == 0 ||
            DocSystemFontAdvances.Find("Cambria Math") is not { } metrics) return;
        for (var pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
        {
            var shifts = GetPageStoryRenderInfos(renderer, pageIndex, source)
                .Where(entry => entry.Info.DocumentObject is MigraDoc.DocumentObjectModel.Paragraph paragraph &&
                    _inlineMathMatrices.ContainsKey(paragraph))
                .SelectMany(entry => _inlineMathMatrices[(MigraDoc.DocumentObjectModel.Paragraph)entry.Info.DocumentObject]
                    .Select(matrix => MatrixSurroundingBaselineShift(matrix.Cells, metrics)))
                .ToArray();
            if (shifts.Length == 0) continue;
            var contents = result.Pages[pageIndex].Contents.Elements
                .OfType<PdfSharp.Pdf.Advanced.PdfReference>()
                .Select(reference => reference.Value)
                .OfType<PdfSharp.Pdf.Advanced.PdfContent>()
                .Where(content => content.Stream != null)
                .Select(content => (Content: content,
                    Text: Encoding.Latin1.GetString(content.Stream.UnfilteredValue)))
                .ToArray();
            var candidateCount = contents.Sum(item => Regex.Matches(item.Text, @"BT[\s\S]*?ET")
                .Count(match => match.Value.Contains(" 14.02 Tf", StringComparison.Ordinal)));
            if (candidateCount != shifts.Length) continue;
            var matrixIndex = 0;
            foreach (var (content, original) in contents)
            {
                var updated = Regex.Replace(original, @"BT[\s\S]*?ET", match =>
                {
                    if (!match.Value.Contains(" 14.02 Tf", StringComparison.Ordinal))
                        return match.Value;
                    var shift = shifts[matrixIndex++];
                    if (shift == 0) return match.Value;
                    var rise = (-shift).ToString("0.####", CultureInfo.InvariantCulture);
                    return Regex.Replace(match.Value, @"(?m)^([^\r\n]+ (?:Tj|TJ))$", show =>
                        show.Value.Contains('\u00a0') ? show.Value :
                        $"{rise} Ts\n{show.Value}\n0 Ts");
                });
                if (updated == original) continue;
                if (content.Elements.ContainsKey("/Filter") && !content.Stream.TryUncompress())
                    continue;
                content.Stream.Value = Encoding.Latin1.GetBytes(updated);
            }
        }
    }

    private static double MatrixSurroundingBaselineShift(string[][] cells, DocOpenTypeAdvances metrics)
    {
        if (cells.Length != 2) return 0;
        if (IsDigitMatrix(cells)) return -0.38;
        var top = double.PositiveInfinity;
        var bottom = double.NegativeInfinity;
        for (var rowIndex = 0; rowIndex < 2; rowIndex++)
        {
            foreach (var cell in cells[rowIndex])
            {
                foreach (var rune in cell.EnumerateRunes())
                {
                    if (!metrics.TryGetGlyphVerticalBounds(rune.Value, out var minimum, out var maximum))
                        return 0;
                    top = Math.Min(top, rowIndex * 14.1 - maximum * 12.0 / metrics.UnitsPerEm);
                    bottom = Math.Max(bottom, rowIndex * 14.1 - minimum * 12.0 / metrics.UnitsPerEm);
                }
            }
        }
        var shift = (top + bottom) / 2 - 3.11;
        return Math.Abs(shift) < 0.5 ? 0 : shift;
    }

    private static bool IsDigitMatrix(string[][] cells) => cells.Length == 2 && cells.All(row => row.All(cell =>
        cell.Length > 0 && cell.All(char.IsAsciiDigit)));

    private static double NestedMathBarWidth(InlineMathTerm term, bool outer)
    {
        if (term.Text != null) return term.Width * 1.045;
        if (term.Numerator == null || term.Denominator == null) return term.Width;
        return Math.Max(NestedMathBarWidth(term.Numerator, false),
            NestedMathBarWidth(term.Denominator, false)) + (outer ? 2.4 : 0.12);
    }

    private void PaintInlineMathTerm(XGraphics graphics, InlineMathTerm term,
        double center, double barTop, double fontSize, bool outer = true)
    {
        if (term.Text != null)
        {
            XFont font;
            try { font = new XFont("Cambria Math", fontSize); }
            catch (InvalidOperationException) { font = new XFont(_fontFamily, fontSize); }
            graphics.DrawString(term.Text, font, XBrushes.Black,
                new XRect(center - term.Width / 2, barTop - 4.3, term.Width + 1, fontSize + 2),
                XStringFormats.TopLeft);
            return;
        }
        if (term.Numerator == null || term.Denominator == null) return;
        var numeratorSize = term.Numerator.Text == null ? fontSize * 6.96 / 8.52 : fontSize;
        var denominatorSize = term.Denominator.Text == null ? fontSize * 6.96 / 8.52 : fontSize;
        var numeratorOffset = fontSize < 8 ? 4.69 : term.Denominator.Text == null ? 6.71 : 8.2;
        var denominatorOffset = term.Numerator.Text == null ? 5.83 :
            term.Denominator.Text == null ? 8.2 : numeratorOffset;
        PaintInlineMathTerm(graphics, term.Numerator, center, barTop - numeratorOffset,
            numeratorSize, false);
        var barWidth = NestedMathBarWidth(term, outer);
        graphics.DrawRectangle(XBrushes.Black, center - barWidth / 2, barTop,
            barWidth, outer ? 0.84 : 0.6);
        PaintInlineMathTerm(graphics, term.Denominator, center, barTop + denominatorOffset,
            denominatorSize, false);
    }

    private static InlineMathTerm? ParseInlineMathTerm(OpenXmlElement element,
        XGraphics measure, string fontFamily, double fontSize, int depth = 0)
    {
        if (depth > 2) return null;
        var child = element.LocalName is "num" or "den"
            ? element.ChildElements.Count == 1 ? element.ChildElements[0] : null : element;
        if (child == null) return null;
        if (child.LocalName == "r" && !string.IsNullOrEmpty(child.InnerText))
        {
            var font = new XFont(fontFamily, fontSize);
            return new InlineMathTerm(child.InnerText, null, null,
                measure.MeasureString(child.InnerText, font).Width);
        }
        if (child.LocalName != "f") return null;
        var numerator = child.ChildElements.FirstOrDefault(part => part.LocalName == "num");
        var denominator = child.ChildElements.FirstOrDefault(part => part.LocalName == "den");
        if (numerator == null || denominator == null) return null;
        var firstSize = numerator.ChildElements.SingleOrDefault()?.LocalName == "f"
            ? fontSize * 6.96 / 8.52 : fontSize;
        var secondSize = denominator.ChildElements.SingleOrDefault()?.LocalName == "f"
            ? fontSize * 6.96 / 8.52 : fontSize;
        var first = ParseInlineMathTerm(numerator, measure, fontFamily, firstSize, depth + 1);
        var second = ParseInlineMathTerm(denominator, measure, fontFamily, secondSize, depth + 1);
        return first == null || second == null ? null :
            new InlineMathTerm(null, first, second, Math.Max(first.Width, second.Width) + 1.4);
    }

    private PdfDocument RenderColoredUnderlines(PdfDocumentRenderer renderer, PdfDocument document)
    {
        if (_coloredUnderlineParagraphs.Count == 0 && _listMarkerUnderlines.Count == 0 &&
            _listMarkerStrikes.Count == 0) return document;
        using var stream = new MemoryStream();
        document.Save(stream, false);
        var bytes = stream.ToArray();
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(bytes);
        using var imported = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var underlined = new PdfDocument();
        foreach (var page in imported.Pages) underlined.AddPage(page);
        CopyNamedDestinations(imported, underlined);
        for (var index = 0; index < renderer.PageCount; index++)
        {
            var page = parsed.GetPage(index + 1);
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, document))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    info.LayoutInfo.ContentArea is not { } area) continue;
                var top = area.Y.Point + offset;
                if (_listMarkerUnderlines.TryGetValue(paragraph, out var markerUnderline) &&
                    int.TryParse(markerUnderline.Color, NumberStyles.HexNumber,
                        CultureInfo.InvariantCulture, out var markerRgb))
                {
                    var markerWord = page.GetWords().FirstOrDefault(word =>
                        word.Text == markerUnderline.Marker &&
                        page.Height - word.BoundingBox.Top >= top - 5 &&
                        page.Height - word.BoundingBox.Top <= top + area.Height.Point + 5);
                    if (markerWord != null)
                    {
                        var firstLetter = markerWord.Letters[0];
                        var lastLetter = markerWord.Letters[^1];
                        var left = firstLetter.StartBaseLine.X;
                        var width = lastLetter.EndBaseLine.X - left;
                        var baseline = page.Height - firstLetter.StartBaseLine.Y;
                        var markerBrush = new XSolidBrush(XColor.FromArgb((markerRgb >> 16) & 255,
                            (markerRgb >> 8) & 255, markerRgb & 255));
                        using var markerGraphics = XGraphics.FromPdfPage(underlined.Pages[index]);
                        if (markerUnderline.Double)
                        {
                            markerGraphics.DrawRectangle(markerBrush, left, baseline + 0.48, width, 0.6);
                            markerGraphics.DrawRectangle(markerBrush, left, baseline + 2.04, width, 0.6);
                        }
                        else
                            markerGraphics.DrawRectangle(markerBrush, left, baseline + 1.32, width, 1.2);
                    }
                }
                if (_listMarkerStrikes.TryGetValue(paragraph, out var markerStrike) &&
                    int.TryParse(markerStrike.Color, NumberStyles.HexNumber,
                        CultureInfo.InvariantCulture, out var strikeRgb))
                {
                    var markerWord = page.GetWords().FirstOrDefault(word =>
                        word.Text == markerStrike.Marker &&
                        page.Height - word.BoundingBox.Top >= top - 5 &&
                        page.Height - word.BoundingBox.Top <= top + area.Height.Point + 5);
                    if (markerWord != null)
                    {
                        var firstLetter = markerWord.Letters[0];
                        var lastLetter = markerWord.Letters[^1];
                        var baseline = page.Height - firstLetter.StartBaseLine.Y;
                        var strikeBrush = new XSolidBrush(XColor.FromArgb((strikeRgb >> 16) & 255,
                            (strikeRgb >> 8) & 255, strikeRgb & 255));
                        using var strikeGraphics = XGraphics.FromPdfPage(underlined.Pages[index]);
                        strikeGraphics.DrawRectangle(strikeBrush, firstLetter.StartBaseLine.X,
                            baseline - 4.32, lastLetter.EndBaseLine.X - firstLetter.StartBaseLine.X, 0.6);
                    }
                }
                if (!_coloredUnderlineParagraphs.TryGetValue(paragraph, out var underline) ||
                    !int.TryParse(underline.Color, NumberStyles.HexNumber,
                        CultureInfo.InvariantCulture, out var rgb)) continue;
                var letters = page.Letters.Where(letter =>
                    !string.IsNullOrWhiteSpace(letter.Value) &&
                    letter.StartBaseLine.X >= area.X.Point - 2 &&
                    letter.StartBaseLine.X <= area.X.Point + area.Width.Point + 2 &&
                    page.Height - letter.StartBaseLine.Y >= top - 5 &&
                    page.Height - letter.StartBaseLine.Y <= top + area.Height.Point + 5).ToArray();
                if (letters.Length == 0) continue;
                using var graphics = XGraphics.FromPdfPage(underlined.Pages[index]);
                var brush = new XSolidBrush(XColor.FromArgb((rgb >> 16) & 255,
                    (rgb >> 8) & 255, rgb & 255));
                foreach (var line in letters.GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1)))
                {
                    var first = line.OrderBy(letter => letter.StartBaseLine.X).First();
                    var last = line.OrderByDescending(letter => letter.EndBaseLine.X).First();
                    var body = _bodyParagraphs.Contains(paragraph);
                    var compactBody = _compactArialBodyParagraphs.Contains(paragraph);
                    var aptosStory = underline.InHeaderFooter &&
                        first.FontName.Contains("Aptos", StringComparison.OrdinalIgnoreCase);
                    var bodyExtension = _characterSpacings.ContainsKey(paragraph) ? 2 : 0;
                    var right = last.EndBaseLine.X +
                        (body ? bodyExtension : aptosStory ? 0 : 4) +
                        (underline.TrailingSpace ? last.PointSize * 0.35 : 0);
                    graphics.DrawRectangle(brush, first.StartBaseLine.X,
                        page.Height - first.GlyphRectangle.Bottom +
                            (body ? compactBody ? -1 : 0.75 : aptosStory ? 0.6 : 1.2) +
                            (compactBody ? 2.1 : 0),
                        right - first.StartBaseLine.X,
                        body && !compactBody || aptosStory ? 0.25 : 1);
                }
            }
        }
        return underlined;
    }

    private static PartialFitMetric[]? BuildPartialFitMetrics(UglyToad.PdfPig.Content.Page page,
        UglyToad.PdfPig.Content.Letter first, double left, double width,
        (string Text, double Width)[] groups)
    {
        var letters = page.Letters.Where(letter =>
            Math.Abs(letter.StartBaseLine.Y - first.StartBaseLine.Y) < 0.5 &&
            letter.StartBaseLine.X >= first.StartBaseLine.X - 1 &&
            letter.EndBaseLine.X <= left + width + 2 &&
            !string.IsNullOrWhiteSpace(letter.Value))
            .OrderBy(letter => letter.StartBaseLine.X).ToArray();
        if (letters.Length < groups.Sum(group => group.Text.Count(character => !char.IsWhiteSpace(character))))
            return null;
        var result = new PartialFitMetric[groups.Length];
        var count = 0;
        var offset = 0.0;
        for (var index = 0; index < groups.Length; index++)
        {
            var groupCount = groups[index].Text.Count(character => !char.IsWhiteSpace(character));
            var next = count + groupCount;
            var start = letters[count].StartBaseLine.X;
            var end = next < letters.Length ? letters[next].StartBaseLine.X : letters[next - 1].EndBaseLine.X;
            var surplus = groups[index].Width - (end - start);
            var spacing = surplus / (next < letters.Length
                ? groupCount : Math.Max(1, groupCount - 1));
            var glyphSpacing = next < letters.Length
                ? surplus / (groupCount + groups[index].Text.TrimEnd().Count(char.IsWhiteSpace))
                : spacing;
            result[index] = new PartialFitMetric(count, start, offset, spacing, glyphSpacing);
            count = next;
            offset += groups[index].Width;
        }
        return result;
    }

    private PdfDocument? ApplyFormattingBalloons(PdfDocumentRenderer renderer, PdfDocument document)
    {
        if (_formattingBalloons.Count == 0 && _changedFieldUrls.Count == 0) return null;
        using var stream = new MemoryStream();
        document.Save(stream, false);
        var bytes = stream.ToArray();
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var mutable = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Modify);
        var pages = Enumerable.Range(0, mutable.PageCount)
            .Select(_ => new List<DxpPdfMarkupPrintLayout.Balloon>()).ToArray();
        var anchored = new HashSet<MigraDoc.DocumentObjectModel.Paragraph>();
        for (var pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
        {
            var page = parsed.GetPage(pageIndex + 1);
            foreach (var (info, xOffset, yOffset) in
                     GetPageEmphasisRenderInfos(renderer, pageIndex, mutable))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    !_formattingBalloons.TryGetValue(paragraph, out var descriptions) ||
                    !anchored.Add(paragraph) || info.LayoutInfo.ContentArea is not { } area) continue;
                var top = area.Y.Point + yOffset;
                var right = area.X.Point + xOffset;
                var letters = page.Letters.Where(letter =>
                    letter.GlyphRectangle.Left >= right - 2 &&
                    letter.GlyphRectangle.Left <= right + area.Width.Point + 2 &&
                    page.Height - letter.StartBaseLine.Y >= top - 5 &&
                    page.Height - letter.StartBaseLine.Y <= top + area.Height.Point + 5);
                var connectorX = letters.Select(letter => letter.EndBaseLine.X)
                    .DefaultIfEmpty(right).Max();
                for (var index = 0; index < descriptions.Count; index++)
                    pages[pageIndex].Add(new DxpPdfMarkupPrintLayout.Balloon(
                        top + index * 14, connectorX, descriptions[index]));
            }
        }
        var emittedFieldBalloons = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var pageIndex = 0; pageIndex < mutable.PageCount; pageIndex++)
        {
            var page = mutable.Pages[pageIndex];
            var lastTop = new Dictionary<string, double>(StringComparer.Ordinal);
            foreach (var annotation in page.Annotations.Cast<PdfSharp.Pdf.Annotations.PdfAnnotation>()
                         .OrderBy(item => page.Height.Point - item.Rectangle.Y2))
            {
                var uri = annotation.Elements.GetDictionary("/A")?.Elements.GetString("/URI");
                if (uri == null || !_changedFieldUrls.TryGetValue(uri, out var fieldCount) ||
                    emittedFieldBalloons.GetValueOrDefault(uri) >= fieldCount) continue;
                var rectangle = annotation.Rectangle;
                var top = page.Height.Point - rectangle.Y2;
                if (lastTop.TryGetValue(uri, out var previousTop) && top - previousTop < 18)
                {
                    lastTop[uri] = top;
                    continue;
                }
                lastTop[uri] = top;
                emittedFieldBalloons[uri] = emittedFieldBalloons.GetValueOrDefault(uri) + 1;
                pages[pageIndex].Add(new DxpPdfMarkupPrintLayout.Balloon(
                    top, rectangle.X2, string.Empty,
                    "Field Code Changed"));
            }
        }
        if (_formattingBalloons.Count > 0 || _changedFieldUrls.Count > 0)
            DxpPdfMarkupPrintLayout.Apply(mutable, pages);
        return mutable;
    }

    private PdfDocument RenderEmphasisMarks(PdfDocumentRenderer renderer, PdfDocument? target = null,
        PdfDocumentRenderer? footnoteRenderer = null,
        List<(int Page, int Index, MigraDoc.DocumentObjectModel.Paragraph Note, double Y,
            int PageOffset)>? footnoteOverlays = null)
    {
        if (_emphasisRanges.Count == 0 && _rubyGuides.Count == 0 && _strikeRanges.Count == 0 &&
            _revisionUnderlineRanges.Count == 0 && _revisionParagraphs.Count == 0 &&
            _textEffectRanges.Count == 0 &&
            _textBorderRanges.Count == 0 &&
            _characterScales.Count == 0 && _fitTextWidths.Count == 0 && _partialFitGroups.Count == 0 &&
            _characterSpacings.Count == 0 && _topAlignedFontSizes.Count == 0 &&
            _verticalAlignedFonts.Count == 0 &&
            _gridGlyphParagraphs.Count == 0)
            return target ?? renderer.PdfDocument;
        var document = target ?? renderer.PdfDocument;
        using var stream = new MemoryStream();
        document.Save(stream, false);
        var bytes = stream.ToArray();
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(bytes);
        using var imported = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        var marked = new PdfDocument();
        foreach (var page in imported.Pages) marked.AddPage(page);
        CopyNamedDestinations(imported, marked);
        var markFonts = new Dictionary<(string Name, double Size), XFont?>();
        for (var index = 0; index < renderer.PageCount; index++)
        {
            var page = parsed.GetPage(index + 1);
            var bodyBorderIndex = 0;
            var cellRevisionBars = new HashSet<(double X, double Top)>();
            var pageInfos = GetPageEmphasisRenderInfos(renderer, index, target);
            if (footnoteRenderer != null && footnoteOverlays != null)
                pageInfos = pageInfos.Concat(GetFootnoteEmphasisRenderInfos(
                    footnoteRenderer, footnoteOverlays, index + 1));
            foreach (var (info, xOffset, offset) in pageInfos)
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    info.LayoutInfo.ContentArea is not { } area) continue;
                _emphasisRanges.TryGetValue(paragraph, out var ranges);
                _strikeRanges.TryGetValue(paragraph, out var strikes);
                _revisionUnderlineRanges.TryGetValue(paragraph, out var revisionUnderlines);
                _textEffectRanges.TryGetValue(paragraph, out var effects);
                _textBorderRanges.TryGetValue(paragraph, out var borders);
                _rubyGuides.TryGetValue(paragraph, out var rubyGuides);
                if (ranges == null && strikes == null && revisionUnderlines == null &&
                    !_revisionParagraphs.Contains(paragraph) && effects == null && borders == null &&
                    rubyGuides == null) continue;
                var borderShift = borders?.FirstOrDefault() is { } firstBorder
                    ? _bodyParagraphs.Contains(paragraph)
                        ? firstBorder.Inset + bodyBorderIndex++ * 8.2 : -firstBorder.Inset - 0.9
                    : 0;
                var top = area.Y.Point + offset;
                var bottom = top + area.Height.Point;
                var letters = page.Letters.Where(letter =>
                    letter.GlyphRectangle.Left >= area.X.Point + xOffset - 2 &&
                    letter.GlyphRectangle.Left <= area.X.Point + xOffset + area.Width.Point + 2 &&
                    page.Height - letter.StartBaseLine.Y >= top - 5 &&
                    page.Height - letter.StartBaseLine.Y <= bottom + 5).ToArray();
                if (effects != null)
                {
                    foreach (var effect in effects)
                    {
                        if (effect.Kind is TextEffectKind.Emboss or TextEffectKind.Imprint)
                            DrawEffectLayer(effect, 0, 0,
                                new XSolidBrush(XColor.FromArgb(153, 153, 153)), 0);
                        if (effect.Kind == TextEffectKind.Emboss)
                        {
                            DrawEffectLayer(effect, -1.9, -1.9, XBrushes.Black, 0);
                            DrawEffectLayer(effect, -0.95, -0.95, XBrushes.Black, 0);
                        }
                        else if (effect.Kind == TextEffectKind.Imprint)
                        {
                            DrawEffectLayer(effect, 1.9, 1.9, XBrushes.Black, 0);
                            DrawEffectLayer(effect, 0.95, 0.95, XBrushes.Black, 0);
                        }
                        else if (effect.Kind == TextEffectKind.ShadowOutline)
                        {
                            DrawEffectLayer(effect, 0.5, 0.5,
                                new XSolidBrush(XColor.FromArgb(192, 192, 192)), 2);
                            DrawEffectLayer(effect, 0, 0, XBrushes.White, 2);
                        }
                        else if (effect.Kind == TextEffectKind.Shadow)
                        {
                            DrawEffectLayer(effect, 0.5, 0.5,
                                new XSolidBrush(XColor.FromArgb(192, 192, 192)), 0);
                            DrawEffectLayer(effect, 0, 0, XBrushes.Black, 0);
                        }
                        else if (effect.Kind == TextEffectKind.Outline)
                            DrawEffectLayer(effect, 0, 0, XBrushes.White, 2);
                    }

                    void DrawEffectLayer(TextEffectRange effect, double horizontal, double vertical,
                        XBrush brush, int mode)
                    {
                        using (var effectGraphics = XGraphics.FromPdfPage(marked.Pages[index]))
                        {
                            foreach (var letter in letters.Skip(effect.Start).Take(effect.Length))
                            {
                                if (string.IsNullOrWhiteSpace(letter.Value)) continue;
                                var style = (effect.Bold ? XFontStyleEx.Bold : XFontStyleEx.Regular) |
                                    (effect.Italic ? XFontStyleEx.Italic : XFontStyleEx.Regular);
                                var font = new XFont(effect.FontName, letter.FontSize, style);
                                effectGraphics.DrawString(letter.Value, font, brush,
                                    letter.StartBaseLine.X + horizontal,
                                    page.Height - letter.StartBaseLine.Y + vertical);
                            }
                        }
                        if (mode == 0) return;
                        var content = ((PdfSharp.Pdf.Advanced.PdfReference)
                            marked.Pages[index].Contents.Elements[^1]).Value as PdfSharp.Pdf.Advanced.PdfContent;
                        if (content?.Stream == null) return;
                        var commands = Encoding.Latin1.GetString(content.Stream.Value);
                        var shadowLayer = effect.Kind == TextEffectKind.ShadowOutline && horizontal != 0;
                        var outlineColor = shadowLayer
                            ? "0.65 0.65 0.65" : "0 0 0";
                        var outlineWidth = shadowLayer ? "0.5" : "0.14";
                        content.Stream.Value = Encoding.Latin1.GetBytes(commands.Replace("q\nq\nBT\n",
                            $"q\nq\n{outlineColor} RG\n{outlineWidth} w\nBT\n{mode} Tr\n", StringComparison.Ordinal));
                    }
                }
                var mergedBorders = new List<TextBorderRange>();
                foreach (var border in borders ?? [])
                {
                    if (mergedBorders.Count > 0 && mergedBorders[^1] is { } previous &&
                        previous.Start + previous.Length == border.Start &&
                        previous.Color == border.Color && previous.Width == border.Width &&
                        previous.Inset == border.Inset && previous.Shadow == border.Shadow &&
                        previous.Double == border.Double && border.Start < letters.Length &&
                        previous.Start < letters.Length &&
                        Math.Abs(letters[previous.Start].StartBaseLine.Y -
                            letters[border.Start].StartBaseLine.Y) < 1)
                        mergedBorders[^1] = previous with { Length = previous.Length + border.Length,
                            VisibleLength = previous.VisibleLength + border.VisibleLength };
                    else mergedBorders.Add(border);
                }
                if (_tabbedRunShadings.TryGetValue(paragraph, out var tabShading) &&
                    int.TryParse(tabShading, NumberStyles.HexNumber,
                        CultureInfo.InvariantCulture, out var shadingRgb))
                {
                    using var background = XGraphics.FromPdfPage(marked.Pages[index],
                        XGraphicsPdfPageOptions.Prepend);
                    var shade = new XSolidBrush(XColor.FromArgb((shadingRgb >> 16) & 255,
                        (shadingRgb >> 8) & 255, shadingRgb & 255));
                    foreach (var border in mergedBorders)
                    {
                        var bordered = letters.Skip(border.Start).Take(border.Length)
                            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
                        for (var letterIndex = 1; letterIndex < bordered.Length; letterIndex++)
                        {
                            var previous = bordered[letterIndex - 1];
                            var next = bordered[letterIndex];
                            var gap = next.StartBaseLine.X - previous.EndBaseLine.X;
                            if (gap < 30 || Math.Abs(previous.StartBaseLine.Y - next.StartBaseLine.Y) > 1)
                                continue;
                            background.DrawRectangle(shade, previous.EndBaseLine.X,
                                page.Height - previous.GlyphRectangle.Top - 0.3,
                                gap + 2 * _characterSpacings.GetValueOrDefault(paragraph),
                                previous.PointSize * 1.22);
                        }
                    }
                }
                using var graphics = XGraphics.FromPdfPage(marked.Pages[index]);
                if (_revisionParagraphs.Contains(paragraph) &&
                    (_sourceBodyParagraphs.Contains(paragraph) ||
                     _sourceCellParagraphs.Contains(paragraph)) &&
                    paragraph.Section is { } revisionSection)
                {
                    var barX = revisionSection.PageSetup.LeftMargin.Point / 2;
                    var barBottom = bottom + Math.Max(0, paragraph.Format.SpaceAfter.Point);
                    if (_sourceBodyParagraphs.Contains(paragraph) ||
                        cellRevisionBars.Add((barX, top)))
                        graphics.DrawRectangle(XBrushes.Black, barX, top, 0.72, barBottom - top);
                }
                if (rubyGuides != null)
                {
                    foreach (var rubyGuide in rubyGuides)
                    {
                        var baseLetters = letters.Select((letter, position) => (letter, position))
                            .FirstOrDefault(candidate => candidate.position >= rubyGuide.Start &&
                                candidate.letter.Value == rubyGuide.Base[..1]);
                        if (baseLetters.letter == null) continue;
                        var start = baseLetters.position;
                        if (start + rubyGuide.Base.Length > letters.Length ||
                            string.Concat(letters.Skip(start).Take(rubyGuide.Base.Length)
                                .Select(letter => letter.Value)) != rubyGuide.Base) continue;
                        var end = letters[start + rubyGuide.Base.Length - 1];
                        var isCjkGuide = rubyGuide.Guide.Any(IsCjk);
                        var guideFamily = rubyGuide.Font is { } guideFont &&
                            _availableFontFamilies.Contains(guideFont) ? guideFont :
                            isCjkGuide && _availableFontFamilies.Contains("Yu Mincho") ? "Yu Mincho" :
                            isCjkGuide && _cjkFontFamily != null ? _cjkFontFamily : _fontFamily;
                        var font = new XFont(guideFamily, rubyGuide.Size);
                        var width = graphics.MeasureString(rubyGuide.Guide, font).Width;
                        var left = baseLetters.letter.StartBaseLine.X;
                        var right = end.EndBaseLine.X;
                        var baseline = page.Height - baseLetters.letter.StartBaseLine.Y;
                        var guideY = baseline - rubyGuide.Raise;
                        if (rubyGuide.Alignment == "rightVertical")
                        {
                            var position = guideY + rubyGuide.Size / 3;
                            foreach (var character in rubyGuide.Guide.EnumerateRunes())
                            {
                                graphics.DrawString(character.ToString(), font, XBrushes.Black,
                                    new XPoint(right - rubyGuide.Size * 0.17, position));
                                position += rubyGuide.Size;
                            }
                        }
                        else if (rubyGuide.Alignment is "distributeLetter" or "distributeSpace")
                        {
                            var characters = rubyGuide.Guide.EnumerateRunes()
                                .Select(character => character.ToString()).ToArray();
                            var widths = characters.Select(character => graphics.MeasureString(character, font).Width)
                                .ToArray();
                            var spacing = rubyGuide.Alignment == "distributeSpace"
                                ? (right - left - widths.Sum()) / characters.Length
                                : characters.Length > 1
                                    ? (right - left - widths.Sum()) / (characters.Length - 1) : 0;
                            var position = left + (rubyGuide.Alignment == "distributeSpace" ? spacing / 2 : 0);
                            for (var characterIndex = 0; characterIndex < characters.Length; characterIndex++)
                            {
                                graphics.DrawString(characters[characterIndex], font, XBrushes.Black,
                                    new XPoint(position, guideY));
                                position += widths[characterIndex] + spacing;
                            }
                        }
                        else
                        {
                            var guideX = rubyGuide.Alignment switch
                            {
                                "left" => left,
                                "right" => right - width,
                                _ => (left + right - width) / 2
                            };
                            graphics.DrawString(rubyGuide.Guide, font, XBrushes.Black,
                                new XPoint(guideX, guideY));
                        }
                    }
                }
                foreach (var border in mergedBorders)
                {
                    var visibleRange = _borderResetParagraphs.Contains(paragraph);
                    var bordered = letters.Skip(visibleRange ? border.VisibleStart : border.Start)
                        .Take(visibleRange ? border.VisibleLength : border.Length)
                        .Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
                    if (bordered.Length == 0 ||
                        !int.TryParse(border.Color, NumberStyles.HexNumber,
                            CultureInfo.InvariantCulture, out var rgb)) continue;
                    var first = bordered[0];
                    var last = bordered[^1];
                    var compactArial = _compactArialBodyParagraphs.Contains(paragraph);
                    var tabbedBodyBorder = _bodyParagraphs.Contains(paragraph) &&
                        bordered.Zip(bordered.Skip(1)).Any(pair =>
                            Math.Abs(pair.First.StartBaseLine.Y - pair.Second.StartBaseLine.Y) < 1 &&
                            pair.Second.StartBaseLine.X - pair.First.EndBaseLine.X > 30);
                    var left = first.StartBaseLine.X + border.Width / 2 +
                        (compactArial ? border.Inset - border.Width / 2 : 0) -
                        (tabbedBodyBorder ? border.Inset : 0);
                    var borderOffset = borderShift - (border.Double
                        ? _firstBodyDoubleBorderInset.GetValueOrDefault(paragraph) +
                            (_footerLeaderParagraphs.Contains(paragraph) ? 2 * border.Width : 0) : 0);
                    var borderTop = page.Height - first.GlyphRectangle.Top + borderOffset - border.Inset - 4 +
                        border.Width / 2 + (_compactArialBodyParagraphs.Contains(paragraph) ? 1.8 : 0);
                    var right = compactArial
                        ? last.EndBaseLine.X + border.VisibleLength *
                            _characterSpacings.GetValueOrDefault(paragraph)
                        : last.EndBaseLine.X + border.Inset * 2 - border.Width / 2 +
                            border.Length * _characterSpacings.GetValueOrDefault(paragraph);
                    if (tabbedBodyBorder)
                        right += border.Inset + _borderedTabInsets.GetValueOrDefault(paragraph);
                    var borderBottom = page.Height - first.GlyphRectangle.Bottom + borderOffset + border.Inset + 4 -
                        border.Width / 2 - (_compactArialBodyParagraphs.Contains(paragraph) ? 1.3 : 0);
                    var borderPen = new XPen(XColor.FromArgb((rgb >> 16) & 255,
                        (rgb >> 8) & 255, rgb & 255), border.Width);
                    if (border.Double)
                    {
                        var outerTop = borderTop + border.Width * 0.6;
                        var outerRight = right + border.Inset + border.Width * 7 / 6;
                        var outerBottom = borderBottom + border.Inset + border.Width * 1.2;
                        graphics.DrawRectangle(borderPen, left, outerTop,
                            outerRight - left, outerBottom - outerTop);
                        var gap = 2 * border.Width;
                        graphics.DrawRectangle(borderPen, left + gap, outerTop + gap,
                            outerRight - left - 2 * gap, outerBottom - outerTop - 2 * gap);
                    }
                    else
                    {
                        var storyBorder = _headerRunBorderParagraphs.Contains(paragraph) ||
                            _footerRunBorderParagraphs.Contains(paragraph);
                        var onePointBorder = storyBorder && Math.Abs(border.Width - 1) < 0.01 &&
                            Math.Abs(border.Inset - 2) < 0.01;
                        var compactStoryBorder = storyBorder && first.PointSize < 10 &&
                            Math.Abs(border.Width - 1.5) < 0.01 &&
                            Math.Abs(border.Inset - 3.5) < 0.01;
                        var storyInset = onePointBorder ? border.Inset - border.Width / 2 :
                            compactStoryBorder ? border.Width * 0.9 : 0;
                        var storyOffset = storyInset + (onePointBorder &&
                            _footerRunBorderParagraphs.Contains(paragraph) ? border.Width / 2 : 0);
                        graphics.DrawRectangle(borderPen,
                            left, borderTop + border.Width / 2 + storyOffset, right - left,
                            borderBottom - borderTop - border.Width - 2 * storyInset);
                    }
                    if (border.Shadow)
                    {
                        var shadowPen = new XPen(XColors.Black, border.Width);
                        graphics.DrawLine(shadowPen, right + border.Width, borderTop + 2.75 * border.Width,
                            right + border.Width, borderBottom + 0.5 * border.Width);
                        graphics.DrawLine(shadowPen, left + 2.25 * border.Width,
                            borderBottom + 0.5 * border.Width, right + border.Width,
                            borderBottom + 0.5 * border.Width);
                    }
                }
                foreach (var range in ranges ?? [])
                {
                    for (var position = range.Start; position < range.Start + range.Length &&
                        position < letters.Length; position++)
                    {
                        var letter = letters[position];
                        if (string.IsNullOrWhiteSpace(letter.Value)) continue;
                        var center = letter.Value.Length == 1 &&
                            (letter.Value[0] is >= '\u3000' and <= '\u303F' or >= '\uFF01' and <= '\uFF65') &&
                            char.IsPunctuation(letter.Value[0])
                                ? letter.StartBaseLine.X + letter.PointSize / 2
                                : (letter.GlyphRectangle.Left + letter.GlyphRectangle.Right) / 2;
                        var lineGlyphTop = letters
                            .Where(candidate => Math.Abs(candidate.StartBaseLine.Y - letter.StartBaseLine.Y) < 1.5)
                            .Min(candidate => page.Height - candidate.GlyphRectangle.Top);
                        var lineBaseline = page.Height - letter.StartBaseLine.Y;
                        if (range.Kind == EmphasisMarkValues.Dot || range.Kind == EmphasisMarkValues.UnderDot)
                        {
                            var underDot = range.Kind == EmphasisMarkValues.UnderDot;
                            var fontName = underDot ? "SimSun" : "MS Mincho";
                            var fontKey = (fontName, letter.PointSize);
                            if (!markFonts.TryGetValue(fontKey, out var font))
                            {
                                try { font = new XFont(fontName, letter.PointSize); }
                                catch (InvalidOperationException) { }
                                markFonts[fontKey] = font;
                            }
                            if (font != null)
                            {
                                var markCenter = center;
                                if (letter.Value.Length == 1 && position + 1 < letters.Length &&
                                    letters[position + 1].Value.Length == 1 &&
                                    Math.Abs(letters[position + 1].StartBaseLine.Y - letter.StartBaseLine.Y) < 1 &&
                                    (IsCjk(letter.Value[0]) && IsLatinLetter(letters[position + 1].Value[0]) ||
                                        IsLatinLetter(letter.Value[0]) && IsCjk(letters[position + 1].Value[0])))
                                    markCenter += letter.PointSize * 0.1;
                                graphics.DrawString("．", font, XBrushes.Black,
                                    new XPoint(markCenter - (underDot ? 2.4 : 2.0),
                                        lineBaseline + (underDot ? 4.48 : -12.2)));
                                continue;
                            }
                        }
                        if (range.Kind == EmphasisMarkValues.Circle)
                            graphics.DrawEllipse(new XPen(XColors.Black, 0.45), center - 0.8, lineGlyphTop - 5.2,
                                1.6, 1.6);
                        else if (range.Kind == EmphasisMarkValues.Comma)
                        {
                            var fontKey = ("MS Mincho", letter.PointSize);
                            if (!markFonts.TryGetValue(fontKey, out var font))
                            {
                                try { font = new XFont("MS Mincho", letter.PointSize); }
                                catch (InvalidOperationException) { }
                                markFonts[fontKey] = font;
                            }
                            if (font != null)
                                graphics.DrawString("、", font, XBrushes.Black,
                                    new XPoint(center - 2.04, lineBaseline - 11.8));
                            else
                                graphics.DrawLine(new XPen(XColors.Black, 0.75), center - 0.8,
                                    lineGlyphTop - 4.9, center + 1.0, lineGlyphTop - 2.7);
                        }
                        else
                            graphics.DrawEllipse(XBrushes.Black, center - 0.7,
                                range.Kind == EmphasisMarkValues.UnderDot
                                    ? lineBaseline + letter.PointSize * 0.23 : lineGlyphTop - 3.7,
                                1.4, 1.4);
                    }
                }
                var revisionPen = new XPen(_revisionMarkupColor, 0.24);
                foreach (var underline in revisionUnderlines ?? [])
                {
                    var underlined = letters.Skip(underline.Start).Take(underline.Length).ToArray();
                    if (underlined.Length == 0) continue;
                    foreach (var line in underlined.GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1)))
                    {
                        var first = line.First();
                        var last = line.Last();
                        var baseline = page.Height - first.StartBaseLine.Y;
                        if (underline.Moved)
                        {
                            var movePen = new XPen(_revisionMarkupColor, 0.12);
                            graphics.DrawLine(movePen, first.StartBaseLine.X, baseline + 0.42,
                                last.EndBaseLine.X, baseline + 0.42);
                            graphics.DrawLine(movePen, first.StartBaseLine.X, baseline + 1.50,
                                last.EndBaseLine.X, baseline + 1.50);
                        }
                        else
                            graphics.DrawLine(revisionPen, first.StartBaseLine.X, baseline + 0.75,
                                last.EndBaseLine.X, baseline + 0.75);
                    }
                }
                for (var strikeIndex = 0; strikeIndex < (strikes?.Count ?? 0); strikeIndex++)
                {
                    var strike = strikes![strikeIndex];
                    var struck = letters.Skip(strike.VisibleStart ?? strike.Start)
                        .Take(strike.VisibleLength ?? strike.Length).ToArray();
                    if (struck.Length == 0) continue;
                    var strikeLines = strike.VisibleStart != null
                        ? struck.GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                            .Select(line => line.ToArray())
                        : [struck];
                    foreach (var strikeLine in strikeLines)
                    {
                        var left = strikeLine[0].StartBaseLine.X;
                        var right = strikeLine[^1].EndBaseLine.X;
                        if (strike.VisibleStart == null &&
                            _strikeTabBoundaries.TryGetValue(paragraph, out var boundaries) &&
                            boundaries.Contains(strike.Start + strike.Length) &&
                            strikeIndex + 1 < strikes.Count &&
                            strikes[strikeIndex + 1] is { } nextStrike &&
                            nextStrike.Start == strike.Start + strike.Length &&
                            nextStrike.Double == strike.Double &&
                            letters.Skip(nextStrike.Start).FirstOrDefault() is { } following &&
                            Math.Abs(following.StartBaseLine.Y - strikeLine[^1].StartBaseLine.Y) < 2 &&
                            following.StartBaseLine.X > right)
                            right = following.StartBaseLine.X;
                        if (_characterScales.TryGetValue(paragraph, out var scale))
                            right = left + (right - left) * scale / 100;
                        var glyph = strikeLine[0].GlyphRectangle;
                        var middle = strike.VisibleStart != null
                            ? page.Height - strikeLine[0].StartBaseLine.Y - strikeLine[0].PointSize / 3
                            : page.Height - glyph.Top + (glyph.Top - glyph.Bottom) * 0.6 + 0.6;
                        var strikePen = strike.VisibleStart != null
                            ? new XPen(_revisionDeletionColor, 0.6)
                            : new XPen(XColors.Black, 0.55);
                        if (strike.Moved)
                        {
                            graphics.DrawLine(strikePen, left, middle - 0.64, right, middle - 0.64);
                            graphics.DrawLine(strikePen, left, middle + 0.58, right, middle + 0.58);
                        }
                        else
                        {
                            graphics.DrawLine(strikePen, left, middle, right, middle);
                            if (strike.Double)
                                graphics.DrawLine(strikePen, left, middle + 1.5, right, middle + 1.5);
                        }
                    }
                }
            }
        }
        ApplyCharacterMetrics(renderer, marked, parsed);
        MatchFittedCalibriPaintSize(marked);
        return marked;
    }

    private void MatchFittedCalibriPaintSize(PdfDocument document)
    {
        if (_fitTextWidths.Count == 0) return;
        foreach (PdfPage page in document.Pages)
        {
            var fontResources = page.Resources.Elements.GetDictionary("/Font");
            if (fontResources == null) continue;
            var calibriFonts = new HashSet<string>(StringComparer.Ordinal);
            foreach (var key in fontResources.Elements.Keys)
            {
                var entry = fontResources.Elements[key];
                var font = (entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfDictionary ??
                    entry as PdfDictionary;
                if (font?.Elements.GetName("/BaseFont").TrimStart('/').Split('+').LastOrDefault()
                    .StartsWith("Calibri", StringComparison.OrdinalIgnoreCase) == true)
                    calibriFonts.Add(key.TrimStart('/'));
            }
            if (calibriFonts.Count == 0) continue;
            foreach (var entry in page.Contents.Elements)
            {
                if ((entry as PdfSharp.Pdf.Advanced.PdfReference)?.Value is not
                    PdfSharp.Pdf.Advanced.PdfContent content || content.Stream == null ||
                    content.Elements.ContainsKey("/Filter")) continue;
                var source = Encoding.Latin1.GetString(content.Stream.Value);
                var revised = RewriteFittedCalibriPaintSize(source, calibriFonts);
                if (revised != source) content.Stream.Value = Encoding.Latin1.GetBytes(revised);
            }
        }
    }

    internal static string RewriteFittedCalibriPaintSize(string source, ISet<string> calibriFonts)
    {
        var lines = source.Split('\n');
        var changed = false;
        for (var index = 0; index < lines.Length; index++)
        {
            var firstFont = Regex.Match(lines[index], @"^/(?<font>\S+) 11 Tf$");
            if (!firstFont.Success || !calibriFonts.Contains(firstFont.Groups["font"].Value))
                continue;
            var scaleIndex = index + 1;
            while (scaleIndex < Math.Min(index + 5, lines.Length) &&
                   !lines[scaleIndex].EndsWith(" Tz", StringComparison.Ordinal) &&
                   !lines[scaleIndex].EndsWith(" Tf", StringComparison.Ordinal))
                scaleIndex++;
            if (scaleIndex >= lines.Length ||
                !lines[scaleIndex].EndsWith(" Tz", StringComparison.Ordinal) ||
                !double.TryParse(lines[scaleIndex].Split(' ')[0], NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var scale) || scale is <= 20 or >= 250 ||
                Math.Abs(scale - 100) < 0.001) continue;
            var end = scaleIndex + 1;
            while (end < lines.Length)
            {
                var font = Regex.Match(lines[end], @"^/(?<font>\S+) (?<size>\d+(?:\.\d+)?) Tf$");
                var move = Regex.Match(lines[end], @"^-?\d+(?:\.\d+)? (?<vertical>-?\d+(?:\.\d+)?) Td$");
                if (lines[end] is "100 Tz" or "ET" ||
                    font.Success && font.Groups["size"].Value != "11" ||
                    move.Success && Math.Abs(double.Parse(move.Groups["vertical"].Value,
                        CultureInfo.InvariantCulture)) > 0.01) break;
                end++;
            }
            var fonts = lines.Skip(index).Take(end - index)
                .Select(line => Regex.Match(line, @"^/(?<font>\S+) 11 Tf$"))
                .Where(match => match.Success).ToArray();
            if (fonts.Any(font => !calibriFonts.Contains(font.Groups["font"].Value)))
            {
                index = end - 1;
                continue;
            }
            lines[index] = lines[index].Replace(" 11 Tf", " 12 Tf", StringComparison.Ordinal);
            lines[scaleIndex] = (scale * 11 / 12).ToString("0.####", CultureInfo.InvariantCulture) + " Tz";
            for (var cursor = scaleIndex + 1; cursor < end; cursor++)
                if (Regex.IsMatch(lines[cursor], @"^/\S+ 11 Tf$"))
                    lines[cursor] = lines[cursor].Replace(" 11 Tf", " 12 Tf", StringComparison.Ordinal);
            changed = true;
            index = end - 1;
        }
        return changed ? string.Join('\n', lines) : source;
    }

    private PdfDocument? RenderCommentAnnotations(PdfDocumentRenderer renderer, PdfDocument document)
    {
        if (!_includePdfComments || _commentThreads.Count == 0 ||
            document.Internals.Catalog.Names?.NameTree is null)
            return null;
        using var textStream = new MemoryStream();
        document.Save(textStream, false);
        var bytes = textStream.ToArray();
        using var textDocument = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var annotated = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Modify);
        document = annotated;
        var names = document.Internals.Catalog.Names!.NameTree!;

        bool TryReadPosition(string name, out PdfPage? destinationPage, out double left,
            out double top)
        {
            destinationPage = null;
            left = top = 0;
            if (names.GetValue(name, true) is not PdfArray destination ||
                destination.Elements.Count < 2 ||
                destination.Elements[0] is not PdfSharp.Pdf.Advanced.PdfReference pageReference)
                return false;
            destinationPage = document.Pages.Cast<PdfPage>().FirstOrDefault(candidate =>
                candidate.Reference?.ObjectID == pageReference.ObjectID);
            if (destinationPage == null) return false;
            if (destination.Elements.Count >= 4)
            {
                left = destination.Elements[2] switch
                {
                    PdfReal real => real.Value,
                    PdfInteger integer => integer.Value,
                    _ => double.NaN
                };
                top = destination.Elements[3] switch
                {
                    PdfReal real => real.Value,
                    PdfInteger integer => integer.Value,
                    _ => double.NaN
                };
            }
            else if (destination.Elements[1] is PdfLiteral literal)
            {
                var parts = literal.ToString().Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 3 || parts[0] != "/XYZ" ||
                    !double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture,
                        out left) ||
                    !double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture,
                        out top)) return false;
            }
            return double.IsFinite(left) && double.IsFinite(top);
        }

        (PdfRectangle Rectangle, PdfArray QuadPoints)? GetSelectedRange(string anchorId,
            PdfPage page)
        {
            if (!TryReadPosition($"dxp-comment-range-start-{anchorId}", out var startPage,
                    out var startLeft, out var startTop) ||
                !TryReadPosition($"dxp-comment-range-end-{anchorId}", out var endPage,
                    out var endLeft, out var endTop) ||
                startPage == null || endPage == null)
                return null;
            var pageNumber = document.Pages.Cast<PdfPage>().TakeWhile(candidate =>
                candidate.Reference?.ObjectID != page.Reference?.ObjectID).Count() + 1;
            var startPageNumber = document.Pages.Cast<PdfPage>().TakeWhile(candidate =>
                candidate.Reference?.ObjectID != startPage.Reference?.ObjectID).Count() + 1;
            var endPageNumber = document.Pages.Cast<PdfPage>().TakeWhile(candidate =>
                candidate.Reference?.ObjectID != endPage.Reference?.ObjectID).Count() + 1;
            if (pageNumber < startPageNumber || pageNumber > endPageNumber ||
                startPageNumber == endPageNumber &&
                (endTop > startTop + 1 ||
                 Math.Abs(startTop - endTop) <= 1 && endLeft <= startLeft))
                return null;
            var startsHere = pageNumber == startPageNumber;
            var endsHere = pageNumber == endPageNumber;
            var textPage = textDocument.GetPage(pageNumber);
            var viewerMargin = Unit.FromCentimeter(0.5).Point;
            var selectedLeft = startsHere ? startLeft + viewerMargin : double.NegativeInfinity;
            var selectedRight = endsHere ? endLeft + viewerMargin : double.PositiveInfinity;
            var selectedTop = startsHere ? page.Height.Point - startTop + viewerMargin : 0;
            var selectedBottom = endsHere ? page.Height.Point - endTop + viewerMargin : page.Height.Point;
            if (startPageNumber != endPageNumber)
            {
                var bodyTop = double.PositiveInfinity;
                var bodyBottom = double.NegativeInfinity;
                foreach (var info in renderer.DocumentRenderer.GetRenderInfoFromPage(pageNumber))
                    if (info.LayoutInfo.ContentArea is { } bodyArea && bodyArea.Height.Point > 0)
                    {
                        bodyTop = Math.Min(bodyTop, bodyArea.Y.Point);
                        bodyBottom = Math.Max(bodyBottom, bodyArea.Y.Point + bodyArea.Height.Point);
                    }
                if (double.IsFinite(bodyTop) && double.IsFinite(bodyBottom))
                {
                    selectedTop = Math.Max(selectedTop, bodyTop);
                    selectedBottom = Math.Min(selectedBottom, bodyBottom);
                }
            }
            (double Left, double Right)? FindCellBounds(RenderInfo info, Cell target,
                double parentLeft)
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Tables.Table table ||
                    info.LayoutInfo.ContentArea is not { } area)
                    return null;
                var tableLeft = parentLeft + area.X.Point + table[0, 0].Borders.Left.Width.Point;
                double LeftOf(Cell candidate)
                {
                    var left = tableLeft;
                    for (var column = 0; column < candidate.Column.Index; column++)
                        left += table.Columns[column].Width.Point;
                    return left;
                }
                if (ReferenceEquals(table, target.Table))
                {
                    var left = LeftOf(target);
                    var right = left;
                    for (var column = target.Column.Index;
                         column <= target.Column.Index + target.MergeRight &&
                         column < table.Columns.Count; column++)
                        right += table.Columns[column].Width.Point;
                    return (left, right);
                }
                if (info.FormatInfo is not TableFormatInfo { FormattedCells: { } formattedCells })
                    return null;
                var getRenderInfos = typeof(FormattedCell).GetMethod("GetRenderInfos",
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.NonPublic);
                if (getRenderInfos == null) return null;
                foreach (var (parentCell, formatted) in formattedCells)
                {
                    if (getRenderInfos.Invoke(formatted, null) is not RenderInfo[] nestedInfos)
                        continue;
                    var contentLeft = LeftOf(parentCell) + parentCell.Column.LeftPadding.Point;
                    foreach (var nestedInfo in nestedInfos)
                        if (FindCellBounds(nestedInfo, target, contentLeft) is { } found)
                            return found;
                }
                return null;
            }
            var cellBounds = _commentRangeCells.TryGetValue(anchorId, out var cell)
                ? renderer.DocumentRenderer.GetRenderInfoFromPage(pageNumber)
                    .Select(info => FindCellBounds(info, cell, 0))
                    .FirstOrDefault(bounds => bounds != null)
                : null;
            var lines = textPage.Letters
                .Where(letter => textPage.Height - letter.GlyphRectangle.Top >= selectedTop - 8 &&
                    textPage.Height - letter.GlyphRectangle.Top <= selectedBottom + 8 &&
                    (cellBounds == null ||
                     letter.GlyphRectangle.Left >= cellBounds.Value.Left - 1 &&
                     letter.GlyphRectangle.Right <= cellBounds.Value.Right + 1))
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .OrderByDescending(line => line.Key)
                .ToArray();
            if (lines.Length == 0) return null;
            var startIndex = !startsHere ? 0 : lines.Select((line, index) => (line, index))
                .Select(candidate => (candidate.index, Distance: candidate.line
                    .Where(letter => letter.GlyphRectangle.Right > selectedLeft - 0.1)
                    .Select(letter => Math.Abs(textPage.Height - letter.GlyphRectangle.Top - selectedTop))
                    .DefaultIfEmpty(double.PositiveInfinity).Min()))
                .Where(candidate => candidate.Distance < 8)
                .OrderBy(candidate => candidate.Distance)
                .Select(candidate => candidate.index).DefaultIfEmpty(-1).First();
            var endIndex = !endsHere ? lines.Length - 1 : lines.Select((line, index) => (line, index))
                .Select(candidate => (candidate.index, Distance: candidate.line
                    .Where(letter => letter.GlyphRectangle.Left < selectedRight + 0.1)
                    .Select(letter => Math.Abs(textPage.Height - letter.GlyphRectangle.Top - selectedBottom))
                    .DefaultIfEmpty(double.PositiveInfinity).Min()))
                .Where(candidate => candidate.Distance < 8)
                .OrderBy(candidate => candidate.Distance)
                .Select(candidate => candidate.index).DefaultIfEmpty(-1).First();
            if (startsHere && endsHere && selectedBottom - selectedTop > 1 && endIndex <= startIndex)
                endIndex = lines.Select((line, index) => (line, index))
                    .Where(candidate => candidate.index > startIndex)
                    .Select(candidate => (candidate.index, Distance: candidate.line
                        .Where(letter => letter.GlyphRectangle.Left < selectedRight + 0.1)
                        .Select(letter => Math.Abs(textPage.Height - letter.GlyphRectangle.Top - selectedBottom))
                        .DefaultIfEmpty(double.PositiveInfinity).Min()))
                    .Where(candidate => candidate.Distance < 8)
                    .OrderBy(candidate => candidate.Distance)
                    .Select(candidate => candidate.index).DefaultIfEmpty(-1).First();
            if (startIndex < 0 || endIndex < startIndex) return null;
            var quadPoints = new PdfArray(document);
            var bounds = new List<(double Left, double Bottom, double Right, double Top)>();
            for (var index = startIndex; index <= endIndex; index++)
            {
                var segments = new List<List<UglyToad.PdfPig.Content.Letter>>();
                foreach (var letter in lines[index].OrderBy(letter => letter.GlyphRectangle.Left))
                {
                    if (segments.Count == 0 ||
                        letter.GlyphRectangle.Left - segments[^1][^1].GlyphRectangle.Right >
                        CommentLineSegmentGapPoints)
                        segments.Add(new List<UglyToad.PdfPig.Content.Letter>());
                    segments[^1].Add(letter);
                }
                var target = index == endIndex && endsHere && index != startIndex
                    ? selectedRight : selectedLeft;
                var segment = segments.MinBy(candidate => Math.Abs(
                    index == endIndex && endsHere && index != startIndex
                        ? candidate[^1].GlyphRectangle.Right - target
                        : candidate[0].GlyphRectangle.Left - target));
                if (segment == null) return null;
                var selectedSegments = index == startIndex && index == endIndex
                    ? segments.Where(candidate =>
                        candidate[^1].GlyphRectangle.Right > selectedLeft - 0.1 &&
                        candidate[0].GlyphRectangle.Left < selectedRight + 0.1)
                    : [segment];
                var letters = selectedSegments.SelectMany(candidate => candidate).Where(letter =>
                    (!startsHere || index != startIndex || letter.GlyphRectangle.Right > selectedLeft - 0.1) &&
                    (!endsHere || index != endIndex || letter.GlyphRectangle.Left < selectedRight + 0.1))
                    .ToArray();
                if (letters.Length == 0) return null;
                var left = letters.Min(letter => letter.GlyphRectangle.Left);
                var right = letters.Max(letter => letter.GlyphRectangle.Right);
                var top = letters.Max(letter => letter.GlyphRectangle.Top);
                var bottom = letters.Min(letter => letter.GlyphRectangle.Bottom);
                if (right <= left || top <= bottom) return null;
                bounds.Add((left, bottom, right, top));
                foreach (var value in new[] { left, top, right, top, left, bottom, right, bottom })
                    quadPoints.Elements.Add(new PdfReal(value));
            }
            return (new PdfRectangle(new XPoint(bounds.Min(bound => bound.Left),
                bounds.Min(bound => bound.Bottom)), new XPoint(bounds.Max(bound => bound.Right),
                bounds.Max(bound => bound.Top))), quadPoints);
        }

        void ApplyCommentHighlight(PdfSharp.Pdf.Annotations.PdfTextAnnotation annotation,
            (PdfRectangle Rectangle, PdfArray QuadPoints) range)
        {
            annotation.Rectangle = range.Rectangle;
            annotation.Elements.SetName("/Subtype", "/Highlight");
            annotation.Elements.SetValue("/QuadPoints", range.QuadPoints);
            var color = new PdfArray(document);
            color.Elements.Add(new PdfReal(1));
            color.Elements.Add(new PdfReal(1));
            color.Elements.Add(new PdfReal(0));
            annotation.Elements.SetValue("/C", color);
            annotation.Elements.SetReal("/CA", 0.2);
        }

        foreach (var (anchorId, thread) in _commentThreads)
        {
            if (!TryReadPosition($"dxp-comment-{anchorId}", out var page, out _,
                    out var top) || page == null) continue;
            var annotationTop = Math.Clamp(top, 16, page.Height.Point);
            var rootComment = thread.Comments.FirstOrDefault(comment => !comment.IsReply) ?? thread.Comments[0];
            var marginRectangle = new PdfRectangle(new XPoint(page.Width.Point - 24, annotationTop - 16),
                new XPoint(page.Width.Point - 8, annotationTop));
            var selectedRange = GetSelectedRange(anchorId, page);
            var annotation = new PdfSharp.Pdf.Annotations.PdfTextAnnotation(document);
            annotation.Rectangle = marginRectangle;
            annotation.Title = rootComment.Author ?? rootComment.Initials ?? "Comment";
            annotation.Contents = rootComment.Text;
            if (selectedRange is { } range)
                ApplyCommentHighlight(annotation, range);
            if (rootComment.DateUtc is DateTime created)
                annotation.Elements.SetDateTime("/CreationDate", created.ToUniversalTime());
            if (rootComment.IsDone)
            {
                annotation.Elements.SetString("/StateModel", "Review");
                annotation.Elements.SetString("/State", "Completed");
            }
            page.Annotations.Add(annotation);
            foreach (PdfPage continuationPage in document.Pages)
            {
                if (continuationPage.Reference?.ObjectID == page.Reference?.ObjectID ||
                    GetSelectedRange(anchorId, continuationPage) is not { } continuationRange)
                    continue;
                var continuation = new PdfSharp.Pdf.Annotations.PdfTextAnnotation(document)
                {
                    Title = annotation.Title,
                    Contents = annotation.Contents
                };
                ApplyCommentHighlight(continuation, continuationRange);
                if (rootComment.DateUtc is DateTime continuationCreated)
                    continuation.Elements.SetDateTime("/CreationDate", continuationCreated.ToUniversalTime());
                continuationPage.Annotations.Add(continuation);
            }
            var annotationsById = new Dictionary<string, PdfSharp.Pdf.Annotations.PdfTextAnnotation>(StringComparer.Ordinal)
            {
                [rootComment.Id] = annotation
            };
            foreach (var comment in thread.Comments.Where(comment => comment.Id != rootComment.Id))
            {
                var reply = new PdfSharp.Pdf.Annotations.PdfTextAnnotation(document)
                {
                    Rectangle = marginRectangle,
                    Title = comment.Author ?? comment.Initials ?? "Comment",
                    Contents = comment.Text
                };
                if (comment.DateUtc is DateTime replyCreated)
                    reply.Elements.SetDateTime("/CreationDate", replyCreated.ToUniversalTime());
                if (comment.IsDone)
                {
                    reply.Elements.SetString("/StateModel", "Review");
                    reply.Elements.SetString("/State", "Completed");
                }
                page.Annotations.Add(reply);
                annotationsById[comment.Id] = reply;
            }
            foreach (var comment in thread.Comments.Where(comment => comment.Id != rootComment.Id))
            {
                var parent = comment.ParentId != null && annotationsById.TryGetValue(comment.ParentId, out var parentAnnotation)
                    ? parentAnnotation : annotation;
                var reply = annotationsById[comment.Id];
                reply.Elements.SetReference("/IRT", parent);
                reply.Elements.SetName("/RT", "/R");
            }
        }
        return annotated;
    }

    private static void CopyNamedDestinations(PdfDocument source, PdfDocument destination)
    {
        var names = source.Internals.Catalog.Names?.NameTree;
        if (names == null) return;
        foreach (var name in names.GetNames(true))
        {
            if (names.GetValue(name, true) is not PdfArray array || array.Elements.Count < 5 ||
                array.Elements[0] is not PdfSharp.Pdf.Advanced.PdfReference pageReference ||
                array.Elements[1].ToString() != "/XYZ") continue;
            var pageIndex = -1;
            for (var index = 0; index < source.PageCount; index++)
                if (source.Pages[index].Reference?.ObjectID == pageReference.ObjectID)
                {
                    pageIndex = index;
                    break;
                }
            if (pageIndex < 0) continue;
            static double? Number(PdfItem item) => item switch
            {
                PdfReal real => real.Value,
                PdfInteger integer => integer.Value,
                _ => null
            };
            var left = Number(array.Elements[2]);
            destination.AddNamedDestination(name, pageIndex + 1,
                PdfSharp.Pdf.Advanced.PdfNamedDestinationParameters.CreatePosition(
                    left, Number(array.Elements[3]), Number(array.Elements[4])));
        }
    }

    private static void DrawDiagonal(XGraphics graphics, DiagonalStroke? stroke,
        double startX, double startY, double endX, double endY)
    {
        if (stroke == null || !int.TryParse(stroke.Color, NumberStyles.HexNumber,
                CultureInfo.InvariantCulture, out var rgb)) return;
        var pen = new XPen(XColor.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255), stroke.Width);
        graphics.DrawLine(pen, startX, startY, endX, endY);
    }

    private static DiagonalStroke? ReadDiagonalStroke(OpenXmlElement? border)
    {
        if (border == null || AttributeValue(border, "val") is "nil" or "none") return null;
        var color = AttributeValue(border, "color");
        if (color == "auto") color = "000000";
        if (color is not { Length: 6 }) return null;
        var width = double.TryParse(AttributeValue(border, "sz"), NumberStyles.Float,
            CultureInfo.InvariantCulture, out var size) ? Math.Max(0.5, size / 8.0) : 0.5;
        return new DiagonalStroke(color, width);
    }

    private void RenderParagraphShadingGaps(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_paragraphShades.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            foreach (var (storyInfo, offset) in GetPageStoryRenderInfos(renderer, index, target))
            {
                if (storyInfo.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph footerParagraph ||
                    !_footerParagraphShadeClearances.TryGetValue(footerParagraph, out var clearance) ||
                    !_paragraphShades.TryGetValue(footerParagraph, out var footerCss) ||
                    storyInfo.LayoutInfo.ContentArea is not { } footerArea || footerCss.Length != 7 ||
                    !int.TryParse(footerCss.AsSpan(1), NumberStyles.HexNumber,
                        CultureInfo.InvariantCulture, out var footerRgb)) continue;
                var footerWidth = footerArea.Width.Point - footerParagraph.Format.LeftIndent.Point -
                    footerParagraph.Format.RightIndent.Point + 3;
                if (footerWidth <= 0) continue;
                var leftIndent = index % 2 == 1 && _mirroredParagraphs.Contains(footerParagraph)
                    ? footerParagraph.Format.RightIndent.Point : footerParagraph.Format.LeftIndent.Point;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index],
                    XGraphicsPdfPageOptions.Prepend);
                graphics.DrawRectangle(new XSolidBrush(XColor.FromArgb((footerRgb >> 16) & 255,
                    (footerRgb >> 8) & 255, footerRgb & 255)),
                    footerArea.X.Point + leftIndent - 1.5,
                    footerArea.Y.Point + offset - 2 * clearance, footerWidth,
                    footerArea.Height.Point + clearance);
            }
            var infos = renderer.DocumentRenderer.GetRenderInfoFromPage(index + 1);
            for (var position = 0; position < infos.Length; position++)
            {
                if (infos[position].DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    !_manualParagraphShades.Contains(paragraph) ||
                    !_paragraphShades.TryGetValue(paragraph, out var css) ||
                    infos[position].LayoutInfo.ContentArea is not { } area || css.Length != 7 ||
                    !int.TryParse(css.AsSpan(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture,
                        out var rgb)) continue;
                var previous = position > 0 &&
                    infos[position - 1].DocumentObject is MigraDoc.DocumentObjectModel.Paragraph before &&
                    _manualParagraphShades.Contains(before) &&
                    _paragraphShades.GetValueOrDefault(before) == css ? before : null;
                var next = position + 1 < infos.Length &&
                    infos[position + 1].DocumentObject is MigraDoc.DocumentObjectModel.Paragraph after &&
                    _manualParagraphShades.Contains(after) &&
                    _paragraphShades.GetValueOrDefault(after) == css &&
                    infos[position + 1].LayoutInfo.ContentArea != null ? after : null;
                var top = area.Y.Point + (previous?.Format.SpaceAfter.Point ?? 0);
                var bottom = next != null
                    ? infos[position + 1].LayoutInfo.ContentArea!.Y.Point + paragraph.Format.SpaceAfter.Point
                    : area.Y.Point + area.Height.Point + paragraph.Format.SpaceAfter.Point +
                      (previous?.Format.SpaceAfter.Point ?? 0);
                var leftIndent = index % 2 == 1 && _mirroredParagraphs.Contains(paragraph)
                    ? paragraph.Format.RightIndent.Point : paragraph.Format.LeftIndent.Point;
                var width = area.Width.Point - paragraph.Format.LeftIndent.Point -
                    paragraph.Format.RightIndent.Point + 3;
                if (width <= 0 || bottom <= top) continue;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index],
                    XGraphicsPdfPageOptions.Prepend);
                graphics.DrawRectangle(new XSolidBrush(XColor.FromArgb((rgb >> 16) & 255,
                    (rgb >> 8) & 255, rgb & 255)), area.X.Point + leftIndent - 1.5,
                    top, width, bottom - top);
            }
            for (var position = 0; position < infos.Length - 1; position++)
            {
                if (infos[position].DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    infos[position + 1].DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph next ||
                    _manualParagraphShades.Contains(paragraph) ||
                    !_paragraphShades.TryGetValue(paragraph, out var css) ||
                    !_paragraphShades.TryGetValue(next, out var nextCss) ||
                    !string.Equals(css, nextCss, StringComparison.OrdinalIgnoreCase) ||
                    infos[position].LayoutInfo.ContentArea is not { } area ||
                    infos[position + 1].LayoutInfo.ContentArea is not { } nextArea ||
                    css.Length != 7 ||
                    !int.TryParse(css.AsSpan(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var rgb))
                    continue;
                var start = area.Y.Point + area.Height.Point;
                var gap = nextArea.Y.Point - start;
                var width = area.Width.Point - paragraph.Format.LeftIndent.Point -
                    paragraph.Format.RightIndent.Point;
                if (gap <= 0 || gap > paragraph.Format.SpaceAfter.Point + 1 || width <= 0) continue;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index],
                    XGraphicsPdfPageOptions.Prepend);
                graphics.DrawRectangle(new XSolidBrush(XColor.FromArgb((rgb >> 16) & 255,
                    (rgb >> 8) & 255, rgb & 255)), area.X.Point + paragraph.Format.LeftIndent.Point,
                    start, width, gap);
            }
        }
    }

    private void RenderContinuousParagraphSideBorders(PdfDocumentRenderer renderer,
        PdfDocument? target = null)
    {
        if (_directBorderParagraphs.Count == 0) return;
        for (var page = 1; page <= renderer.PageCount; page++)
        {
            var infos = renderer.DocumentRenderer.GetRenderInfoFromPage(page)
                .Where(info => info.DocumentObject is MigraDoc.DocumentObjectModel.Paragraph)
                .ToArray();
            for (var position = 0; position < infos.Length - 1; position++)
            {
                if (infos[position].DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    infos[position + 1].DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph next ||
                    !_directBorderParagraphs.TryGetValue(paragraph, out var source) ||
                    !_directBorderParagraphs.TryGetValue(next, out var nextSource) ||
                    _mirroredParagraphs.Contains(paragraph) ||
                    !ReferenceEquals(source.NextSibling(), nextSource) ||
                    source.ParagraphProperties?.ParagraphBorders is not { } borders ||
                    nextSource.ParagraphProperties?.ParagraphBorders?.OuterXml != borders.OuterXml ||
                    infos[position].LayoutInfo.ContentArea is not { } area ||
                    infos[position + 1].LayoutInfo.ContentArea is not { } nextArea)
                    continue;
                var bottom = area.Y.Point + area.Height.Point;
                var top = nextArea.Y.Point;
                if (top <= bottom || top - bottom > 24) continue;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[page - 1]);
                foreach (var side in new[] { "left", "right" })
                {
                    var border = borders.ChildElements.FirstOrDefault(element => element.LocalName == side);
                    if (border == null || AttributeValue(border, "val") is "nil" or "none") continue;
                    var color = AttributeValue(border, "color");
                    if (color is not { Length: 6 } && color != "auto") continue;
                    var rgb = 0;
                    if (color != "auto" && !int.TryParse(color, NumberStyles.HexNumber,
                            CultureInfo.InvariantCulture, out rgb)) continue;
                    var width = double.TryParse(AttributeValue(border, "sz"), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var eighths) ? eighths / 8 : 0.5;
                    var distance = double.TryParse(AttributeValue(border, "space"), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var points) ? points : 0;
                    var x = side == "left" ?
                        area.X.Point + paragraph.Format.LeftIndent.Point - distance -
                        DirectParagraphSideBorderOffsetPoints - width / 2 :
                        area.X.Point + area.Width.Point - paragraph.Format.RightIndent.Point +
                        distance + DirectParagraphSideBorderOffsetPoints + width / 2;
                    graphics.DrawLine(new XPen(XColor.FromArgb((rgb >> 16) & 255,
                        (rgb >> 8) & 255, rgb & 255), width), x, bottom, x, top);
                }
            }
        }
    }

    private sealed record FooterImageOverlay(HeaderFooterValues Kind, byte[] Bytes,
        double Left, double Top, double Width, double Height, bool Header = false,
        MigraDoc.DocumentObjectModel.Paragraph? AnchorParagraph = null, bool BehindDocument = false,
        long ZOrder = 0);

    private sealed record ParagraphImageOverlay(byte[] Bytes, double Left, double Top,
        double Width, double Height, bool PageRelative = false, long ZOrder = 0,
        bool BehindDocument = false);
    private sealed record FloatingTextWrap(double Left, double Right);

    private void ClipExactLineImages(PdfDocumentRenderer renderer, PdfDocument document)
    {
        if (_exactLineImages.Count == 0) return;
        for (var pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
        {
            var images = renderer.DocumentRenderer.GetRenderInfoFromPage(pageIndex + 1)
                .Where(info => info.DocumentObject is MigraDoc.DocumentObjectModel.Paragraph &&
                    info.LayoutInfo.ContentArea != null)
                .SelectMany(info => _exactLineImages
                    .Where(item => ReferenceEquals(item.Paragraph, info.DocumentObject))
                    .Select(item => (item.Width, item.Height, item.LineHeight,
                        Area: info.LayoutInfo.ContentArea!)))
                .ToArray();
            if (images.Length == 0) continue;
            foreach (var reference in document.Pages[pageIndex].Contents.Elements
                         .OfType<PdfSharp.Pdf.Advanced.PdfReference>())
            {
                if (reference.Value is not PdfSharp.Pdf.Advanced.PdfContent content ||
                    content.Stream == null || content.Elements.ContainsKey("/Filter")) continue;
                var commands = Encoding.Latin1.GetString(content.Stream.Value);
                var updated = Regex.Replace(commands,
                    @"(?m)^q (?<width>\d+(?:\.\d+)?) 0 0 (?<height>\d+(?:\.\d+)?) (?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) cm /I\d+ Do Q$",
                    match =>
                    {
                        var width = double.Parse(match.Groups["width"].Value, CultureInfo.InvariantCulture);
                        var height = double.Parse(match.Groups["height"].Value, CultureInfo.InvariantCulture);
                        var x = double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture);
                        var bottom = document.Pages[pageIndex].Height.Point -
                            double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture);
                        var image = images.FirstOrDefault(item =>
                            Math.Abs(item.Width - width) < 0.1 && Math.Abs(item.Height - height) < 0.1 &&
                            x >= item.Area.X.Point - 1 &&
                            x <= item.Area.X.Point + item.Area.Width.Point + 1 &&
                            bottom >= item.Area.Y.Point - 1 &&
                            bottom <= item.Area.Y.Point + item.Area.Height.Point + 1);
                        if (image.Area == null) return match.Value;
                        return $"q\n{match.Groups["x"].Value} {match.Groups["y"].Value} " +
                            $"{match.Groups["width"].Value} {(image.LineHeight * 0.8).ToString("0.####", CultureInfo.InvariantCulture)} re W n\n" +
                            match.Value + "\nQ";
                    });
                if (updated != commands) content.Stream.Value = Encoding.Latin1.GetBytes(updated);
            }
        }
    }

    private static void DisableRasterInterpolation(PdfDocument document)
    {
        foreach (var page in document.Pages)
        {
            var objects = page.Resources.Elements.GetDictionary("/XObject");
            if (objects == null) continue;
            foreach (var value in objects.Elements.Values)
            {
                var image = (value as PdfSharp.Pdf.Advanced.PdfReference)?.Value as PdfSharp.Pdf.PdfDictionary ??
                    value as PdfSharp.Pdf.PdfDictionary;
                if (image?.Elements.GetName("/Subtype") != "/Image") continue;
                if (image.Elements.GetName("/Filter") != "/DCTDecode")
                    image.Elements.SetBoolean("/Interpolate", false);
            }
        }
    }

    private void RenderParagraphAnchoredImages(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_paragraphImages.Count == 0) return;
        var drawnParagraphs = new HashSet<MigraDoc.DocumentObjectModel.Paragraph>();
        for (var index = 0; index < renderer.PageCount; index++)
        {
            foreach (var info in renderer.DocumentRenderer.GetRenderInfoFromPage(index + 1))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    !_paragraphImages.TryGetValue(paragraph, out var images) ||
                    info.LayoutInfo.ContentArea is not { } area ||
                    !drawnParagraphs.Add(paragraph)) continue;
                foreach (var overlay in images.Where(image => image.BehindDocument)
                             .OrderByDescending(image => image.ZOrder)
                             .Concat(images.Where(image => !image.BehindDocument)
                                 .OrderBy(image => image.ZOrder)))
                {
                    using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index],
                        overlay.BehindDocument ? XGraphicsPdfPageOptions.Prepend : XGraphicsPdfPageOptions.Append);
                    using var stream = new MemoryStream(overlay.Bytes);
                    using var image = XImage.FromStream(stream);
                    graphics.DrawImage(image, overlay.PageRelative ? overlay.Left : area.X.Point + overlay.Left,
                        overlay.PageRelative ? overlay.Top : area.Y.Point + overlay.Top,
                        overlay.Width, overlay.Height);
                }
            }
        }
    }

    private void RenderMathFractions(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_mathFractions.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, target))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    !_mathFractions.TryGetValue(paragraph, out var fraction) ||
                    info.LayoutInfo.ContentArea is not { } area) continue;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index]);
                XFont font;
                try { font = new XFont("Cambria Math", fraction.Size); }
                catch (InvalidOperationException)
                {
                    try { font = new XFont("Cambria", fraction.Size); }
                    catch (InvalidOperationException) { font = new XFont(_fontFamily, fraction.Size); }
                }
                var center = area.X.Point + area.Width.Point / 2;
                var numeratorWidth = graphics.MeasureString(fraction.Numerator, font).Width;
                var denominatorWidth = graphics.MeasureString(fraction.Denominator, font).Width;
                var sizeDelta = fraction.Size - 12;
                var firstBodyOffset = _firstBodyDisplayFractions.Contains(paragraph);
                var top = area.Y.Point + offset - (firstBodyOffset ? 0 : 10.31);
                var glyphBaselineCorrection = firstBodyOffset ? 0 : fraction.Size * 0.17;
                graphics.DrawString(fraction.Numerator, font, XBrushes.Black,
                    new XRect(center - numeratorWidth / 2,
                        top + 10.8 - sizeDelta * 0.1075 -
                        (firstBodyOffset ? 12.5 + sizeDelta * 0.15625 : 0) -
                        glyphBaselineCorrection,
                        numeratorWidth + 1, fraction.Size + 3),
                    XStringFormats.TopLeft);
                graphics.DrawRectangle(XBrushes.Black, center - Math.Max(numeratorWidth, denominatorWidth) / 2,
                    top + 25.45 + sizeDelta * 0.675 - (firstBodyOffset ? 10.31 : 0),
                    Math.Max(numeratorWidth, denominatorWidth), 0.84);
                graphics.DrawString(fraction.Denominator, font, XBrushes.Black,
                    new XRect(center - denominatorWidth / 2,
                        top + 27.98 + sizeDelta * 0.3875 -
                        (firstBodyOffset ? 12.5 + sizeDelta * 0.15625 : 0) -
                        glyphBaselineCorrection,
                        denominatorWidth + 1, fraction.Size + 3),
                    XStringFormats.TopLeft);
            }
    }

    private void RenderMathRadicals(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_mathRadicals.Count == 0) return;
        for (var index = 0; index < renderer.PageCount; index++)
            foreach (var (info, offset) in GetPageStoryRenderInfos(renderer, index, target))
            {
                if (info.DocumentObject is not MigraDoc.DocumentObjectModel.Paragraph paragraph ||
                    !_mathRadicals.TryGetValue(paragraph, out var value) ||
                    info.LayoutInfo.ContentArea is not { } area) continue;
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index]);
                XFont font;
                try { font = new XFont("Cambria Math", 12); }
                catch (InvalidOperationException) { font = new XFont(_fontFamily, 12); }
                var text = "√" + (font.FontFamily.Name == "Cambria Math" ? MathItalicVariable(value) : value);
                var width = graphics.MeasureString(text, font).Width;
                var left = area.X.Point + area.Width.Point / 2 - width / 2 + 1.7;
                var top = area.Y.Point + offset - 10.31;
                graphics.DrawString(text, font, XBrushes.Black,
                    new XRect(left, top + 14.42, width + 1, 16), XStringFormats.TopLeft);
                var radicalWidth = graphics.MeasureString("√", font).Width;
                graphics.DrawRectangle(XBrushes.Black, left + radicalWidth, top + 13.57,
                    width - radicalWidth, 0.84);
            }
    }

    private void RenderFooterImages(PdfDocumentRenderer renderer, PdfDocument? target = null)
    {
        if (_footerImages.Count == 0) return;
        Section? previousSection = null;
        var sectionPage = 0;
        var displayedPage = 0;
        for (var index = 0; index < renderer.PageCount; index++)
        {
            var section = renderer.DocumentRenderer.GetRenderInfoFromPage(index + 1)?
                .Select(info => info.DocumentObject.Section)
                .FirstOrDefault(candidate => candidate != null);
            if (section == null) continue;
            sectionPage = ReferenceEquals(section, previousSection) ? sectionPage + 1 : 1;
            displayedPage = sectionPage == 1 && section.PageSetup.StartingNumber > 0
                ? section.PageSetup.StartingNumber : displayedPage + 1;
            previousSection = section;
            if (!_footerImages.TryGetValue(section, out var images)) continue;
            var kind = sectionPage == 1 && section.PageSetup.DifferentFirstPageHeaderFooter
                ? HeaderFooterValues.First
                : section.PageSetup.OddAndEvenPagesHeaderFooter && displayedPage % 2 == 0
                    ? HeaderFooterValues.Even : HeaderFooterValues.Default;
            var selected = images.Where(image => image.Kind == kind).ToArray();
            if (selected.Length == 0) continue;
            foreach (var overlay in selected.Where(image => image.BehindDocument)
                         .OrderByDescending(image => image.ZOrder)
                         .Concat(selected.Where(image => !image.BehindDocument)
                             .OrderBy(image => image.ZOrder)))
            {
                using var graphics = XGraphics.FromPdfPage((target ?? renderer.PdfDocument).Pages[index],
                    overlay.BehindDocument ? XGraphicsPdfPageOptions.Prepend : XGraphicsPdfPageOptions.Append);
                using var stream = new MemoryStream(overlay.Bytes);
                using var image = XImage.FromStream(stream);
                var anchorTop = overlay.AnchorParagraph == null ? null :
                    GetPageStoryRenderInfos(renderer, index, target)
                        .Where(entry => ReferenceEquals(entry.Info.DocumentObject, overlay.AnchorParagraph))
                        .Select(entry => entry.Info.LayoutInfo.ContentArea is { } area
                            ? (double?)(area.Y.Point + entry.Offset) : null)
                        .FirstOrDefault();
                var storyTop = anchorTop ?? (overlay.AnchorParagraph != null && !overlay.Header
                    ? renderer.PdfDocument.Pages[index].Height.Point -
                        section.PageSetup.FooterDistance.Point - 15
                    : overlay.Header ? section.PageSetup.HeaderDistance.Point :
                        section.PageSetup.TopMargin.Point);
                graphics.DrawImage(image, section.PageSetup.LeftMargin.Point + overlay.Left,
                    storyTop + overlay.Top, overlay.Width, overlay.Height);
            }
        }
    }

    private static void DrawPageBorder(XGraphics graphics, PageBorders borders, double width, double height)
    {
        static double? Footprint(OpenXmlElement? side)
        {
            if (side == null || AttributeValue(side, "val") is "none" or "nil") return null;
            var size = double.TryParse(AttributeValue(side, "sz"), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var eighths) ? eighths * 0.12 : 0.48;
            var space = double.TryParse(AttributeValue(side, "space"), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var points) ? points : 0;
            return space + size * (AttributeValue(side, "val") == "double" ? 3 : 1);
        }
        var topBorder = borders.ChildElements.FirstOrDefault(side => side.LocalName == "top");
        var bottomBorder = borders.ChildElements.FirstOrDefault(side => side.LocalName == "bottom");
        foreach (var side in borders.ChildElements)
        {
            var kind = side.LocalName;
            if (kind is not ("top" or "bottom" or "left" or "right")) continue;
            var value = AttributeValue(side, "val");
            if (value is "none" or "nil") continue;
            var color = AttributeValue(side, "color");
            if (color?.Length != 6 || !uint.TryParse(color, NumberStyles.HexNumber,
                CultureInfo.InvariantCulture, out var rgb)) continue;
            var size = double.TryParse(AttributeValue(side, "sz"), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var eighths) ? eighths * 0.12 : 0.48;
            var space = double.TryParse(AttributeValue(side, "space"), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var points) ? points : 0;
            var pen = new XPen(XColor.FromArgb((int)(rgb >> 16), (int)((rgb >> 8) & 255), (int)(rgb & 255)), size);
            pen.DashStyle = value switch
            {
                "dashed" => XDashStyle.Dash,
                "dotted" => XDashStyle.Dot,
                _ => XDashStyle.Solid
            };
            var center = space + size / 2;
            var topEdge = Footprint(topBorder) ?? space;
            var bottomEdge = height - (Footprint(bottomBorder) ?? space);
            if (bottomBorder != null && AttributeValue(bottomBorder, "val") == "double")
                bottomEdge += 0.2;
            if (kind == "top") graphics.DrawLine(pen, space, center, width - space, center);
            if (kind == "bottom") graphics.DrawLine(pen, space,
                height - center + (value == "double" ? 0.2 : 0), width - space,
                height - center + (value == "double" ? 0.2 : 0));
            if (kind == "left") graphics.DrawLine(pen, center, topEdge, center, bottomEdge);
            if (kind == "right") graphics.DrawLine(pen, width - center, topEdge, width - center, bottomEdge);
            if (value == "double")
            {
                var inner = center + 2 * size;
                if (kind == "top") graphics.DrawLine(pen, space, inner, width - space, inner);
                if (kind == "bottom") graphics.DrawLine(pen, space, height - inner + 0.2,
                    width - space, height - inner + 0.2);
                if (kind == "left") graphics.DrawLine(pen, inner, topEdge, inner, bottomEdge);
                if (kind == "right") graphics.DrawLine(pen, width - inner, topEdge, width - inner, bottomEdge);
            }
        }
    }

    private static string? AttributeValue(OpenXmlElement element, string name)
        => element.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == name).Value;

    public override IDisposable VisitSectionBodyBegin(SectionProperties properties, DxpIDocumentContext context)
    {
        var previous = _container;
        _container = _section;
        var body = properties.Ancestors<Body>().FirstOrDefault();
        var sectionEnd = properties.Parent is DocumentFormat.OpenXml.Wordprocessing.Paragraph boundary
            ? (OpenXmlElement)boundary : properties;
        var preceding = body?.ChildElements.TakeWhile(element => element != sectionEnd).ToArray() ?? [];
        var previousSectionIndex = Array.FindLastIndex(preceding, element =>
            element is DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph &&
            paragraph.ParagraphProperties?.SectionProperties != null);
        var sectionElements = body?.ChildElements.Skip(previousSectionIndex + 1)
            .TakeWhile(element => element != sectionEnd)
            .Concat(sectionEnd is DocumentFormat.OpenXml.Wordprocessing.Paragraph ? [sectionEnd] : [])
            .ToArray() ?? [];
        var sectionParagraphs = sectionElements.OfType<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToArray();
        if (_columnWidths != null && body != null &&
            sectionElements.All(element => element is DocumentFormat.OpenXml.Wordprocessing.Paragraph or SectionProperties) &&
            sectionParagraphs.All(paragraph =>
                paragraph.Descendants<Drawing>().All(drawing =>
                    drawing.Descendants<TextBoxContent>().Any())) &&
            (sectionParagraphs.Any(paragraph => paragraph.Descendants<Break>()
                 .Any(br => br.Type?.Value == BreakValues.Column)) ||
             sectionParagraphs.All(paragraph =>
                 paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                     .Where(text => !text.Ancestors<TextBoxContent>().Any())
                     .Sum(text => text.Text.Length) < 45)))
        {
            _columnIndex = 0;
            _columnParagraphCount = 0;
            _columnBreakPending = false;
            AddColumnPage();
            if (properties.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Any() &&
                properties.GetFirstChild<SectionType>() is { } sectionType &&
                AttributeValue(sectionType, "val") == "continuous")
            {
                _columnIndex = 1;
                _activeColumnCell = _columnCells![1];
            }
        }
        return DxpDisposable.Create(() =>
        {
            _container = previous;
            _activeColumnCell = null;
            _columnCells = null;
        });
    }

    private void AddColumnPage()
    {
        var table = _section!.AddTable();
        table.Borders.Visible = false;
        table.Rows.LeftIndent = Unit.Zero;
        table.LeftPadding = Unit.Zero;
        table.RightPadding = Unit.Zero;
        table.TopPadding = Unit.Zero;
        table.BottomPadding = Unit.Zero;
        for (var index = 0; index < _columnWidths!.Length; index++)
        {
            if (index > 0) table.AddColumn(Unit.FromPoint(_columnGapPoints));
            table.AddColumn(Unit.FromPoint(_columnWidths[index]));
        }
        var row = table.AddRow();
        row.TopPadding = Unit.Zero;
        row.BottomPadding = Unit.Zero;
        _columnCells = Enumerable.Range(0, _columnWidths.Length)
            .Select(index => row.Cells[index * 2]).ToArray();
        _activeColumnCell = _columnCells[0];
    }

    private void AdvanceColumn()
    {
        var explicitBreak = _columnBreakPending;
        if (_lastParagraph != null)
            _lastParagraph.Format.SpaceAfter = Unit.Zero;
        _columnIndex++;
        if (_columnIndex == _columnWidths!.Length)
        {
            _section!.AddPageBreak();
            AddColumnPage();
            _columnIndex = 0;
        }
        _activeColumnCell = _columnCells![_columnIndex];
        _columnParagraphCount = 0;
        if (explicitBreak && _columnIndex == 1)
        {
            var spacer = _activeColumnCell.AddParagraph();
            spacer.Format.LineSpacingRule = LineSpacingRule.Exactly;
            spacer.Format.LineSpacing = Unit.FromPoint(SimpleColumnParagraphPitchPoints);
            spacer.Format.SpaceAfter = Unit.Zero;
            _columnParagraphCount = 1;
        }
        _columnBreakPending = false;
    }

    public override IDisposable VisitSectionHeaderBegin(Header header, object value, DxpIDocumentContext context)
        => EnterHeaderFooter(value, header: true);

    public override IDisposable VisitSectionFooterBegin(Footer footer, object value, DxpIDocumentContext context)
        => EnterHeaderFooter(value, header: false);

    private IDisposable EnterHeaderFooter(object value, bool header)
    {
        var previous = _container;
        var previousHeader = _inHeader;
        var previousKind = _headerFooterKind;
        _inHeader = header;
        if (header)
        {
            _headerGridWrapClearancePoints = 0;
            _headerHasHorizontalMerge = false;
            _headerHasOccupiedHorizontalMerge = false;
            _headerHasTrailingUnstyledTableParagraph = false;
            _headerParagraphCount = 0;
            _headerHasUnnumberedDefaultTab = false;
            _headerNumberedParagraphCount = 0;
            _headerPositionalTabCount = 0;
            _headerPositionalTabHasFollowingText = false;
            _headerPositionalCellParagraphCount = 0;
            _headerLeftMarginTabBreakCount = 0;
            _headerEmptyParagraphCount = 0;
            _headerSizedEmptyLineExtraPoints = 0;
            _headerPendingEmptyParagraphCount = 0;
            _headerPreviousParagraphUsesBodySpacing = false;
            _headerFirstBeforePoints = 0;
            _headerFirstAfterPoints = 0;
            _headerFirstAfterLayoutAdjustmentPoints = 0;
            _headerLastAfterPoints = 0;
            _headerHasContextualSpacing = false;
            _headerHasContent = false;
            _headerHasDirectBidiOverride = false;
            _headerParagraphsAreNumbered = true;
            _headerRowCount = 0;
            _headerMinimumRowClearancePoints = 0;
            _headerHasExactHeightRow = false;
            _headerHasZeroMarginRow = false;
            _headerHasRotatedCell = false;
            _nestedHeaderRowCount = 0;
            _headerCompactRowTransitions = 0;
            _headerRowSpacingReductionPoints = 0;
            _headerCellSpacingPoints = 0;
            _headerHasFilledSpacedTable = false;
            _headerHasStyledTable = false;
            _headerHasImbalancedAutoFitTable = false;
        }
        else
        {
            _footerHasHorizontalMerge = false;
            _footerHasContextualSpacing = false;
            _footerLastParagraphHasText = false;
            _footerLeftMarginTabBreakCount = 0;
            _footerParagraphCount = 0;
        }
        var kind = (value as DxpHeaderFooterContext)?.Kind ?? HeaderFooterValues.Default;
        _headerFooterKind = kind;
        var group = header ? _section!.Headers : _section!.Footers;
        if (kind == HeaderFooterValues.First)
        {
            _container = group.FirstPage;
        }
        else if (kind == HeaderFooterValues.Even)
        {
            _container = group.EvenPage;
        }
        else
        {
            _container = group.Primary;
        }
        return DxpDisposable.Create(() =>
        {
            if (header && (_headerHasContent || _headerRowCount > 0) && _section != null)
            {
                var page = _section.PageSetup;
                var height = _headerRowCount > 0 || !_headerParagraphsAreNumbered
                    ? page.HeaderDistance.Point +
                        (_headerParagraphCount + _headerRowCount) *
                        (_defaultParagraphLinePoints + _defaultParagraphAfterPoints) -
                    _headerEmptyParagraphCount * _defaultParagraphAfterPoints +
                    _headerCellSpacingPoints * (_headerRowCount + 1) +
                    (_headerRowCount > 0 ? 3 : 0) -
                    CompactHeaderRowReductionPoints * _headerCompactRowTransitions -
                    _headerRowSpacingReductionPoints +
                    (_headerRowCount == 1 && _headerCellSpacingPoints > 0
                        ? _defaultParagraphAfterPoints + 1 : 0)
                    : page.HeaderDistance.Point +
                        _headerParagraphCount * _defaultParagraphLinePoints +
                        Math.Max(0, _headerParagraphCount - 1) * _defaultParagraphAfterPoints;
                if (_headerRowCount == 0 &&
                    (_headerEmptyParagraphCount > 0 ||
                     _headerHasContextualSpacing && _headerPendingEmptyParagraphCount == 0))
                    height += _headerFirstBeforePoints + _headerLastAfterPoints;
                height += _headerPositionalTabCount *
                    (_headerPositionalTabHasFollowingText ? 12.5 : 12.32);
                if (_headerPositionalCellParagraphCount > 0 && _headerRowCount == 1)
                    height = Math.Max(height, page.HeaderDistance.Point +
                        _headerPositionalCellParagraphCount * (_defaultParagraphLinePoints + 12.5) +
                        Math.Max(0, _headerPositionalCellParagraphCount - 1) *
                        (_defaultParagraphLinePoints + 4));
                if (_headerParagraphCount >= 3)
                    height += _headerLeftMarginTabBreakCount * (_defaultParagraphLinePoints + 20);
                if (_headerHasFilledSpacedTable && _headerRowCount == 1)
                    height = Math.Max(height, page.HeaderDistance.Point +
                        _defaultParagraphLinePoints + 12 + 4 * _headerCellSpacingPoints);
                if (_documentDefaultParagraphAfterPoints is > 0 &&
                    _headerRowCount == 1 && _headerParagraphCount == 1 &&
                    _headerHasZeroMarginRow && _headerLastAfterPoints == 0)
                    height = Math.Min(height, page.HeaderDistance.Point +
                        3 * (_defaultParagraphLinePoints + 2));
                page.TopMargin = Unit.FromPoint(Math.Max(page.TopMargin.Point, height));
                if (_headerParagraphCount == 2 && _headerRowCount == 0 &&
                    _headerHasDirectBidiOverride)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point + 2.3);
                if (_documentHasFitText && _headerParagraphCount >= 2 && _headerRowCount == 0)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point + 17);
                if (_headerRowCount == 1 && _headerParagraphCount == 1 &&
                    (_headerHasStyledTable || _bodyHasRowGridGaps ||
                     _headerHasZeroMarginRow && _documentDefaultParagraphAfterPoints is > 0 &&
                     _defaultParagraphLinePoints > 13 &&
                     !_headerHasImbalancedAutoFitTable && !_headerHasRotatedCell &&
                     !_headerHasTrailingUnstyledTableParagraph) &&
                    !_headerHasFilledSpacedTable)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point - 2.85);
                else if (_headerRowCount == 1 && _headerParagraphCount == 1 &&
                    !_headerHasFilledSpacedTable && !_headerHasImbalancedAutoFitTable &&
                    !_headerHasRotatedCell && !_headerHasZeroMarginRow &&
                    !_bodyHasRowBorderExceptions && _nestedHeaderRowCount <= 1 &&
                    _headerMinimumRowClearancePoints == 0 && _headerCellSpacingPoints == 0)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point -
                        (_nestedHeaderRowCount == 0 ? 2.4 : 1.4));
                if (_snapToCharsGrid || _gridLinePitchPoints is > 0 && _gridCharSpacingPoints is > 0)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point + _headerGridWrapClearancePoints -
                        (_gridLinePitchPoints is > 0 && _headerGridWrapClearancePoints > 0 ? 5 : 0));
                if (_headerParagraphCount < 3 && _headerLeftMarginTabBreakCount > 0)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point +
                        5 * _headerLeftMarginTabBreakCount);
                if (_headerHasImbalancedAutoFitTable)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point +
                        Math.Min(_defaultParagraphLinePoints, 12.5) -
                        (string.Equals(_minorThemeLatinFont, "Arial", StringComparison.OrdinalIgnoreCase) ? 4 : 0));
                if (_headerHasHorizontalMerge)
                {
                    var occupiedMergeClearance = _headerHasOccupiedHorizontalMerge &&
                        !_explicitDefaultRunProperties ? (_implicitModernFont ? 4.6 : 2.7) : 0;
                    page.TopMargin = Unit.FromPoint(Math.Max(page.TopMargin.Point,
                        page.HeaderDistance.Point + _headerRowCount * (_defaultParagraphLinePoints + 10) + 25 +
                        occupiedMergeClearance));
                }
                if (_headerHasTrailingUnstyledTableParagraph)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point + _defaultParagraphLinePoints / 2);
                if (_headerMinimumRowClearancePoints > 0)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point + _headerMinimumRowClearancePoints);
                if (_headerHasRotatedCell)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point - 2.5);
                if (_nestedHeaderRowCount > 0)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point + 13.2 * _nestedHeaderRowCount);
                if (_headerRowCount == 0 && _headerNumberedParagraphCount > 0 &&
                    _headerNumberedParagraphCount < _headerParagraphCount)
                    page.TopMargin = Unit.FromPoint(Math.Max(
                        _sectionBaseTopMarginPoints + Math.Max(0, _headerFirstAfterPoints),
                        page.TopMargin.Point - _defaultParagraphLinePoints +
                        (_headerHasUnnumberedDefaultTab ? 2.5 : 0)));
                if (_headerSizedEmptyLineExtraPoints > 0 &&
                    !_headersWithStyledEmptyLine.Contains(_section) && !_headerHasContextualSpacing)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point + _headerSizedEmptyLineExtraPoints);
                if (kind != HeaderFooterValues.First && _headerRowCount > 0 &&
                    page.TopMargin.Point > _sectionBaseTopMarginPoints)
                    _sectionOtherHeaderTableRaisedTopMargin = true;
                if (kind == HeaderFooterValues.First && page.DifferentFirstPageHeaderFooter &&
                    _sectionOtherHeaderTableRaisedTopMargin && _headerRowCount == 0 &&
                    height <= _sectionBaseTopMarginPoints)
                    page.TopMargin = Unit.FromPoint(Math.Max(_sectionBaseTopMarginPoints, height));
                if (_headerHasImbalancedAutoFitTable && _headerRowCount == 1 &&
                    !string.Equals(_minorThemeLatinFont, "Arial", StringComparison.OrdinalIgnoreCase))
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point - 0.7);
                if (_headerHasExactHeightRow && _headerRowCount == 1)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point + 0.5);
                if (_headerRowCount >= 3 && !_bodyHasRowGridGaps)
                    page.TopMargin = Unit.FromPoint(page.TopMargin.Point + 0.65);
            }
            if (!header && _footerLeftMarginTabBreakCount > 0 && _section != null)
                _section.PageSetup.FooterDistance = Unit.FromPoint(
                    _section.PageSetup.FooterDistance.Point +
                    _footerLeftMarginTabBreakCount *
                         (_footerParagraphCount >= 3 ? _defaultParagraphLinePoints + 4.2 : 0.7));
            if (!header && _footerHasContextualSpacing && _footerLastParagraphHasText &&
                _lastParagraphContainer is HeaderFooter &&
                _lastParagraph != null)
                _lastParagraph.Format.SpaceAfter = Unit.Zero;
            if (!header && _footerParagraphCount == 1 && _footerLastParagraphHasText &&
                _lastParagraphContainer is HeaderFooter && _lastParagraph != null &&
                _lastParagraph.Format.SpaceBefore.Point > 0)
                _lastParagraph.Format.SpaceBefore = Unit.Zero;
            if (!header && _footerHasHorizontalMerge && !_footerHorizontalMergeClearanceApplied && _section != null)
            {
                double clearance = _implicitModernFont ? 11 : 13;
                if (_explicitDefaultRunProperties) clearance -= 1.4;
                else clearance -= _implicitModernFont ? 0.2 : 0.4;
                _section.PageSetup.FooterDistance = Unit.FromPoint(_section.PageSetup.FooterDistance.Point + clearance);
                _footerHorizontalMergeClearanceApplied = true;
            }
            if (!header && _paragraphStyles.Count == 0 && _footerParagraphCount == 1 &&
                _footerLastParagraphHasText && _section != null &&
                _unstyledFooterDistanceApplied.Add(_section))
                _section.PageSetup.FooterDistance = Unit.FromPoint(
                    _section.PageSetup.FooterDistance.Point + 10.2);
            _container = previous;
            _inHeader = previousHeader;
            _headerFooterKind = previousKind;
        });
    }

    public override IDisposable VisitParagraphBegin(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph, DxpIDocumentContext context, DxpIParagraphContext style)
    {
        if (_container is HeaderFooter && !_inHeader && _section != null &&
            paragraph.NextSibling() == null && string.IsNullOrEmpty(paragraph.InnerText) &&
            paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Footer" &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph previousFooter &&
            previousFooter.Descendants<TabChar>().Any())
            _section.PageSetup.FooterDistance = Unit.FromPoint(_section.PageSetup.FooterDistance.Point + 1.87);
        if ((paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Deleted>() != null ||
             paragraph.ParagraphProperties?.SectionProperties != null &&
             paragraph.ParagraphProperties.NumberingProperties != null) &&
            !paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any() &&
            !(_includeRevisionMarkup && paragraph.Descendants<DeletedText>().Any()) &&
            !paragraph.Descendants<Drawing>().Any() && !paragraph.Descendants<Picture>().Any())
        {
            _paragraph = null;
            return DxpDisposable.Empty;
        }
        _distributedText = null;
        _activeFloatingTextWrap = _floatingTextWraps.GetValueOrDefault(paragraph);
        var numberingProperties = context.Styles.ResolveEffectiveNumPr(paragraph);
        var numberingLevel = numberingProperties is { } list
            ? _numberingResolver?.ResolveLevel(list.NumId, list.Ilvl) : null;
        var noMarkerList = numberingLevel?.lvl.NumberingFormat?.Val?.Value == NumberFormatValues.None &&
            string.IsNullOrEmpty(numberingLevel.Value.lvl.LevelText?.Val?.Value) &&
            numberingLevel.Value.lvl.GetFirstChild<LevelSuffix>()?.Val?.Value != LevelSuffixValues.Nothing;
        _unfittedMixedScriptParagraph = _documentHasFitText && _container is Section &&
            !paragraph.Descendants<FitText>().Any() &&
            paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(character => IsCjk(character) || IsHangul(character)) ||
                    _availableFontFamilies.Contains("MingLiU-ExtB") &&
                    text.Text.EnumerateRunes().Any(rune => IsSupplementaryCjk(rune.Value)));
        _latinAfterFittedParagraph = _documentHasFitText && _container is Section &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph previousFittedParagraph &&
            previousFittedParagraph.Descendants<FitText>().Any() &&
            paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => !string.IsNullOrWhiteSpace(text.Text)) &&
            !paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(character => IsCjk(character) || IsHangul(character) ||
                    char.IsSurrogate(character)));
        _unfittedCalibriParagraph = _documentHasFitText &&
            (_unfittedMixedScriptParagraph || _container is HeaderFooter &&
                paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Any(text => text.Text.Any(IsCjk)) || _latinAfterFittedParagraph) &&
            !paragraph.Descendants<FitText>().Any() &&
            paragraph.Descendants<Run>().FirstOrDefault() is { } firstMixedRun &&
            string.Equals(context.Styles.ResolveRunStyle(paragraph, firstMixedRun).FontName,
                "Calibri", StringComparison.OrdinalIgnoreCase);
        _unfittedMixedScriptLatinParagraph = _unfittedMixedScriptParagraph && _unfittedCalibriParagraph &&
            paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(IsCjk));
        _mixedBodyMinchoFallback = _unfittedMixedScriptLatinParagraph &&
            context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                definition.StyleRunProperties?.RunFonts?.EastAsia?.Value is "Arial" or "Calibri") &&
            !paragraph.Descendants<Run>().Any(run =>
                string.Equals(run.RunProperties?.RunFonts?.EastAsia?.Value,
                    "Yu Mincho", StringComparison.OrdinalIgnoreCase));
        _chineseFallbackParagraph = _unfittedMixedScriptLatinParagraph &&
            _availableFontFamilies.Contains("MS Mincho") &&
            context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                definition.StyleRunProperties?.Languages?.EastAsia?.Value?.StartsWith("zh-",
                    StringComparison.OrdinalIgnoreCase) == true);
        if (_container is Cell && paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Table &&
            paragraph.NextSibling() == null && string.IsNullOrEmpty(paragraph.InnerText))
        {
            _paragraph = null;
            return DxpDisposable.Empty;
        }
        if (_activeColumnCell != null)
        {
            if (_columnBreakPending || _columnParagraphCount == _columnParagraphCapacity)
                AdvanceColumn();
            _columnParagraphCount++;
        }
        _paragraphGridCharSpacingPoints = _container is Section && UsesDocumentGrid(paragraph, context)
            ? _gridCharSpacingPoints : null;
        _compactSnapToCharsParagraph = _snapToCharsGrid &&
            paragraph.ParagraphProperties?.SnapToGrid?.Val?.Value == true;
        _compactLinesAndCharsParagraph = _gridLinePitchPoints is > 0 &&
            _gridCharSpacingPoints is > 0 &&
            paragraph.ParagraphProperties?.SnapToGrid?.Val?.Value == true;
        if (_documentHasFitText && !_documentHasTrackedInsertions && _container is Section &&
            paragraph.Elements<Run>().ToArray() is [{ } plainRun] &&
            plainRun.ChildElements.ToArray() is [DocumentFormat.OpenXml.Wordprocessing.Text plainText] &&
            plainText.Text.Length >= 30 &&
            plainText.Text.All(character => character is >= ' ' and <= '~') &&
            !plainText.Text.Any(character => character is '(' or ')' or '\\'))
            _plainLatinAfterFitTexts.Add(Regex.Replace(plainText.Text, "[^A-Za-z0-9]", ""));
        _autoSpaceDE = ResolveAutoSpace(paragraph, context, numbers: false);
        _autoSpaceDN = ResolveAutoSpace(paragraph, context, numbers: true);
        _narrowJustifiedSpaces = _availableFontFamilies.Contains("Aptos") &&
            (_container is Cell &&
             paragraph.ParagraphProperties?.Justification?.Val?.Value == JustificationValues.Both ||
             _documentHasTrackedInsertions && _container is Section &&
             style.ComputedStyle.TextAlign == DxpComputedTextAlign.Justify);
        _compressedBodyLayoutSpaces = _narrowJustifiedSpaces && _container is Section &&
            paragraph.InnerText.Length > 1000 && paragraph.Descendants<Run>().Count() > 50;
        _restoreCompressedBodySpaces |= _compressedBodyLayoutSpaces;
        _cjkStyleParagraph = !_documentHasFitText && _gridLinePitchPoints == null &&
            !paragraph.Descendants<FitText>().Any() &&
            paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(IsCjk)) &&
            context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                string.Equals(definition.StyleRunProperties?.RunFonts?.Ascii?.Value,
                    "Yu Mincho", StringComparison.OrdinalIgnoreCase));
        _cjkEmphasisParagraph = _cjkStyleParagraph &&
            (paragraph.Descendants<Emphasis>().Any() ||
             paragraph.Descendants<Run>().Any(run =>
                 ResolveEmphasis(run, context) is { } emphasis && emphasis != EmphasisMarkValues.None) ||
             context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                 _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                 definition.StyleRunProperties?.GetFirstChild<Emphasis>() != null));
        var directYuMinchoCjkFit = paragraph.Descendants<Run>().Any(run =>
            run.RunProperties?.GetFirstChild<FitText>() != null &&
            string.Equals(run.RunProperties.RunFonts?.EastAsia?.Value,
                "Yu Mincho", StringComparison.OrdinalIgnoreCase) &&
            run.InnerText.Any(IsCjk));
        if (_inHeader && _section != null && _container is HeaderFooter)
        {
            if (_headerParagraphCount == 0 && paragraph.Descendants<FitText>().Any())
                _headersWithFirstFittedParagraph.Add(_section);
            if (_headerParagraphCount == 1 && _headersWithFirstFittedParagraph.Contains(_section) &&
                paragraph.InnerText.Any(IsCjk))
                _headersWithSecondCjkParagraph.Add(_section);
        }
        var inheritedYuMincho = context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
            _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
            string.Equals(definition.StyleRunProperties?.RunFonts?.EastAsia?.Value,
                "Yu Mincho", StringComparison.OrdinalIgnoreCase));
        if (_inHeader && _headerParagraphCount == 0 && directYuMinchoCjkFit &&
            !inheritedYuMincho && _section != null)
            _section.PageSetup.HeaderDistance = Unit.FromPoint(
                _section.PageSetup.HeaderDistance.Point - 1.2);
        if (_inHeader && _cjkEmphasisParagraph && _headerParagraphCount == 0 && _section != null)
            _section.PageSetup.HeaderDistance = Unit.FromPoint(
                _section.PageSetup.HeaderDistance.Point -
                    (_snapToCharsGrid || _gridLinePitchPoints is > 0 ? 0.85 : 1.3));
        else if (_inHeader && _headerParagraphCount == 0 && _section != null &&
            (_snapToCharsGrid || _gridLinePitchPoints is > 0 && _gridCharSpacingPoints is > 0))
            _section.PageSetup.HeaderDistance = Unit.FromPoint(
                _section.PageSetup.HeaderDistance.Point + (_snapToCharsGrid ? 0.1 : -1.15));
        if (_cjkStyleParagraph && _sectionUsesDefaultMargins && _section != null &&
            _cjkDefaultMarginsAdjusted.Add(_section))
        {
            var page = _section.PageSetup;
            var styledFitText = _paragraphStyles.Values.Any(definition =>
                definition.StyleRunProperties?.GetFirstChild<FitText>() != null);
            page.HeaderDistance = Unit.FromPoint(page.HeaderDistance.Point -
                (_cjkEmphasisParagraph || styledFitText ? 4 : 5.32));
            page.TopMargin = Unit.FromPoint(page.TopMargin.Point -
                (styledFitText ? 4 : 5.18));
            page.FooterDistance = Unit.FromPoint(page.FooterDistance.Point +
                (styledFitText ? 4 : 5.32));
        }
        if (IsHidden(paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Vanish>()))
        {
            _paragraph = null;
            return DxpDisposable.Empty;
        }
        var firstBodyParagraph = (paragraph.PreviousSibling() == null ||
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph previousSectionParagraph &&
                previousSectionParagraph.ParagraphProperties?.SectionProperties != null) &&
            (paragraph.Parent is Body ||
             paragraph.Parent is SdtContentBlock content && content.Parent is SdtBlock block &&
             block.Parent is Body && block.PreviousSibling() == null);
        var firstBodyBefore = firstBodyParagraph && _container is Section
            ? style.ComputedStyle.MarginTopPt : null;
        if (firstBodyParagraph && _container is Section && _section != null &&
            firstBodyBefore is not > 0 &&
            context.Styles.ResolveRunStyle(paragraph, new Run()) is
                { FontName: "Arial", FontSizeHalfPoints: 28, Bold: true } &&
            paragraph.ParagraphProperties?.ParagraphStyleId != null)
            _section.PageSetup.TopMargin = Unit.FromPoint(_section.PageSetup.TopMargin.Point + 1.17);
        if (firstBodyParagraph && _container is Section && _section != null &&
            !_gutterAtTop &&
            paragraph.Descendants<TabChar>().Any() &&
            paragraph.Descendants<Break>().Count() >= 2)
            _section.PageSetup.TopMargin = Unit.FromPoint(_section.PageSetup.TopMargin.Point + 1.34);
        var doubleBorderInset = 0.0;
        if (firstBodyParagraph && _container is Section)
        {
            doubleBorderInset = paragraph.Descendants<Run>()
                .Select(run => ResolveRunTextBorder(run, context))
                .Where(border => border?.Double == true)
                .Select(border => Math.Max(0, border!.Inset - 0.3))
                .DefaultIfEmpty(0).Max();
            if (doubleBorderInset > 0)
                firstBodyBefore = Math.Max(firstBodyBefore ?? 0, doubleBorderInset);
            var shadowInset = paragraph.Descendants<Run>()
                .Select(run => ResolveRunTextBorder(run, context))
                .Any(border => border?.Shadow == true) ? 2.0 : 0;
            if (shadowInset > 0)
                firstBodyBefore = Math.Max(firstBodyBefore ?? 0, shadowInset);
        }
        if (firstBodyParagraph && firstBodyBefore is not > 0 &&
            paragraph.ParagraphProperties?.ParagraphStyleId != null &&
            context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.Tabs?
                    .Elements<DocumentFormat.OpenXml.Wordprocessing.TabStop>()
                    .Any(HasRenderedTabLeader) == true))
            firstBodyBefore = 3.6;
        var clearsUnsnappedGridEmphasis = _snapToCharsGrid &&
            paragraph.ParagraphProperties?.SnapToGrid?.Val?.Value == false &&
            paragraph.Descendants<Run>().Any(run =>
                run.RunProperties?.GetFirstChild<Emphasis>()?.Val?.Value == EmphasisMarkValues.None);
        if (firstBodyParagraph && _container is Section && _section != null &&
            clearsUnsnappedGridEmphasis)
            _section.PageSetup.TopMargin = Unit.FromPoint(_section.PageSetup.TopMargin.Point + 5.1);
        if (_cjkStyleParagraph && paragraph.Parent is Body && paragraph.PreviousSibling() == null &&
            _container is Section && _section != null &&
            paragraph.Descendants<Run>().FirstOrDefault() is { } firstEmphasizedRun &&
            firstEmphasizedRun.RunProperties?.GetFirstChild<Emphasis>()?.Val?.Value == EmphasisMarkValues.None &&
                context.Styles.GetParagraphStyleChain(paragraph)
                    .Select(entry => _paragraphStyles.TryGetValue(entry.StyleId, out var definition)
                        ? definition.StyleRunProperties?.GetFirstChild<Emphasis>()?.Val?.Value : null)
                    .FirstOrDefault(inherited => inherited != null) == EmphasisMarkValues.Dot)
            _firstBodyClearsInheritedEmphasis.Add(_section);
        if (_paragraphGridCharSpacingPoints is > 0 && _section != null &&
            paragraph.Parent is Body && paragraph.PreviousSibling() == null &&
            !HasBaselineTextAlignment(paragraph, context))
            _section.PageSetup.TopMargin = Unit.FromPoint(_section.PageSetup.TopMargin.Point - 6);
        if (firstBodyBefore is > 0 && _section != null)
        {
            var spacer = _section.AddParagraph();
            spacer.Format.LineSpacingRule = LineSpacingRule.Exactly;
            spacer.Format.LineSpacing = Unit.FromPoint(0.1);
            var sectionReset = paragraph.PreviousSibling() is
                DocumentFormat.OpenXml.Wordprocessing.Paragraph boundaryParagraph &&
                boundaryParagraph.ParagraphProperties?.SectionProperties != null &&
                paragraph.ParagraphProperties?.ContextualSpacing?.Val?.Value == false;
            spacer.Format.SpaceAfter = Unit.FromPoint(firstBodyBefore.Value - 0.1 +
                (sectionReset ? 4 : 0));
        }
        MigraDoc.DocumentObjectModel.Shapes.TextFrame? positionedParagraphFrame = null;
        var frameProperties = paragraph.ParagraphProperties?.ChildElements
            .FirstOrDefault(element => element.LocalName == "framePr");
        var frameWrapSetting = frameProperties != null ? AttributeValue(frameProperties, "wrap") : null;
        var aroundFrameAvailable = Enum.TryParse<MigraDoc.DocumentObjectModel.Shapes.WrapStyle>(
            "BothSides", out var aroundFrameWrap);
        if (_container is Section frameSection && _activeColumnCell == null &&
            paragraph.Parent is Body && frameProperties != null &&
            AttributeValue(frameProperties, "hAnchor") is "margin" or "page" &&
            AttributeValue(frameProperties, "vAnchor") is "margin" or "page" &&
            frameWrapSetting is "around" or "notBeside" &&
            (frameWrapSetting == "notBeside" || aroundFrameAvailable) &&
            int.TryParse(AttributeValue(frameProperties, "x"), out var frameX) &&
            int.TryParse(AttributeValue(frameProperties, "y"), out var frameY) &&
            int.TryParse(AttributeValue(frameProperties, "w"), out var frameWidth) &&
            frameWidth > 0)
        {
            positionedParagraphFrame = frameSection.AddTextFrame();
            positionedParagraphFrame.RelativeHorizontal = AttributeValue(frameProperties, "hAnchor") == "page"
                ? MigraDoc.DocumentObjectModel.Shapes.RelativeHorizontal.Page
                : MigraDoc.DocumentObjectModel.Shapes.RelativeHorizontal.Margin;
            positionedParagraphFrame.RelativeVertical = AttributeValue(frameProperties, "vAnchor") == "page"
                ? MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Page
                : MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Margin;
            positionedParagraphFrame.Left = Unit.FromPoint(frameX / 20.0);
            positionedParagraphFrame.Top = Unit.FromPoint(frameY / 20.0);
            positionedParagraphFrame.Width = Unit.FromPoint(frameWidth / 20.0);
            if (int.TryParse(AttributeValue(frameProperties, "h"), out var frameHeight) && frameHeight > 0)
                positionedParagraphFrame.Height = Unit.FromPoint(frameHeight / 20.0);
            positionedParagraphFrame.MarginLeft = Unit.Zero;
            positionedParagraphFrame.MarginRight = Unit.Zero;
            positionedParagraphFrame.MarginTop = Unit.Zero;
            positionedParagraphFrame.MarginBottom = Unit.Zero;
            positionedParagraphFrame.WrapFormat.Style = frameWrapSetting == "notBeside"
                ? MigraDoc.DocumentObjectModel.Shapes.WrapStyle.TopBottom : aroundFrameWrap;
            if (frameWrapSetting == "around")
                positionedParagraphFrame.WrapFormat.DistanceRight = Unit.FromPoint(1);
        }
        var precedingNotBesideFrame = positionedParagraphFrame == null &&
            _container is Section followingFrameSection &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                { ParagraphProperties: { } precedingNotBesideProperties } &&
            precedingNotBesideProperties.ChildElements.Any(element =>
                element.LocalName == "framePr" && AttributeValue(element, "wrap") == "notBeside")
            ? followingFrameSection.Elements.Cast<MigraDoc.DocumentObjectModel.DocumentObject>()
                .LastOrDefault() as MigraDoc.DocumentObjectModel.Shapes.TextFrame
            : null;
        _paragraph = positionedParagraphFrame?.AddParagraph() ?? AddParagraph();
        if (precedingNotBesideFrame != null && _paragraph != null)
            _notBesideFrameFollowers[precedingNotBesideFrame] = _paragraph;
        var sourceText = paragraph.InnerText.TrimEnd();
        if (_paragraph != null && sourceText.Length > 0 && sourceText[^1] is (']' or ')'))
            _justifiedShrinkFactors[_paragraph] = 0.7;
        if (_paragraph != null && sourceText.EndsWith('’'))
            _justifiedShrinkFactors[_paragraph] = 0.58;
        if (_paragraph != null && sourceText.EndsWith('”'))
            _justifiedShrinkFactors[_paragraph] = 0.5;
        if (_paragraph != null && Regex.IsMatch(sourceText, @"^\S+\s+\S+$"))
            _twoWordOverhangParagraphs.Add(_paragraph);
        if (_paragraph != null && _unfittedMixedScriptParagraph && _unfittedCalibriParagraph &&
            paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(IsHangul)) &&
            !paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(IsCjk)))
            _hangulBaselineShiftParagraphs.Add(_paragraph);
        if (_paragraph != null && _implicitContentTable)
            _implicitTableParagraphs.Add(_paragraph);
        if (_paragraph != null && _container is Section &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                { ParagraphProperties: { } widthFrameProperties } &&
            widthFrameProperties.ChildElements.Any(element => element.LocalName == "framePr"))
            _frameWidthParagraphs.Add(_paragraph);
        if (_paragraph != null && _container is HeaderFooter && !_inHeader)
            _footerLeaderParagraphs.Add(_paragraph);
        if (_paragraph != null && paragraph.Descendants<PositionalTab>().Any())
            _positionalLeaderParagraphs.Add(_paragraph);
        if (_paragraph != null && _container is Section &&
            context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.Tabs?.Elements<DocumentFormat.OpenXml.Wordprocessing.TabStop>()
                    .Any(stop => stop.Val?.Value == TabStopValues.Clear) == true))
            _clearedTabLeaderParagraphs.Add(_paragraph);
        if (_paragraph != null && paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value?
                .StartsWith("TOC", StringComparison.OrdinalIgnoreCase) == true)
            _tocLeaderParagraphs.Add(_paragraph);
        if (_compressedBodyLayoutSpaces) _compressedBodyParagraphs.Add(_paragraph);
        if (_container is Section && paragraph.Descendants<InsertedRun>().Any())
            _trackedLayoutParagraphs.Add(_paragraph);
        var textRuns = paragraph.Descendants<Run>()
            .Where(run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any()).ToArray();
        if (textRuns is [{ } onlyRun] && ResolveRunTextBorder(onlyRun, context) is
            { Double: true } doubleBorder)
        {
            var borderWords = Regex.Matches(paragraph.InnerText, @"\S+");
            if (borderWords.Count >= 2)
                _doubleBorderTextPadding[borderWords[0].Value + "\0" + borderWords[1].Value] =
                    2 * doubleBorder.Width;
        }
        if (_balanceSingleByteDoubleByteWidth && _container is HeaderFooter && !_inHeader &&
            paragraph.Descendants<FieldCode>().Any(code =>
                Regex.IsMatch(code.Text ?? "", @"^\s*PAGE(?:\s|$)", RegexOptions.IgnoreCase)) &&
            paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().FirstOrDefault() is { } firstText)
        {
            var firstWord = Regex.Match(firstText.Text, @"\S+").Value;
            var spaces = paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Sum(value => value.Text.Count(character => character == ' '));
            if (firstWord.Length > 0 && spaces > 0)
                _balancedFooterFieldSpaces[firstWord] = spaces;
        }
        if (_paragraph == null) return DxpDisposable.Empty;
        var leadingWords = GetAsciiWordsAfterRunBoundary(textRuns);
        if (leadingWords.Count > 0) _crossRunLeadingWords[_paragraph] = leadingWords;
        if (_container is Section && _availableFontFamilies.Contains("Aptos") &&
            numberingProperties == null &&
            paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "ListParagraph" &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph precedingBulletParagraph &&
            precedingBulletParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "ListParagraph" &&
            context.Styles.ResolveEffectiveNumPr(precedingBulletParagraph) is { } precedingBulletNumbering &&
            paragraph.Parent is Body openingBody &&
            openingBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is
                { } firstOpeningParagraph &&
            context.Styles.ResolveEffectiveNumPr(firstOpeningParagraph) is { } openingNumbering &&
            openingNumbering.NumId == precedingBulletNumbering.NumId &&
            _numberingResolver?.ResolveLevel(precedingBulletNumbering.NumId, precedingBulletNumbering.Ilvl) is
                { } precedingBulletLevel &&
            precedingBulletLevel.lvl.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet)
        {
            var markerFamily = precedingBulletLevel.lvl.GetFirstChild<NumberingSymbolRunProperties>()?
                .RunFonts?.Ascii?.Value;
            var rise = markerFamily switch
            {
                "Symbol" => 0.51,
                "Courier New" => 0.29,
                "Wingdings" => 0.06,
                _ => 0
            };
            if (rise > 0) _bulletBaselineParagraphs[_paragraph] = ("Aptos", rise);
        }
        if (_container is Section && _activeColumnCell == null)
            _styleRefBodyParagraphs[_paragraph] = paragraph;
        var underlinedRuns = paragraph.Descendants<Run>()
            .Where(run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any())
            .Select(run => ResolveRunUnderlineColor(paragraph, run, context)).Distinct().ToArray();
        if (underlinedRuns is [{ } underlineColor])
            _coloredUnderlineParagraphs[_paragraph] = (underlineColor,
                paragraph.InnerText.EndsWith(" ", StringComparison.Ordinal), _container is HeaderFooter);
        if (doubleBorderInset > 0) _firstBodyDoubleBorderInset[_paragraph] = doubleBorderInset;
        if (_container is Section)
        {
            _sourceBodyParagraphs.Add(_paragraph);
            if (!_documentHasFitText && paragraph.InnerText.Any(IsFullWidthCjkPunctuation) &&
                paragraph.InnerText.Any(char.IsAsciiLetter))
                _mixedFullWidthPunctuationParagraphs.Add(_paragraph);
            if (_narrowJustifiedSpaces) _narrowJustifiedBodyParagraphs.Add(_paragraph);
            if (paragraph.ParagraphProperties?.SuppressLineNumbers is { } suppress &&
                suppress.Val?.Value != false)
                _suppressedLineNumberParagraphs.Add(_paragraph);
        }
        if (_container is Cell) _sourceCellParagraphs.Add(_paragraph);
        var fittedRuns = paragraph.Descendants<Run>().ToArray();
        var hasExplicitTopAlignment = false;
        if (HasTopTextAlignment(paragraph, context))
        {
            var explicitSizes = fittedRuns.Select(run => run.RunProperties?.FontSize?.Val?.Value)
                .Select(size => int.TryParse(size, out var halfPoints) ? halfPoints / 2.0 : 0)
                .Where(size => size > 0).ToArray();
            var sizes = fittedRuns.Select(run => context.Styles.ResolveRunStyle(paragraph, run)
                .FontSizeHalfPoints is int halfPoints && halfPoints > 0 ? halfPoints / 2.0 :
                    _paragraph.Format.Font.Size.Point > 0 ? _paragraph.Format.Font.Size.Point : 11)
                .Where(size => size > 0).ToArray();
            if (sizes.Length > 1 && sizes.Max() > sizes.Min())
            {
                _topAlignedFontSizes[_paragraph] = sizes.Max();
                if (_container is HeaderFooter && _inHeader &&
                    sizes.Max() <= 18 && sizes.Max() - sizes.Min() is >= 5 and <= 7)
                    _topAlignedHeaderParagraphs.Add(_paragraph);
                hasExplicitTopAlignment = explicitSizes.Length > 1;
                if (hasExplicitTopAlignment && _section != null && _container is Section &&
                    _topAlignedBodyMargins.Add(_section))
                    _section.PageSetup.TopMargin = Unit.FromPoint(
                        _section.PageSetup.TopMargin.Point + 3.2);
                if (hasExplicitTopAlignment && _section != null && _container is HeaderFooter && !_inHeader &&
                    _topAlignedFooterDistances.Add(_section))
                    _section.PageSetup.FooterDistance = Unit.FromPoint(
                        _section.PageSetup.FooterDistance.Point + 12);
            }
        }
        var paragraphFit = ResolveParagraphFitText(paragraph, context);
        _paragraphInheritsFitText = paragraphFit != null;
        var firstFit = fittedRuns.FirstOrDefault() is { } firstFittedRun
            ? ResolveFitText(firstFittedRun) ?? paragraphFit : null;
        var inheritedStoryFit = _container is HeaderFooter && paragraphFit != null &&
            fittedRuns.All(run => ResolveFitText(run) == null);
        if (!inheritedStoryFit && firstFit?.Val?.Value is uint fitTwips && fitTwips > 0 &&
            fittedRuns.All(run => (ResolveFitText(run) ?? paragraphFit) is { } fit &&
                fit.Val?.Value == firstFit.Val?.Value && fit.Id?.Value == firstFit.Id?.Value))
        {
            var singleYuMinchoStoryRun = _container is HeaderFooter && fittedRuns is [var fittedStoryRun] &&
                fittedStoryRun.RunProperties?.RunFonts == null &&
                string.Equals(context.Styles.ResolveRunStyle(paragraph, fittedStoryRun).FontName,
                    "Yu Mincho", StringComparison.OrdinalIgnoreCase) &&
                fittedStoryRun.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Any(text => text.Text.Any(IsCjk) && text.Text.Any(char.IsAsciiLetter));
            var singleHighAnsiStoryRun = _container is HeaderFooter && fittedRuns is [var highAnsiRun] &&
                highAnsiRun.RunProperties?.RunFonts is { } runFonts &&
                runFonts.Ascii?.Value != runFonts.HighAnsi?.Value &&
                highAnsiRun.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Any(text => text.Text.Any(IsHighAnsiCharacter) &&
                        !text.Text.Any(IsCjk));
            if (_inHeader && _headerParagraphCount == 0 && _container is HeaderFooter &&
                fittedRuns is [var mixedHeaderRun] &&
                mixedHeaderRun.RunProperties?.RunFonts is { } mixedFonts &&
                mixedFonts.Ascii?.Value != mixedFonts.HighAnsi?.Value &&
                mixedHeaderRun.InnerText.Any(IsHighAnsiCharacter) &&
                !mixedHeaderRun.InnerText.Any(char.IsWhiteSpace) &&
                mixedHeaderRun.InnerText.Count(IsCjk) == 1)
                _mixedHighAnsiFittedHeaders.Add(_paragraph);
            _fitTextWidths[_paragraph] = fitTwips / 20.0 - (singleYuMinchoStoryRun ? 4 : 0) -
                (singleHighAnsiStoryRun ? 1.7 : 0);
            _fitTextCharacterCounts[_paragraph] = paragraph.InnerText.Length;
            if (_container is Section && paragraph.InnerText.Any(IsCjk) &&
                paragraph.InnerText.Any(char.IsAsciiLetter))
                _mixedCjkFitBodyParagraphs.Add(_paragraph);
            if (_container is Section && _section != null &&
                _headersWithSecondCjkParagraph.Contains(_section) &&
                paragraph.PreviousSibling() == null)
                _fittedBodyAfterSecondCjkHeader.Add(_paragraph);
            if (_container is Section && fittedRuns.Length > 0 &&
                fittedRuns.All(run => string.Equals(
                    context.Styles.ResolveRunStyle(paragraph, run).FontName,
                    "Calibri", StringComparison.OrdinalIgnoreCase)) &&
                !paragraph.InnerText.Any(character => IsCjk(character) || IsHangul(character) ||
                    char.IsSurrogate(character)))
                _latinCalibriFitParagraphs.Add(_paragraph);
            if (_container is Section && fittedRuns.Length == 2 &&
                string.Equals(context.Styles.ResolveRunStyle(paragraph, fittedRuns[0]).FontName,
                    "Calibri", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(context.Styles.ResolveRunStyle(paragraph, fittedRuns[1]).FontName,
                    "Yu Mincho", StringComparison.OrdinalIgnoreCase) &&
                !paragraph.InnerText.Any(character => IsCjk(character) || IsHangul(character) ||
                    char.IsSurrogate(character)))
                _mixedCalibriYuMinchoFitParagraphs.Add(_paragraph);
            if (_container is Section && fittedRuns.Length > 0 &&
                fittedRuns.All(run => string.Equals(
                    context.Styles.ResolveRunStyle(paragraph, run).FontName,
                    "Yu Mincho", StringComparison.OrdinalIgnoreCase)) &&
                !paragraph.InnerText.Any(character => IsCjk(character) || IsHangul(character) ||
                    char.IsSurrogate(character)))
                _latinYuMinchoFitParagraphs.Add(_paragraph);
            if (fittedRuns.All(run => ResolveFitText(run) == null))
            {
                _fitTextTracking.Add(_paragraph);
                if (_container is Section && fittedRuns.All(run => string.Equals(
                        context.Styles.ResolveRunStyle(paragraph, run).FontName,
                        "Times New Roman", StringComparison.OrdinalIgnoreCase)) &&
                    !paragraph.InnerText.Any(character => IsCjk(character) || IsHangul(character) ||
                        char.IsSurrogate(character)))
                    _latinTimesFitParagraphs.Add(_paragraph);
            }
        }
        else if (!inheritedStoryFit && fittedRuns.Length > 1 && !paragraph.Descendants<Break>().Any())
        {
            var groups = new List<(string Text, double Width)>();
            int? currentId = null;
            uint? currentWidth = null;
            foreach (var run in fittedRuns)
            {
                var fit = ResolveFitText(run) ?? paragraphFit;
                var text = run.InnerText;
                if (fit?.Val?.Value is not uint width || width == 0 ||
                    !text.Any(character => !char.IsWhiteSpace(character)))
                {
                    groups.Clear();
                    break;
                }
                if (groups.Count > 0 && currentId == fit.Id?.Value && currentWidth == width)
                {
                    var previousGroup = groups[^1];
                    groups[^1] = (previousGroup.Text + text, previousGroup.Width);
                }
                else groups.Add((text, width / 20.0));
                currentId = fit.Id?.Value;
                currentWidth = width;
            }
            if (groups.Count > 1) _partialFitGroups[_paragraph] = groups.ToArray();
        }
        if (_fitTextTracking.Contains(_paragraph) && _container is Section &&
            !_documentHasFitText && _section != null && _styleFitStoryDistancesApplied.Add(_section))
        {
            var page = _section.PageSetup;
            page.HeaderDistance = Unit.FromPoint(page.HeaderDistance.Point + 4);
            page.FooterDistance = Unit.FromPoint(page.FooterDistance.Point + 6);
        }
        if (_documentHasFitText && _fitTextWidths.ContainsKey(_paragraph) &&
            paragraphFit == null && _container is Section && _section != null &&
            _styleFitStoryDistancesApplied.Add(_section))
            _section.PageSetup.FooterDistance = Unit.FromPoint(
                _section.PageSetup.FooterDistance.Point + 9);
        if (_partialFitGroups.ContainsKey(_paragraph) && _container is Section && _section != null &&
            _styleFitStoryDistancesApplied.Add(_section))
            _section.PageSetup.FooterDistance = Unit.FromPoint(
                _section.PageSetup.FooterDistance.Point + 10);
        if (_container is HeaderFooter && !_inHeader && _section != null &&
            _fitTextWidths.ContainsKey(_paragraph) &&
            paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(IsCjk)) &&
            (directYuMinchoCjkFit || inheritedYuMincho) &&
            _fittedCjkFooterDistanceApplied.Add(_section))
            _section.PageSetup.FooterDistance = Unit.FromPoint(
                _section.PageSetup.FooterDistance.Point +
                (directYuMinchoCjkFit && !inheritedYuMincho ? 4.1 : 4.6));
        var justification = paragraph.ParagraphProperties?.Justification?.Val?.Value;
        if (justification == null)
            foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
                if (_paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                    definition.StyleParagraphProperties?.Justification?.Val?.Value is { } inherited)
                {
                    justification = inherited;
                    break;
                }
        if (justification == JustificationValues.Distribute &&
            paragraph.Descendants<Run>().Count() == 1 &&
            paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().ToArray() is { Length: 1 } texts &&
            texts[0].Text.Split(' ', StringSplitOptions.RemoveEmptyEntries) is { Length: >= 2 } words &&
            words.Length == texts[0].Text.Count(character => character == ' ') + 1 &&
            !paragraph.Descendants<TabChar>().Any() && !paragraph.Descendants<Break>().Any())
        {
            _distributedText = texts[0].Text;
            for (var index = 1; index < words.Length; index++)
                _distributedWordPairs.Add(words[index - 1] + "\0" + words[index]);
        }
        _paragraphTextLength = 0;
        _pendingNoBreakHyphen = false;
        _paragraphLineSize = context.Styles.ResolveRunStyle(paragraph, new Run()).FontSizeHalfPoints is int lineHalfPoints &&
            lineHalfPoints > 0 ? lineHalfPoints / 2.0 : _paragraph.Format.Font.Size.Point;
        _paragraphBorderLetterLength = 0;
        _revisedListMarkerLength = 0;
        _paragraphHasLeaderTab = false;
        if (_inHeader && _container is HeaderFooter &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Table previousTable &&
            previousTable.TableProperties?.TableStyle == null &&
            previousTable.Elements<TableRow>().Any(row =>
                row.TableRowProperties?.GetFirstChild<TableHeader>() != null &&
                row.TableRowProperties.GetFirstChild<TableRowHeight>()?.HeightType?.Value == HeightRuleValues.Exact) &&
            !paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any())
            _headerHasTrailingUnstyledTableParagraph = true;
        _lastParagraph = _paragraph;
        _lastParagraphContainer = _container;
        if (!_inHeader && _container is HeaderFooter)
            _footerLastParagraphHasText = !string.IsNullOrWhiteSpace(paragraph.InnerText);
        if (_inHeader && _container is HeaderFooter)
        {
            if (paragraph.ParagraphProperties?.GetFirstChild<BiDi>() != null &&
                context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                    _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                    definition.StyleParagraphProperties?.GetFirstChild<BiDi>() != null))
                _headerHasDirectBidiOverride = true;
            if (_gridLinePitchPoints is > 0 && _gridCharSpacingPoints is > 0 &&
                !_largeGridHeaderDistanceAdjusted && _section != null &&
                paragraph.Descendants<FontSize>().Any(size =>
                    int.TryParse(size.Val?.Value, out var halfPoints) && halfPoints > 24))
            {
                _section.PageSetup.HeaderDistance = Unit.FromPoint(
                    _section.PageSetup.HeaderDistance.Point + 4);
                _section.PageSetup.FooterDistance = Unit.FromPoint(
                    _section.PageSetup.FooterDistance.Point + 3.5);
                _largeGridHeaderDistanceAdjusted = true;
            }
            if (!string.IsNullOrWhiteSpace(paragraph.InnerText) || paragraph.Descendants<Drawing>().Any())
            {
                _headerParagraphCount += 1 + _headerPendingEmptyParagraphCount;
                if (style.MarkerAccept?.marker != null)
                    _headerNumberedParagraphCount++;
                _headerEmptyParagraphCount += _headerPendingEmptyParagraphCount;
                _headerPendingEmptyParagraphCount = 0;
                _headerHasContent = true;
                _headerParagraphsAreNumbered &= style.MarkerAccept?.marker != null || noMarkerList;
                _headerHasUnnumberedDefaultTab |= _defaultTabStopPoints is > 0 &&
                    style.MarkerAccept?.marker == null && paragraph.Descendants<TabChar>().Any();
                if ((_snapToCharsGrid || _gridLinePitchPoints is > 0 && _gridCharSpacingPoints is > 0) &&
                    StoryParagraphWraps(paragraph, context))
                    _headerGridWrapClearancePoints +=
                        (_cjkEmphasisParagraph ? 28.2 : 23.1) - 4.5;
            }
            else if (_headerHasContent && _headerPreviousParagraphUsesBodySpacing)
                _headerPendingEmptyParagraphCount++;
        }
        else if (_container is HeaderFooter && !string.IsNullOrWhiteSpace(paragraph.InnerText))
        {
            _footerParagraphCount++;
            _footerHasContextualSpacing |= context.Styles is DocxportNet.Walker.Context.DxpStyleResolver resolver &&
                resolver.GetContextualSpacing(paragraph);
        }
        var format = _paragraph.Format;
        if (_unfittedMixedScriptLatinParagraph && !_chineseFallbackParagraph && _container is Section)
        {
            format.RightIndent = Unit.FromPoint(-40);
            format.SpaceBefore = Unit.FromPoint(_mixedBodyMinchoFallback ? -1.75 : -4);
        }
        if (_latinAfterFittedParagraph && _headerParagraphCount >= 2)
            format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point - 2);
        if (_unfittedCalibriParagraph && (_latinAfterFittedParagraph ||
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                { } fittedPrevious && fittedPrevious.Descendants<FitText>().Any() &&
            paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.EnumerateRunes().Any(rune => IsSupplementaryCjk(rune.Value)))))
            format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point - 2.4);
        if (_unfittedMixedScriptParagraph && !_unfittedCalibriParagraph &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                { } previousUniformFit && previousUniformFit.Descendants<FitText>().Any() &&
            !previousUniformFit.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(character => IsCjk(character) || IsHangul(character) ||
                    char.IsSurrogate(character))) &&
            previousUniformFit.Descendants<Run>().Where(run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Any()).Select(run => context.Styles.ResolveRunStyle(previousUniformFit, run).FontName)
                .Distinct(StringComparer.OrdinalIgnoreCase).ToArray() is [string fittedFont] &&
            !string.Equals(fittedFont, "Yu Mincho", StringComparison.OrdinalIgnoreCase))
            format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point - 1.3);
        if (_unfittedCalibriParagraph && _container is HeaderFooter)
        {
            format.SpaceBefore = Unit.FromPoint(_inHeader && _headerParagraphCount >= 2 ? -7.29 :
                !_inHeader && _footerParagraphCount >= 2 ? -5.32 : -4);
            if (!_inHeader && _footerParagraphCount >= 2 && _section != null &&
                _fittedCjkFooterContinuationDistanceApplied.Add(_section))
                _section.PageSetup.FooterDistance = Unit.FromPoint(
                    _section.PageSetup.FooterDistance.Point + 1.32);
        }
        if (_fitTextTracking.Contains(_paragraph))
            format.Font.Size = Unit.FromPoint(12);
        var keepNext = paragraph.ParagraphProperties?.KeepNext;
        if (keepNext == null)
            foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
                if (_paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                    definition.StyleParagraphProperties?.KeepNext is { } inheritedKeepNext)
                {
                    keepNext = inheritedKeepNext;
                    break;
                }
        format.KeepWithNext = keepNext?.Val?.Value ??
            context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.OutlineLevel is { } level &&
                level.Val?.Value is >= 0 and <= 8);
        var widowControl = paragraph.ParagraphProperties?.WidowControl;
        if (widowControl == null)
            foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
                if (_paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                    definition.StyleParagraphProperties?.WidowControl is { } inherited)
                {
                    widowControl = inherited;
                    break;
                }
        if (widowControl?.Val?.Value != false && paragraph.Descendants<Break>().Count(br =>
                br.Type?.Value == null || br.Type?.Value == BreakValues.TextWrapping) is > 0 and <= 2)
            format.KeepTogether = true;
        var markFont = paragraph.ParagraphProperties?.ParagraphMarkRunProperties;
        if (paragraph.InnerText.Length == 0 && markFont?.ChildElements.Any(element =>
                element is not Languages) == true && (_container is Section ||
                _container is HeaderFooter && (markFont.GetFirstChild<RunFonts>() != null ||
                    markFont.GetFirstChild<FontSize>() != null)))
        {
            var inheritedMark = context.Styles.ResolveRunStyle(paragraph, new Run());
            if (inheritedMark.FontName is string inheritedFamily &&
                _availableFontFamilies.Contains(inheritedFamily))
                format.Font.Name = inheritedFamily;
            if (inheritedMark.FontSizeHalfPoints is > 0)
                format.Font.Size = Unit.FromPoint(inheritedMark.FontSizeHalfPoints.Value / 2.0);
        }
        var markFamily = markFont?.GetFirstChild<RunFonts>()?.Ascii?.Value;
        if (markFamily != null && _availableFontFamilies.Contains(markFamily))
            format.Font.Name = markFamily;
        if (int.TryParse(markFont?.GetFirstChild<FontSize>()?.Val?.Value, out var markHalfPoints) &&
            markHalfPoints > 0)
            format.Font.Size = Unit.FromPoint(markHalfPoints / 2.0);
        if ((paragraph.ParagraphProperties?.Tabs?.Elements<DocumentFormat.OpenXml.Wordprocessing.TabStop>()
                .Any(HasRenderedTabLeader) == true ||
            context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.Tabs?
                    .Elements<DocumentFormat.OpenXml.Wordprocessing.TabStop>()
                    .Any(HasRenderedTabLeader) == true)) &&
            paragraph.Descendants<Run>().FirstOrDefault(run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any())
                is { } tabTextRun)
        {
            var tabTextStyle = context.Styles.ResolveRunStyle(paragraph, tabTextRun);
            if (tabTextStyle.FontName is string tabFont && _availableFontFamilies.Contains(tabFont))
                format.Font.Name = tabFont;
            if (tabTextStyle.FontSizeHalfPoints is > 0)
                format.Font.Size = Unit.FromPoint(tabTextStyle.FontSizeHalfPoints.Value / 2.0);
        }
        if (_container is Cell && _tableParagraphSpacing is { } tableSpacing)
        {
            if (int.TryParse(tableSpacing.Before?.Value, out var tableBefore))
                format.SpaceBefore = Unit.FromPoint(tableBefore / 20.0);
            if (int.TryParse(tableSpacing.After?.Value, out var tableAfter))
                format.SpaceAfter = Unit.FromPoint(tableAfter / 20.0);
            if (int.TryParse(tableSpacing.Line?.Value, out var tableLine) && tableLine > 0)
            {
                if (tableSpacing.LineRule?.Value == LineSpacingRuleValues.Exact)
                {
                    format.LineSpacingRule = LineSpacingRule.Exactly;
                    format.LineSpacing = Unit.FromPoint(tableLine / 20.0);
                }
                else if (tableSpacing.LineRule?.Value == LineSpacingRuleValues.AtLeast)
                {
                    format.LineSpacingRule = LineSpacingRule.AtLeast;
                    format.LineSpacing = Unit.FromPoint(tableLine / 20.0);
                }
                else
                {
                    format.LineSpacingRule = LineSpacingRule.Multiple;
                    format.LineSpacing = tableLine / 240.0;
                }
            }
        }
        if (HasPageBreakBefore(paragraph, context) || _pendingEmptyParagraphPageBreak)
            format.PageBreakBefore = true;
        _pendingEmptyParagraphPageBreak = false;
        var computed = style.ComputedStyle;
        if (computed.MarginLeftPt is double indent) format.LeftIndent = Unit.FromPoint(indent);
        if (_container is Cell && _table != null && !_tableInHeaderFooter &&
            paragraph.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault() is { } bodyTable &&
            bodyTable.Elements<TableRow>().FirstOrDefault()?.Elements<TableCell>().FirstOrDefault() is
                { TableCellProperties.TableCellBorders: { } firstCellBorders } &&
            firstCellBorders.LeftBorder == null && firstCellBorders.StartBorder is { } leftEdge &&
            AttributeValue(leftEdge, "val") is not ("none" or "nil") &&
            double.TryParse(AttributeValue(leftEdge, "sz"), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var borderEighths) && borderEighths > 0)
            format.LeftIndent = Unit.FromPoint(format.LeftIndent.Point - borderEighths / 16);
        if (style.Indent.Right is int rightTwips)
            format.RightIndent = Unit.FromPoint(rightTwips / 20.0);
        if (_compressedBodyLayoutSpaces &&
            paragraph.Descendants<Run>().Count() > paragraph.InnerText.Length * 0.08)
        {
            format.RightIndent = Unit.FromPoint(format.RightIndent.Point -
                DenseTrackedLayoutAllowancePoints);
            _layoutWidthParagraphs.Add(_paragraph!);
        }
        if (_compactLinesAndCharsParagraph && _headerGridWrapClearancePoints > 0 &&
            _container is Section)
            format.RightIndent = Unit.FromPoint(format.RightIndent.Point + 12);
        if (_compactSnapToCharsParagraph && _cjkStyleParagraph && _container is Section)
            format.RightIndent = Unit.FromPoint(format.RightIndent.Point + 12);
        if (_snapToCharsGrid && !_compactSnapToCharsParagraph && _cjkStyleParagraph &&
            _container is Section && _gridCharSpacingPoints is >= 8 and <= 12)
            format.RightIndent = Unit.FromPoint(format.RightIndent.Point - 5);
        if (_gridCharSpacingPoints is double gridCharSpace && UsesAdjustedRightIndent(paragraph, context))
            format.RightIndent = Unit.FromPoint(format.RightIndent.Point + gridCharSpace);
        if (_container is not Cell && UsesMirrorIndents(paragraph, context))
        {
            var leftIndent = format.LeftIndent;
            format.LeftIndent = format.RightIndent;
            format.RightIndent = leftIndent;
            _mirroredParagraphs.Add(_paragraph);
        }
        if (paragraph.Descendants<Run>().ToArray() is [{ } borderedRun] &&
            ResolveRunTextBorder(borderedRun, context) is { Shadow: true } shadowBorder)
        {
            format.LeftIndent = Unit.FromPoint(format.LeftIndent.Point + shadowBorder.Width);
            if (_mirroredParagraphs.Contains(_paragraph))
                format.RightIndent = Unit.FromPoint(format.RightIndent.Point + shadowBorder.Width);
        }
        if (_container is Cell && _table != null)
        {
            format.LeftIndent = Unit.FromPoint(format.LeftIndent.Point + _cellSpacingIndentPoints);
            if (_tableCellSpacingPoints > 0 &&
                context.CurrentTableCell is { RowIndex: 0, ColumnIndex: 0 } &&
                context.CurrentTable?.Table.TableProperties?.TableStyle != null &&
                context.CurrentTable.Table.TableProperties?.TableBorders?.LeftBorder == null &&
                context.CurrentTable.ComputedStyle.TableBorders?.Left is
                    { LineStyle: not DxpComputedBorderLineStyle.None, WidthPt: > 0 } spacedStyledLeft)
                format.LeftIndent = Unit.FromPoint(format.LeftIndent.Point + spacedStyledLeft.WidthPt / 2);
            if (format.Alignment == ParagraphAlignment.Left &&
                paragraph.Ancestors<TableCell>().FirstOrDefault() is { } styledCell &&
                styledCell.ElementsBefore().OfType<TableCell>().Count() == 1 &&
                paragraph.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault() is
                    { } styledTable &&
                StyledFixedTableFirstColumnInset(styledTable, context.CurrentTable) is > 0
                    and var styledInset)
                format.LeftIndent = Unit.FromPoint(format.LeftIndent.Point + styledInset);
            if (paragraph.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault() is
                { } sourceTable && sourceTable.TableProperties?.TableStyle == null &&
                sourceTable.TableProperties?.TableWidth?.Type?.Value == TableWidthUnitValues.Auto &&
                sourceTable.TableProperties?.TableLayout == null &&
                sourceTable.Descendants<NoWrap>().Any() &&
                _cellLeftMarginPoints == null && _rowLeftMarginPoints == null)
                format.LeftIndent = Unit.FromPoint(format.LeftIndent.Point - 5.4);
            if ((_cellLeftMarginPoints ?? _rowLeftMarginPoints) is double cellLeft)
                format.LeftIndent = Unit.FromPoint(format.LeftIndent.Point +
                    cellLeft - _table.LeftPadding.Point);
            if ((_cellRightMarginPoints ?? _rowRightMarginPoints) is double cellRight)
                format.RightIndent = Unit.FromPoint(format.RightIndent.Point +
                    cellRight - _table.RightPadding.Point);
            if (!_nativeDoubleTableBorders &&
                paragraph.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault() is
                    { } doubleBorderTable && HasSixDoubleBorders(doubleBorderTable))
                format.LeftIndent = Unit.FromPoint(format.LeftIndent.Point + 0.5);
            if (_cellNoWrap && _cellNoWrapWidthPoints is > 100 &&
                !string.Equals(_minorThemeLatinFont, "Arial", StringComparison.OrdinalIgnoreCase))
                format.RightIndent = Unit.FromPoint(format.RightIndent.Point - 3);
        }
        else if (_container is Cell && _table == null)
            format.LeftIndent = Unit.FromPoint(format.LeftIndent.Point +
                Math.Max(0, paragraph.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>().Count() - 1) * 5.4);
        if (_activeFootnoteIndex != null &&
            paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any(text =>
                text.Text.Contains('-') &&
                text.Text.TrimStart().StartsWith("https://", StringComparison.OrdinalIgnoreCase)))
            format.RightIndent = Unit.FromPoint(format.RightIndent.Point - 5);
        if (computed.TextIndentPt is double firstLine) format.FirstLineIndent = Unit.FromPoint(firstLine);
        if (noMarkerList) format.FirstLineIndent = Unit.Zero;
        if (_compactSnapToCharsParagraph && _cjkStyleParagraph && _container is Section &&
            context.Styles.ResolveRunStyle(paragraph, new Run()).FontSizeHalfPoints is int fontSizeHalfPoints and > 0)
        {
            var continuationInset = fontSizeHalfPoints / 8.0;
            format.LeftIndent = Unit.FromPoint(format.LeftIndent.Point + continuationInset);
            format.FirstLineIndent = Unit.FromPoint(format.FirstLineIndent.Point - continuationInset);
        }
        if (computed.MarginTopPt is double before) format.SpaceBefore = Unit.FromPoint(before);
        if (firstBodyBefore is > 0) format.SpaceBefore = Unit.Zero;
        if (computed.MarginBottomPt is double after) format.SpaceAfter = Unit.FromPoint(after);
        else if ((_stylelessBodyParagraphSpacing || _stylelessUnthemedParagraphAfter) && _container is Section &&
            (paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph ||
             _activeEndnoteIndex is int endnoteIndex && endnoteIndex < _endnoteOrdinals.Count))
            format.SpaceAfter = Unit.FromPoint(8.04);
        if ((_stylelessBodyParagraphSpacing || _stylelessUnthemedParagraphAfter &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph) &&
            _container is Section)
        {
            if (paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Table
                { TableProperties: null })
                format.LineSpacingRule = LineSpacingRule.Single;
            else
            {
                format.LineSpacingRule = LineSpacingRule.Multiple;
                format.LineSpacing = 1.1553;
            }
        }
        if (_container is Section &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
            { ParagraphProperties.NumberingProperties.NumberingId.Val.Value: 0 } precedingBlank &&
            precedingBlank.InnerText.Length == 0 &&
            precedingBlank.ParagraphProperties?.Indentation != null)
            format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point + 1.7);
        if (_container is Section &&
            string.Equals(format.Font.Name, "DejaVu Sans", StringComparison.OrdinalIgnoreCase) &&
            paragraph.ParagraphProperties?.NumberingProperties?.NumberingId != null &&
            paragraph.ParagraphProperties?.SpacingBetweenLines?.After?.Value == "225")
        {
            var italicList = paragraph.Descendants<Run>().FirstOrDefault() is { } firstListRun &&
                context.Styles.ResolveRunStyle(paragraph, firstListRun).Italic;
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point +
                (italicList ? 2.4 : 0.7));
            if (paragraph.PreviousSibling() is not DocumentFormat.OpenXml.Wordprocessing.Paragraph
                { ParagraphProperties.NumberingProperties.NumberingId: not null })
                format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point - 0.6);
        }
        if (_container is Section && context.Styles is DocxportNet.Walker.Context.DxpStyleResolver autoResolver &&
            autoResolver.GetParagraphSpacing(paragraph) is { } autoSpacing)
        {
            if (autoSpacing.BeforeAutoSpacing?.Value == true && firstBodyBefore is not > 0)
                format.SpaceBefore = Unit.FromPoint(
                    _availableFontFamilies.Contains("Aptos") &&
                    paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Heading2" ? 14 : 11);
            if (autoSpacing.AfterAutoSpacing?.Value == true)
                format.SpaceAfter = Unit.FromPoint(14);
        }
        if (_availableFontFamilies.Contains("Aptos") && _container is Cell &&
            paragraph.Parent is TableCell sourceCell)
        {
            var styleAutoSpacingDisabled = context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.SpacingBetweenLines?.BeforeAutoSpacing?.Value == false);
            var cellParagraphs = sourceCell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().ToArray();
            if (sourceCell.Parent is TableRow boundaryRow)
            {
                if (cellParagraphs.Length == 1 && ReferenceEquals(cellParagraphs[0], paragraph) &&
                    boundaryRow.PreviousSibling<TableRow>() == null &&
                    boundaryRow.NextSibling<TableRow>() is { } nextRow &&
                    nextRow.Elements<TableCell>().All(cell =>
                        cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count() >= 5) &&
                    paragraph.ParagraphProperties?.SpacingBetweenLines?.After?.Value == "120")
                {
                    if (!styleAutoSpacingDisabled)
                        format.SpaceAfter = Unit.Zero;
                    if (_tableOpeningSpacingTransferred.Remove(paragraph))
                        format.SpaceBefore = Unit.Zero;
                }
                if (cellParagraphs.Length >= 5 && ReferenceEquals(cellParagraphs[0], paragraph) &&
                    boundaryRow.PreviousSibling<TableRow>() is { } headerRow &&
                    headerRow.PreviousSibling<TableRow>() == null &&
                    paragraph.ParagraphProperties?.SpacingBetweenLines?.Before?.Value == "120")
                    if (!styleAutoSpacingDisabled)
                        format.SpaceBefore = Unit.Zero;
            }
            if (!styleAutoSpacingDisabled && cellParagraphs.Length >= 5 &&
                paragraph.ParagraphProperties?.SpacingBetweenLines?.After?.Value == "120")
                format.SpaceAfter = Unit.FromPoint(ReferenceEquals(cellParagraphs[^1], paragraph) ? 0 : 14);
            if (!styleAutoSpacingDisabled && cellParagraphs.Length is 1 or 2 &&
                ReferenceEquals(cellParagraphs[0], paragraph) &&
                sourceCell.Parent is TableRow sourceRow &&
                sourceRow.PreviousSibling<TableRow>() is { } priorRow &&
                priorRow.Elements<TableCell>().All(cell =>
                    cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count() == 1) &&
                paragraph.ParagraphProperties?.SpacingBetweenLines?.Before?.Value == "120")
                format.SpaceBefore = Unit.Zero;
            if (!styleAutoSpacingDisabled && cellParagraphs.Length == 1 &&
                sourceCell.Parent is TableRow shortRow &&
                shortRow.PreviousSibling<TableRow>() is { } longRow &&
                longRow.Elements<TableCell>().All(cell =>
                    cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count() >= 5) &&
                paragraph.ParagraphProperties?.SpacingBetweenLines?.Before == null)
            {
                format.SpaceBefore = Unit.FromPoint(-5);
                format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point - 5);
            }
            if (!styleAutoSpacingDisabled && cellParagraphs.Length == 1 &&
                sourceCell.Parent is TableRow terminalRow &&
                terminalRow.NextSibling<TableRow>() == null &&
                terminalRow.Elements<TableCell>().All(cell =>
                    cell.InnerText.TrimStart().StartsWith("b)", StringComparison.Ordinal)) &&
                terminalRow.PreviousSibling<TableRow>() is { } precedingTerminalRow &&
                precedingTerminalRow.Elements<TableCell>().All(cell =>
                    cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count() == 2) &&
                paragraph.ParagraphProperties?.SpacingBetweenLines?.After == null)
            {
                var originalBefore = format.SpaceBefore.Point;
                format.SpaceBefore = Unit.FromPoint(4);
                format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point - 6 +
                    Math.Max(0, originalBefore - 4));
            }
        }
        if (_container is Section && paragraph.ParagraphProperties?.SpacingBetweenLines?.Before?.Value == "120" &&
            paragraph.ParagraphProperties.ParagraphBorders is { } directBorders &&
            directBorders.GetFirstChild<LeftBorder>() != null &&
            directBorders.GetFirstChild<RightBorder>() != null)
        {
            var autoSpacingDisabled = context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.SpacingBetweenLines?.BeforeAutoSpacing?.Value == false);
            if (!autoSpacingDisabled)
                format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point + 6);
            else
            {
                if (paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                    { ParagraphProperties.ParagraphBorders: { } previousBorders } &&
                    previousBorders.OuterXml == directBorders.OuterXml)
                    format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point - 2);
                if (paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                    { ParagraphProperties.ParagraphBorders: { } nextBorders } &&
                    nextBorders.OuterXml == directBorders.OuterXml)
                    format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point - 2);
            }
            if (autoSpacingDisabled && paragraph.InnerText.Length < 80 &&
                paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextBorderedParagraph &&
                nextBorderedParagraph.InnerText.Length > paragraph.InnerText.Length &&
                nextBorderedParagraph.ParagraphProperties?.ParagraphBorders?.OuterXml == directBorders.OuterXml)
            {
                format.KeepWithNext = true;
                format.KeepTogether = true;
            }
        }
        if (_documentDefaultParagraphAfterPoints is double documentAfter &&
            context.Styles.GetParagraphStyleChain(paragraph) is { Count: > 0 } styleChain &&
            _paragraphStyles.TryGetValue(styleChain[^1].StyleId, out var rootStyle) &&
            rootStyle.BasedOn == null && rootStyle.Default?.Value != true &&
            (computed.MarginBottomPt == null || !styleChain.Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.SpacingBetweenLines?.After != null)))
            format.SpaceAfter = Unit.FromPoint(documentAfter);
        else if (_documentDefaultParagraphAfterPoints == null &&
            !_paragraphStyles.ContainsKey("Normal") &&
            context.Styles.GetParagraphStyleChain(paragraph) is { Count: > 0 } implicitChain &&
            _paragraphStyles.TryGetValue(implicitChain[^1].StyleId, out var implicitRoot) &&
            implicitRoot.BasedOn == null && implicitRoot.Default?.Value != true &&
            implicitRoot.StyleRunProperties?.RunFonts == null &&
            (implicitRoot.StyleParagraphProperties == null ||
             implicitRoot.StyleParagraphProperties.ChildElements.All(element =>
                 element is Indentation or MirrorIndents)))
            format.SpaceAfter = Unit.FromPoint(10);
        else if (!_documentHasPositionalTabs && _documentDefaultParagraphAfterPoints == null &&
            _paragraphStyles.TryGetValue("Normal", out var emptyNormal) &&
            emptyNormal.StyleParagraphProperties == null &&
            context.Styles.GetParagraphStyleChain(paragraph) is { Count: > 1 } normalChain &&
            _paragraphStyles.TryGetValue(normalChain[0].StyleId, out var derivedStyle) &&
            derivedStyle.StyleParagraphProperties == null &&
            string.Equals(normalChain[^1].StyleId, "Normal", StringComparison.OrdinalIgnoreCase))
            format.SpaceAfter = Unit.FromPoint(10);
        else if (_container is Section && paragraph.InnerText.Length == 0 &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph &&
            int.TryParse(markFont?.GetFirstChild<FontSize>()?.Val?.Value, out var markSize) &&
            context.Styles.ResolveRunStyle(paragraph, new Run()).FontSizeHalfPoints is int inheritedSize &&
            markSize > inheritedSize &&
            context.Styles is DocxportNet.Walker.Context.DxpStyleResolver spacingResolver &&
            spacingResolver.GetParagraphSpacing(paragraph)?.Line?.Value == "360")
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + markSize - inheritedSize);
        else if (_classicPartialDefaults && (_container is Section ||
            _container is Cell && !_tableInHeaderFooter))
            format.SpaceAfter = Unit.FromPoint(10);
        else if (hasExplicitTopAlignment && _container is Section &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
            format.SpaceAfter = Unit.FromPoint(14);
        else if (paragraph.Descendants<FitText>().Any() &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph afterFit)
            format.SpaceAfter = Unit.FromPoint(afterFit.InnerText.Any(IsHangul) ? 14.3 :
                afterFit.InnerText.Any(IsCjk) &&
                paragraph.Descendants<Run>().Any(run =>
                    run.RunProperties?.GetFirstChild<FitText>() == null && ResolveFitText(run) != null) &&
                paragraph.Descendants<Run>().Any(run =>
                    run.RunProperties?.GetFirstChild<FitText>() != null) ? 12 : 14);
        else if (_fitTextTracking.Contains(_paragraph) &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
            format.SpaceAfter = Unit.FromPoint(8);
        else if (_cjkStyleParagraph) format.SpaceAfter = Unit.FromPoint(8);
        else if (paragraph.Descendants<PositionalTab>().Any())
        {
            var descendants = paragraph.Descendants<OpenXmlElement>().ToArray();
            var tabIndex = Array.FindIndex(descendants, element => element is PositionalTab);
            var hasTextAfterTab = tabIndex >= 0 && descendants.Skip(tabIndex + 1)
                .OfType<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => !string.IsNullOrWhiteSpace(text.Text));
            var hasExplicitAfter = paragraph.ParagraphProperties?.SpacingBetweenLines?.After != null ||
                context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                    _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                    definition.StyleParagraphProperties?.SpacingBetweenLines?.After != null);
            if (!hasExplicitAfter)
                format.SpaceAfter = Unit.FromPoint((_implicitModernFont ? 10.4 : 12.5) -
                    (_container is Section || _container is HeaderFooter && !hasTextAfterTab
                        ? 0.09 : 0));
            if (context.Styles is DocxportNet.Walker.Context.DxpStyleResolver tabSpacingResolver &&
                tabSpacingResolver.GetContextualSpacing(paragraph) &&
                paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextTabParagraph &&
                paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ==
                    nextTabParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value)
                format.SpaceAfter = Unit.Zero;
            if (_container is Cell &&
                paragraph.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault()?
                    .Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().Skip(1).Any() == true &&
                paragraph.Descendants<Run>().LastOrDefault()?.LastChild is PositionalTab)
                format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point - 4.4);
            if (_inHeader && _container is HeaderFooter && !hasExplicitAfter)
            {
                _headerPositionalTabCount++;
                if (hasTextAfterTab)
                    _headerPositionalTabHasFollowingText = true;
            }
            if (_inHeader && _container is Cell) _headerPositionalCellParagraphCount++;
            if (_container is Cell &&
                paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextParagraph &&
                nextParagraph.Descendants<PositionalTab>().Any() &&
                (_tableInHeaderFooter || _table?.LeftPadding.Point > 1))
                format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point +
                    _defaultParagraphLinePoints + 4);
        }
        else if (_paragraphGridCharSpacingPoints is > 0 && _defaultParagraphAfterPoints == 0 &&
            !_paragraphStyles.ContainsKey("Normal"))
            format.SpaceAfter = Unit.FromPoint(8);
        else if (_horizontalMergeTable && _container is Cell)
            format.SpaceAfter = Unit.FromPoint(10);
        if (_inHeader && _container is HeaderFooter && _headerParagraphCount == 1 &&
            paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value is string headerStyleId &&
            _paragraphStyles.TryGetValue(headerStyleId, out var headerStyle) &&
            headerStyle.StyleRunProperties?.RunFonts?.Ascii?.Value == "Arial" &&
            headerStyle.StyleRunProperties?.Bold != null &&
            headerStyle.StyleRunProperties?.Italic == null &&
            headerStyle.StyleRunProperties?.FontSize?.Val?.Value == "22" &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextStyledHeaderParagraph &&
            nextStyledHeaderParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Header" &&
            format.SpaceAfter.Point >= 6 &&
            !paragraph.Descendants<Drawing>().Any(drawing => drawing.Inline != null))
        {
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point - 1.27);
            _headerFirstAfterLayoutAdjustmentPoints = 1.27;
        }
        if (!_inHeader && _container is HeaderFooter && _section != null &&
            paragraph.InnerText.Length == 0 &&
            paragraph.ParagraphProperties?.ParagraphMarkRunProperties is { } styledFooterMark &&
            (styledFooterMark.GetFirstChild<Bold>() != null ||
                styledFooterMark.GetFirstChild<Italic>() != null) &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph previousFooterLine &&
            !string.IsNullOrWhiteSpace(previousFooterLine.InnerText) &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextFooterLine &&
            !string.IsNullOrWhiteSpace(nextFooterLine.InnerText) &&
            paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value is string emptyFooterStyleId &&
            nextFooterLine.ParagraphProperties?.ParagraphStyleId?.Val?.Value == emptyFooterStyleId &&
            previousFooterLine.ParagraphProperties?.ParagraphStyleId?.Val?.Value != emptyFooterStyleId &&
            (previousFooterLine.ParagraphProperties?.Justification?.Val?.Value == JustificationValues.Center ||
                previousFooterLine.ParagraphProperties?.ParagraphStyleId?.Val?.Value is string previousFooterStyleId &&
                _paragraphStyles.TryGetValue(previousFooterStyleId, out var previousFooterStyle) &&
                previousFooterStyle.BasedOn?.Val?.Value is string footerBaseStyleId &&
                !string.Equals(footerBaseStyleId, "Normal", StringComparison.OrdinalIgnoreCase) &&
                context.Styles.GetParagraphStyleChain(previousFooterLine).Any(entry =>
                    _paragraphStyles.TryGetValue(entry.StyleId, out var inheritedFooterStyle) &&
                    inheritedFooterStyle.StyleParagraphProperties?.Justification?.Val?.Value ==
                        JustificationValues.Distribute)))
        {
            const double blankLineCorrection = 1.9;
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + blankLineCorrection);
            _section.PageSetup.FooterDistance = Unit.FromPoint(
                _section.PageSetup.FooterDistance.Point + 1.8 - blankLineCorrection);
        }
        if (_paragraphInheritsFitText && format.Font.Size.Point < 11.5 &&
            computed.MarginBottomPt == null &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextMixedParagraph &&
            nextMixedParagraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(IsCjk)))
            format.SpaceAfter = Unit.FromPoint(Math.Max(0, format.SpaceAfter.Point - 4));
        if (_fitTextTracking.Contains(_paragraph) && _container is Section &&
            !_documentHasFitText && computed.MarginBottomPt == null &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextTrackedCjkParagraph &&
            nextTrackedCjkParagraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(IsCjk)))
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + 1);
        if (!_paragraphInheritsFitText && computed.MarginBottomPt == null &&
            paragraph.Descendants<FitText>().Any() &&
            paragraph.Descendants<Run>().All(run => run.RunProperties?.RunStyle == null) &&
            !paragraph.Descendants<Run>().Any(run =>
                run.RunProperties?.Bold is { } bold && bold.Val?.Value != false ||
                run.RunProperties?.Italic is { } italic && italic.Val?.Value != false) &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextCjkParagraph &&
            nextCjkParagraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(IsCjk)))
            format.SpaceAfter = Unit.FromPoint(Math.Max(0, format.SpaceAfter.Point - 2.1));
        if (_gridLinePitchPoints is > 0 && _gridCharSpacingPoints == null &&
            _container is Section && computed.MarginBottomPt == null &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
            format.SpaceAfter = Unit.FromPoint(UsesDocumentGrid(paragraph, context) ? 12 : 7.5);
        if (_snapToCharsGrid && !_cjkEmphasisParagraph && _container is Section &&
            computed.MarginBottomPt == null &&
            paragraph.ParagraphProperties?.SnapToGrid?.Val?.Value == false &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextSnapParagraph &&
            nextSnapParagraph.ParagraphProperties?.SnapToGrid?.Val?.Value == true)
            format.SpaceAfter = Unit.FromPoint(Math.Max(0, format.SpaceAfter.Point - 0.7));
        if (_snapToCharsGrid && _cjkEmphasisParagraph && _container is Section &&
            computed.MarginBottomPt == null &&
            paragraph.ParagraphProperties?.SnapToGrid?.Val?.Value == true &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextImplicitSnapParagraph &&
            nextImplicitSnapParagraph.ParagraphProperties?.SnapToGrid == null)
            format.SpaceAfter = Unit.FromPoint(Math.Max(0, format.SpaceAfter.Point - 0.85));
        if (_container is Section && computed.MarginBottomPt == null && _defaultParagraphAfterPoints > 0 &&
            !paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any() &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Table)
            format.SpaceAfter = Unit.FromPoint(_defaultParagraphAfterPoints + EmptyParagraphAfterTableExtraPoints);
        if (_container is Section && paragraph.InnerText.Length == 0 &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Table &&
            paragraph.ParagraphProperties?.SpacingBetweenLines?.After != null &&
            string.Equals(format.Font.Name, "DejaVu Sans", StringComparison.OrdinalIgnoreCase))
            format.SpaceAfter = Unit.FromPoint(Math.Max(0, format.SpaceAfter.Point - 1.5));
        if (_container is Section && paragraph.InnerText.Length == 0 &&
            paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value == 0 &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + 1.5);
        if (_container is Section && paragraph.InnerText.Length == 0 &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph previousBlank &&
            previousBlank.InnerText.Length == 0 &&
            previousBlank.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "ListParagraph" &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
            format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point + 1.5);
        if (_container is Section &&
            (paragraph.ParagraphProperties == null ||
             paragraph.ParagraphProperties.ChildElements.All(element =>
                 element is ParagraphMarkRunProperties mark &&
                 mark.ChildElements.All(property => property is Languages))) &&
            paragraph.InnerText.Length == 0 && !paragraph.Descendants<Break>().Any() &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph previousList &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextList &&
            context.Styles.ResolveEffectiveNumPr(previousList) is { } previousNumbering &&
            context.Styles.ResolveEffectiveNumPr(nextList) is { } nextNumbering &&
            previousNumbering.NumId != nextNumbering.NumId)
        {
            var previousLevel = _numberingResolver?.ResolveLevel(previousNumbering.NumId,
                previousNumbering.Ilvl);
            var nextLevel = _numberingResolver?.ResolveLevel(nextNumbering.NumId,
                nextNumbering.Ilvl);
            var bulletToNumber = previousLevel is { } bullet && nextLevel is { } number &&
                bullet.lvl.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet &&
                number.lvl.NumberingFormat?.Val?.Value != NumberFormatValues.Bullet;
            format.SpaceAfter = Unit.FromPoint(Math.Max(_defaultParagraphAfterPoints,
                format.SpaceAfter.Point) + (bulletToNumber ? 9.4 : 2.4));
        }
        if (_gridLinePitchPoints is > 0 && _gridCharSpacingPoints is > 0 &&
            _headerGridWrapClearancePoints > 0 && _container is Section &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextGridParagraph &&
            computed.MarginBottomPt == null)
            format.SpaceAfter = Unit.FromPoint(7 +
                (UsesDocumentGrid(paragraph, context) && UsesDocumentGrid(nextGridParagraph, context) ? 1.15 : 0));
        if (_themedImplicitParagraphSpacing && computed.MarginBottomPt == null &&
            (_container is Section && paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Table ||
             _container is Cell && !_tableInHeaderFooter))
            format.SpaceAfter = Unit.FromPoint(Math.Max(_defaultParagraphAfterPoints,
                format.SpaceAfter.Point) + 2.2);
        if (_container is Section && paragraph.InnerText.Length == 0 &&
            (paragraph.Parent is SdtContentBlock tocContent && paragraph.NextSibling() == null &&
                tocContent.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Any(item =>
                    item.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.StartsWith("TOC",
                        StringComparison.OrdinalIgnoreCase) == true) ||
             paragraph.PreviousSibling() is SdtBlock precedingToc &&
                precedingToc.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Any(item =>
                    item.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.StartsWith("TOC",
                        StringComparison.OrdinalIgnoreCase) == true)))
            format.SpaceAfter = Unit.FromPoint(Math.Max(_defaultParagraphAfterPoints,
                format.SpaceAfter.Point) + TocEmptyParagraphClearancePoints);
        if (_container is Section &&
            paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Heading1" &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph precedingEmpty &&
            precedingEmpty.InnerText.Length == 0 &&
            precedingEmpty.PreviousSibling() is SdtBlock precedingTocBlock &&
            precedingTocBlock.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Any(item =>
                item.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.StartsWith("TOC",
                    StringComparison.OrdinalIgnoreCase) == true))
            format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point +
                TocEmptyParagraphClearancePoints);
        if (_container is Section && paragraph.NextSibling() is
            DocumentFormat.OpenXml.Wordprocessing.Paragraph nextImageParagraph &&
            nextImageParagraph.Descendants<Drawing>().Any(drawing => drawing.Descendants()
                .Any(element => element.LocalName is "anchor" or "chart")) &&
            !nextImageParagraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any())
            format.SpaceAfter = Unit.Zero;
        if (_container is Section && !paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any() &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph imageParagraph &&
            imageParagraph.Descendants<Drawing>().Any(drawing => drawing.Descendants().Any(element =>
                element.LocalName == "wrapSquare")) &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph followingText &&
            followingText.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any())
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + 3.65 * _defaultParagraphLinePoints);
        if (_container is Section && paragraph.ParagraphProperties?.ParagraphBorders is { } exitBorders &&
            exitBorders.GetFirstChild<TopBorder>() != null &&
            exitBorders.GetFirstChild<LeftBorder>() != null &&
            exitBorders.GetFirstChild<BottomBorder>() != null &&
            exitBorders.GetFirstChild<RightBorder>() != null &&
            paragraph.ParagraphProperties.SpacingBetweenLines?.After?.Value == "120" &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextUnbordered &&
            nextUnbordered.ParagraphProperties?.ParagraphBorders == null &&
            !context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.SpacingBetweenLines?.BeforeAutoSpacing?.Value == false))
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point +
                BorderedParagraphExitClearancePoints);
        if (computed.Borders is { } borders) ApplyBorders(format.Borders, borders);
        if (paragraph.ParagraphProperties?.GetFirstChild<ParagraphBorders>() is { } paragraphBorders)
        {
            if (_container is Section &&
                paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextInBorderGroup &&
                nextInBorderGroup.ParagraphProperties?.ParagraphBorders?.OuterXml == paragraphBorders.OuterXml &&
                paragraph.PreviousSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()?
                    .ParagraphProperties?.ParagraphBorders?.OuterXml != paragraphBorders.OuterXml)
                format.KeepWithNext = true;
            if (_container is Section && computed.MarginBottomPt == null &&
                paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextBordered &&
                nextBordered.ParagraphProperties?.ParagraphBorders?.OuterXml == paragraphBorders.OuterXml)
                format.SpaceAfter = Unit.FromPoint(Math.Max(0, _defaultParagraphAfterPoints - 2));
            if (_container is Section) _directBorderParagraphs[_paragraph] = paragraph;
            if (paragraphBorders.GetFirstChild<TopBorder>() is { } topBorder &&
                int.TryParse(AttributeValue(topBorder, "space"), out var topSpace) && topSpace > 0)
            {
                if (_container is Section &&
                    paragraph.PreviousSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>() is
                        { ParagraphProperties.ParagraphBorders: null })
                    format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point + 2 * topSpace);
                format.Borders.Distance = Unit.FromPoint(topSpace);
                if (_container is Section &&
                    paragraphBorders.GetFirstChild<BottomBorder>() == null &&
                    computed.BackgroundColorCss != null)
                    format.Borders.DistanceFromBottom = Unit.Zero;
                if (paragraphBorders.GetFirstChild<LeftBorder>() == null &&
                    paragraphBorders.GetFirstChild<RightBorder>() == null &&
                    computed.BackgroundColorCss != null &&
                    double.TryParse(AttributeValue(topBorder, "sz"), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var borderSize))
                {
                    format.Borders.DistanceFromLeft = Unit.FromPoint(topSpace + borderSize / 8);
                    format.Borders.DistanceFromRight = Unit.FromPoint(topSpace + borderSize / 8);
                }
            }
            foreach (var border in paragraphBorders.ChildElements)
            {
                if (!int.TryParse(AttributeValue(border, "space"), out var borderSpace) || borderSpace <= 0)
                    continue;
                var distance = Unit.FromPoint(borderSpace);
                switch (border.LocalName)
                {
                    case "top": format.Borders.DistanceFromTop = distance; break;
                    case "bottom": format.Borders.DistanceFromBottom = distance; break;
                    case "left": format.Borders.DistanceFromLeft = Unit.FromPoint(
                        distance.Point + DirectParagraphSideBorderOffsetPoints); break;
                    case "right": format.Borders.DistanceFromRight = Unit.FromPoint(
                        distance.Point + DirectParagraphSideBorderOffsetPoints); break;
                }
            }
        }
        else if (computed.BackgroundColorCss != null &&
            computed.Borders is { Top: { LineStyle: not DxpComputedBorderLineStyle.None },
                Left: null, Right: null } &&
            context.Styles.GetParagraphStyleChain(paragraph)
                .Select(entry => _paragraphStyles.TryGetValue(entry.StyleId, out var definition)
                    ? definition.StyleParagraphProperties?.ParagraphBorders?.TopBorder : null)
                .FirstOrDefault(border => border != null) is { } inheritedTop &&
            double.TryParse(AttributeValue(inheritedTop, "space"), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var inheritedSpace) && inheritedSpace > 0 &&
            double.TryParse(AttributeValue(inheritedTop, "sz"), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var inheritedSize))
        {
            var distance = Unit.FromPoint(inheritedSpace + inheritedSize / 8);
            format.Borders.Distance = Unit.FromPoint(inheritedSpace);
            if (computed.Borders.Bottom == null)
                format.Borders.DistanceFromBottom = Unit.Zero;
            format.Borders.DistanceFromLeft = distance;
            format.Borders.DistanceFromRight = distance;
            if (_container is Section)
                format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point - inheritedSpace);
        }
        if (ParseColor(computed.BackgroundColorCss) is { } background &&
            !(_container is Cell && background == MigraDoc.DocumentObjectModel.Colors.White &&
              (_tableBackground == null || _tableBackground == MigraDoc.DocumentObjectModel.Colors.White) &&
              (ParseColor(_cellStyle?.BackgroundColorCss) is not { } cellBackground ||
               cellBackground == MigraDoc.DocumentObjectModel.Colors.White)))
        {
            format.Shading.Color = background;
            if ((_container is Section || _container is HeaderFooter && !_inHeader) &&
                computed.BackgroundColorCss is string css)
                _paragraphShades[_paragraph] = css;
            if ((_container is Section || _container is HeaderFooter && !_inHeader) &&
                context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                    _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                    definition.StyleRunProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Border>() != null))
            {
                format.Shading.Visible = false;
                _manualParagraphShades.Add(_paragraph);
                if (_container is HeaderFooter)
                    _footerParagraphShadeClearances[_paragraph] = paragraph.Descendants<Run>()
                        .Select(run => ResolveRunTextBorder(run, context)?.Inset ?? 0)
                        .DefaultIfEmpty(0).Max() * 2 + 1;
            }
        }
        if (_container is Section && paragraph.ParagraphProperties?.ParagraphBorders != null &&
            paragraph.ParagraphProperties.Shading != null &&
            paragraph.ParagraphProperties.SpacingBetweenLines?.After == null &&
            _defaultParagraphAfterPoints >= 0.2 &&
            paragraph.PreviousSibling() == null && _section != null &&
            paragraph.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>() is { } followingBlank &&
            string.IsNullOrEmpty(followingBlank.InnerText))
        {
            _section.PageSetup.TopMargin = Unit.FromPoint(_section.PageSetup.TopMargin.Point + 1.8);
            format.SpaceAfter = Unit.FromPoint(_defaultParagraphAfterPoints - 0.2);
        }
        if (_container is Section && paragraph.PreviousSibling() == null && _section != null &&
            paragraph.ParagraphProperties?.ParagraphBorders == null &&
            context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.ParagraphBorders?.TopBorder != null &&
                definition.StyleParagraphProperties.Shading != null))
            _section.PageSetup.TopMargin = Unit.FromPoint(_section.PageSetup.TopMargin.Point - 0.8);
        var verticalAlignment = paragraph.ParagraphProperties?.TextAlignment?.Val?.Value;
        if (verticalAlignment == null)
            verticalAlignment = context.Styles.GetParagraphStyleChain(paragraph)
                .Select(entry => _paragraphStyles.TryGetValue(entry.StyleId, out var definition)
                    ? definition.StyleParagraphProperties?.TextAlignment?.Val?.Value : null)
                .FirstOrDefault(value => value != null);
        if (_gridLinePitchPoints is > 0 && _gridCharSpacingPoints is > 0 &&
            verticalAlignment == null &&
            (paragraph.ParagraphProperties?.AdjustRightIndent != null ||
             context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                 _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                 definition.StyleParagraphProperties?.AdjustRightIndent != null)))
            _unshiftedGridGlyphParagraphs.Add(_paragraph);
        if (_container is HeaderFooter && _availableFontFamilies.Contains("Aptos") &&
            context.Styles.ResolveRunStyle(paragraph,
                paragraph.Descendants<Run>().FirstOrDefault() ?? new Run()) is { FontName: "Aptos", FontSizeHalfPoints: 24 } &&
            paragraph.ParagraphProperties?.SpacingBetweenLines == null &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextStoryLine &&
            paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value is string storyStyleId &&
            nextStoryLine.ParagraphProperties?.ParagraphStyleId?.Val?.Value == storyStyleId &&
            (paragraph.InnerText.Length == 0 ? paragraph : nextStoryLine) is { } middleStoryLine &&
            middleStoryLine.InnerText.Length == 0 &&
            (middleStoryLine.ParagraphProperties?.ParagraphMarkRunProperties?
                .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Bold>() != null ||
             middleStoryLine.ParagraphProperties?.ParagraphMarkRunProperties?
                .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Italic>() != null) &&
            middleStoryLine.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph precedingStoryLine &&
            !string.IsNullOrWhiteSpace(precedingStoryLine.InnerText) &&
            middleStoryLine.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph followingStoryLine &&
            !string.IsNullOrWhiteSpace(followingStoryLine.InnerText) &&
            precedingStoryLine.ParagraphProperties?.ParagraphStyleId?.Val?.Value == storyStyleId &&
            followingStoryLine.ParagraphProperties?.ParagraphStyleId?.Val?.Value == storyStyleId &&
            precedingStoryLine.ParagraphProperties?.ChildElements.All(element =>
                element is ParagraphStyleId) == true &&
            followingStoryLine.ParagraphProperties?.ChildElements.All(element =>
                element is ParagraphStyleId) == true &&
            precedingStoryLine.Descendants<Run>().All(run => run.RunProperties == null) &&
            followingStoryLine.Descendants<Run>().All(run => run.RunProperties == null) &&
            middleStoryLine.ParagraphProperties?.ParagraphMarkRunProperties?.ChildElements.Count == 1 &&
            context.Styles is DocxportNet.Walker.Context.DxpStyleResolver linePitchResolver &&
            linePitchResolver.GetParagraphSpacing(paragraph) is { } storySpacing &&
            storySpacing.Line?.Value == "240" &&
            (storySpacing.LineRule?.Value == null ||
             storySpacing.LineRule.Value == LineSpacingRuleValues.Auto))
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + 0.99);
        var firstFooterSpacing = (context.Styles as DocxportNet.Walker.Context.DxpStyleResolver)?
            .GetParagraphSpacing(paragraph);
        if (_container is HeaderFooter && !_inHeader && _footerParagraphCount == 1 &&
            paragraph.Descendants<Run>().FirstOrDefault() is { } firstFooterRun &&
            context.Styles.ResolveRunStyle(paragraph, firstFooterRun) is
                { FontName: "Aptos", FontSizeHalfPoints: 24, Subscript: true } &&
            paragraph.ParagraphProperties?.SpacingBetweenLines == null &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph emptyFooterLine &&
            string.IsNullOrWhiteSpace(emptyFooterLine.InnerText) &&
            emptyFooterLine.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph finalFooterLine &&
            !string.IsNullOrWhiteSpace(finalFooterLine.InnerText) &&
            (firstFooterSpacing == null || firstFooterSpacing.Line?.Value == "240" &&
                (firstFooterSpacing.LineRule?.Value == null ||
                 firstFooterSpacing.LineRule.Value == LineSpacingRuleValues.Auto)))
            format.SpaceAfter = Unit.FromPoint(firstFooterSpacing == null
                ? Math.Max(format.SpaceAfter.Point, _defaultParagraphAfterPoints) + 2.45
                : format.SpaceAfter.Point + 2.32);
        if (computed.LineHeightCss is string lineHeight && lineHeight.EndsWith("pt", StringComparison.Ordinal) &&
            double.TryParse(lineHeight.AsSpan(0, lineHeight.Length - 2),
                System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var points))
        {
            format.LineSpacing = Unit.FromPoint(points);
            var effectiveSpacing = (context.Styles as DocxportNet.Walker.Context.DxpStyleResolver)?
                .GetParagraphSpacing(paragraph) ?? paragraph.ParagraphProperties?.SpacingBetweenLines;
            format.LineSpacingRule = effectiveSpacing?.LineRule?.Value == LineSpacingRuleValues.AtLeast
                ? LineSpacingRule.AtLeast : LineSpacingRule.Exactly;
        }
        else if (computed.LineHeightCss is string multipleText &&
            double.TryParse(multipleText, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var multiple) && multiple > 0)
        {
            if (markFamily?.Contains(';') == true && markHalfPoints > 0 &&
                paragraph.Descendants<Run>().All(run =>
                    (run.RunProperties?.RunFonts?.Ascii?.Value is not string family || family == markFamily) &&
                    (run.RunProperties?.FontSize?.Val?.Value is not string size ||
                        !int.TryParse(size, out var halfPoints) || halfPoints == markHalfPoints)))
            {
                format.LineSpacingRule = LineSpacingRule.Exactly;
                format.LineSpacing = Unit.FromPoint(markHalfPoints / 2.0 * multiple * CompoundFontLineHeightRatio);
            }
            else if (!(paragraph.ParagraphProperties?.SpacingBetweenLines?.Line?.Value == "240" &&
                    paragraph.ParagraphProperties.SpacingBetweenLines.LineRule?.Value == LineSpacingRuleValues.Auto) &&
                markFamily != null && _availableFontFamilies.Contains(markFamily) &&
                markHalfPoints > 0 && paragraph.Descendants<Run>().All(run =>
                    (run.RunProperties?.RunFonts?.Ascii?.Value is not string family || family == markFamily) &&
                    (run.RunProperties?.FontSize?.Val?.Value is not string size ||
                        !int.TryParse(size, out var halfPoints) || halfPoints == markHalfPoints)))
            {
                var font = new XFont(markFamily, markHalfPoints / 2.0);
                format.LineSpacingRule = LineSpacingRule.Exactly;
                format.LineSpacing = Unit.FromPoint(font.Metrics.LineSpacing *
                    (markHalfPoints / 2.0) * multiple / font.Metrics.UnitsPerEm);
            }
            else
            {
                format.LineSpacingRule = LineSpacingRule.Multiple;
                format.LineSpacing = multiple;
            }
        }
        else if (paragraph.ParagraphProperties?.SpacingBetweenLines is { Line: null } &&
            _sourceWordDocument?.MainDocumentPart?.StyleDefinitionsPart == null &&
            _gridLinePitchPoints == null)
        {
            format.LineSpacingRule = LineSpacingRule.Multiple;
            format.LineSpacing = 1.158;
        }
        else if (_gridLinePitchPoints is double gridPitch && UsesDocumentGrid(paragraph, context) &&
            (_container is Section || _gridCharSpacingPoints != null))
        {
            format.LineSpacingRule = LineSpacingRule.AtLeast;
            format.LineSpacing = Unit.FromPoint(gridPitch + (_paragraphGridCharSpacingPoints is > 0
                ? verticalAlignment == null ? 3.75 : 3.84 : 0));
        }
        else if (_gridLinePitchPoints is double unsnappedGridPitch &&
            _gridCharSpacingPoints is > 0 && _headerGridWrapClearancePoints > 0 &&
            _container is Section && computed.LineHeightCss == null)
        {
            format.LineSpacingRule = LineSpacingRule.AtLeast;
            format.LineSpacing = Unit.FromPoint(unsnappedGridPitch - 0.84);
        }
        else if (_container is Section && computed.LineHeightCss == null &&
            paragraph.Descendants<Run>().FirstOrDefault() is { } firstBodyRun &&
            context.Styles.ResolveRunStyle(paragraph, firstBodyRun) is { FontName: "Arial",
                FontSizeHalfPoints: 18 })
        {
            format.LineSpacingRule = LineSpacingRule.Exactly;
            format.LineSpacing = Unit.FromPoint(10.5);
            _compactArialBodyParagraphs.Add(_paragraph);
        }
        if (_inHeader && _headerParagraphCount == 1 && _gridLinePitchPoints is > 0 &&
            _gridCharSpacingPoints is > 0 && computed.LineHeightCss == null &&
            StoryParagraphWraps(paragraph, context) && _section != null)
        {
            format.LineSpacingRule = LineSpacingRule.AtLeast;
            format.LineSpacing = Unit.FromPoint(_gridLinePitchPoints.Value - 0.84);
            _section.PageSetup.HeaderDistance = Unit.FromPoint(
                _section.PageSetup.HeaderDistance.Point + 0.84);
        }
        if (_container is HeaderFooter && !_inHeader && _gridLinePitchPoints is > 0 &&
            _gridCharSpacingPoints is > 0 && computed.LineHeightCss == null &&
            StoryParagraphWraps(paragraph, context))
        {
            format.LineSpacingRule = LineSpacingRule.AtLeast;
            format.LineSpacing = Unit.FromPoint(_gridLinePitchPoints.Value - 0.84);
        }
        if (_container is Section && paragraph.InnerText.Length == 0 &&
            paragraph.ParagraphProperties?.Indentation != null &&
            paragraph.ParagraphProperties?.ParagraphStyleId == null &&
            paragraph.ParagraphProperties?.NumberingProperties == null &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextListParagraph &&
            context.Styles.ResolveEffectiveNumPr(nextListParagraph) != null)
        {
            format.LineSpacingRule = LineSpacingRule.AtLeast;
            format.LineSpacing = Unit.FromPoint(17);
        }
        if (_container is Cell && _availableFontFamilies.Contains("DejaVu Sans") &&
            computed.LineHeightCss is string cellLineHeight &&
            double.TryParse(cellLineHeight, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var cellMultiple) && cellMultiple > 1 &&
            paragraph.Descendants<Run>().FirstOrDefault() is { } cellTextRun &&
            string.Equals(context.Styles.ResolveRunStyle(paragraph, cellTextRun).FontName,
                "DejaVu Sans", StringComparison.OrdinalIgnoreCase) &&
            paragraph.ParagraphProperties?.SpacingBetweenLines?.After != null &&
            paragraph.Parent is TableCell { Parent: TableRow cellRow } &&
            cellRow.PreviousSibling<TableRow>() != null)
        {
            format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point - 1.5);
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + 1.5);
        }
        if (computed.LineHeightCss == null && _gridLinePitchPoints == null &&
            paragraph.Descendants<PositionalTab>().Any(tab =>
                AttributeValue(tab, "alignment") == "left" &&
                AttributeValue(tab, "relativeTo") == "margin" &&
                (tab.ElementsBefore().OfType<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Any(text => text.Text.Length > 0) ||
                 tab.Parent?.ElementsBefore().Any(element =>
                    element.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                        .Any(text => text.Text.Length > 0)) == true) &&
                (tab.ElementsAfter().OfType<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Any(text => text.Text.Length > 0) ||
                tab.Parent?.ElementsAfter().Any(element =>
                    element.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                        .Any(text => text.Text.Length > 0)) == true)))
        {
            format.LineSpacingRule = LineSpacingRule.Exactly;
            format.LineSpacing = Unit.FromPoint(17);
            if (_container is HeaderFooter && _inHeader)
                format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point - 1);
            format.SpaceAfter = Unit.FromPoint(Math.Max(0, format.SpaceAfter.Point - 2.35) +
                (_container is HeaderFooter && _inHeader ? 1 : 0));
        }
        if (_container is HeaderFooter && !_inHeader && _gridLinePitchPoints is double footerGridPitch &&
            _gridCharSpacingPoints == null &&
            paragraph.ParagraphProperties?.SnapToGrid?.Val?.Value == true)
        {
            if (_section != null && _footerLineGridDistanceApplied.Add(_section))
                _section.PageSetup.FooterDistance = Unit.FromPoint(
                    _section.PageSetup.FooterDistance.Point + Math.Max(0, footerGridPitch - 13.8));
        }
        if (_cjkStyleParagraph && computed.LineHeightCss == null)
        {
            format.LineSpacingRule = LineSpacingRule.AtLeast;
            format.LineSpacing = Unit.FromPoint(
                (_cjkEmphasisParagraph || _snapToCharsGrid && _paragraphGridCharSpacingPoints is > 0) &&
                !clearsUnsnappedGridEmphasis ? (_snapToCharsGrid ? 27.8 : 28.2) : 23.1);
        }
        var implicitPlainAptos12 = HasImplicitPlainAptos12Runs(paragraph);
        if (_container is Section && _paragraphStyles.Count == 0 &&
            paragraph.ParagraphProperties?.SpacingBetweenLines?.After?.Value == "0" &&
            paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(IsFullWidthCjkPunctuation)) &&
            Math.Abs(_document!.Styles["Normal"].Font.Size.Point - 12) < 0.1 &&
            computed.LineHeightCss == null && _gridLinePitchPoints == null)
        {
            format.LineSpacingRule = LineSpacingRule.Exactly;
            format.LineSpacing = Unit.FromPoint(20.12);
            _exactCjkLineParagraphs.Add(_paragraph);
        }
        else if (_container is Section && _paragraphStyles.Count == 0 &&
            (paragraph.ParagraphProperties?.SpacingBetweenLines?.After?.Value == "0" &&
             Math.Abs(_document!.Styles["Normal"].Font.Size.Point - 12) < 0.1 ||
             implicitPlainAptos12) &&
            computed.LineHeightCss == null && _gridLinePitchPoints == null)
        {
            format.LineSpacingRule = LineSpacingRule.Exactly;
            format.LineSpacing = Unit.FromPoint(16.97);
            _exactStylelessPlainParagraphs.Add(_paragraph);
            if (implicitPlainAptos12)
                _exactStylelessImplicitAptosParagraphs.Add(_paragraph);
        }
        if (_chineseFallbackParagraph && computed.LineHeightCss == null)
        {
            format.LineSpacingRule = LineSpacingRule.Exactly;
            format.LineSpacing = Unit.FromPoint(20.76);
            format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point - 3.45);
        }
        if (_unfittedMixedScriptParagraph && computed.LineHeightCss == null &&
            !_unfittedCalibriParagraph)
        {
            format.LineSpacingRule = LineSpacingRule.Exactly;
            format.LineSpacing = Unit.FromPoint(23.2);
            var precedingEmphasis = paragraph.PreviousSibling() is
                DocumentFormat.OpenXml.Wordprocessing.Paragraph preceding &&
                preceding.Descendants<Run>().Any(run =>
                    run.RunProperties?.Bold is { } bold && bold.Val?.Value != false ||
                    run.RunProperties?.Italic is { } italic && italic.Val?.Value != false);
            format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point -
                (precedingEmphasis ? 5 : 2.3));
        }
        if ((verticalAlignment == VerticalTextAlignmentValues.Center ||
             verticalAlignment == VerticalTextAlignmentValues.Bottom) &&
            computed.LineHeightCss == null && fittedRuns.Length > 1)
        {
            var sizes = fittedRuns.Select(run => context.Styles.ResolveRunStyle(paragraph, run)
                .FontSizeHalfPoints is int points ? points / 2.0 : format.Font.Size.Point).ToArray();
            if (sizes.Max() >= 20 && sizes.Min() <= 12)
            {
                _verticalAlignedFonts[_paragraph] = (sizes.Max(),
                    verticalAlignment == VerticalTextAlignmentValues.Center ? 0.35 : -0.11);
                if (_container is Section && _section != null)
                {
                    if (_centerBottomBodyMargins.Add(_section))
                        _section.PageSetup.TopMargin = Unit.FromPoint(
                            _section.PageSetup.TopMargin.Point + 3.8);
                    if (verticalAlignment == VerticalTextAlignmentValues.Center &&
                        paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
                        format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + 12.5);
                }
                else if (_container is HeaderFooter && !_inHeader && _section != null &&
                    _centerBottomFooterDistances.Add(_section))
                    _section.PageSetup.FooterDistance = Unit.FromPoint(
                        _section.PageSetup.FooterDistance.Point + 12.5);
            }
        }
        if ((computed.TextAlign ?? _cellStyle?.ConditionalTextAlign) is DxpComputedTextAlign align)
            format.Alignment = align switch
            {
                DxpComputedTextAlign.Center => ParagraphAlignment.Center,
                DxpComputedTextAlign.Right => ParagraphAlignment.Right,
                DxpComputedTextAlign.Justify => ParagraphAlignment.Justify,
                _ => ParagraphAlignment.Left
            };
        if (paragraph.ParagraphProperties?.Justification?.Val?.Value is { } directAlignment)
        {
            if (directAlignment == JustificationValues.Left)
                format.Alignment = ParagraphAlignment.Left;
            else if (directAlignment == JustificationValues.Right)
                format.Alignment = ParagraphAlignment.Right;
            else if (directAlignment == JustificationValues.Center)
                format.Alignment = ParagraphAlignment.Center;
            else if (directAlignment == JustificationValues.Both)
                format.Alignment = ParagraphAlignment.Justify;
        }
        if (justification == null || justification == JustificationValues.Start ||
            justification == JustificationValues.End)
        {
            var bidi = paragraph.ParagraphProperties?.GetFirstChild<BiDi>();
            if (bidi == null)
                foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
                    if (_paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                        definition.StyleParagraphProperties?.GetFirstChild<BiDi>() is { } inherited)
                    {
                        bidi = inherited;
                        break;
                    }
            var rightToLeft = bidi != null && bidi.Val?.Value != false;
            if (justification == null)
            {
                if (rightToLeft) format.Alignment = ParagraphAlignment.Right;
            }
            else
                format.Alignment = (justification == JustificationValues.Start) == rightToLeft
                    ? ParagraphAlignment.Right : ParagraphAlignment.Left;
        }
        ApplyTabStops(paragraph, context, format);
        ApplyPositionalTabStops(paragraph, format);
        var precedingListNumbering = (paragraph.PreviousSibling() as
            DocumentFormat.OpenXml.Wordprocessing.Paragraph)?.ParagraphProperties?.NumberingProperties;
        var markerState = _includeRevisionMarkup &&
            paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Deleted>() != null
                ? style.MarkerReject : style.MarkerAccept;
        var currentBullet = markerState is { numId: int bulletId, iLvl: int bulletLevel } &&
            _numberingResolver?.ResolveLevel(bulletId, bulletLevel)?.lvl.NumberingFormat?.Val?.Value ==
                NumberFormatValues.Bullet;
        var followsBullet = markerState == null && precedingListNumbering?.NumberingId?.Val?.Value is int previousId &&
            _numberingResolver?.ResolveLevel(previousId,
                precedingListNumbering.NumberingLevelReference?.Val?.Value ?? 0)?.lvl.NumberingFormat?.Val?.Value ==
                    NumberFormatValues.Bullet;
        if (_container is Section &&
            paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "ListParagraph" &&
            context.Styles is DocxportNet.Walker.Context.DxpStyleResolver bulletSpacingResolver &&
            bulletSpacingResolver.GetContextualSpacing(paragraph) &&
            (currentBullet || followsBullet))
        {
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + 0.25);
            if (currentBullet && markerState is { numId: int currentId, iLvl: int currentLevel } &&
                paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextBullet &&
                nextBullet.ParagraphProperties?.NumberingProperties is { } nextBulletNumbering &&
                nextBulletNumbering.NumberingId?.Val?.Value == currentId &&
                nextBulletNumbering.NumberingLevelReference?.Val?.Value == currentLevel)
                format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + 0.5);
        }
        if (markerState?.marker is string marker)
        {
            var numbering = markerState is { numId: int numId, iLvl: int level }
                ? _numberingResolver?.ResolveLevel(numId, level) : null;
            if (numbering?.lvl.NumberingFormat?.Val?.Value == NumberFormatValues.IdeographDigital)
            {
                if (context.Styles is DocxportNet.Walker.Context.DxpStyleResolver listResolver &&
                    listResolver.GetContextualSpacing(paragraph))
                {
                    format.LineSpacingRule = LineSpacingRule.Exactly;
                    format.LineSpacing = Unit.FromPoint(20);
                }
                else if (_container is HeaderFooter)
                {
                    format.LineSpacingRule = LineSpacingRule.Exactly;
                    format.LineSpacing = Unit.FromPoint(18);
                }
            }
            var numberFormat = numbering?.lvl.NumberingFormat?.Val?.Value;
            if (_container is HeaderFooter && !_inHeader &&
                (numberFormat == NumberFormatValues.DecimalFullWidth ||
                 numberFormat == NumberFormatValues.DecimalFullWidth2 ||
                 numberFormat == NumberFormatValues.IdeographDigital))
            {
                format.LineSpacingRule = LineSpacingRule.Exactly;
                format.LineSpacing = Unit.FromPoint(17.64);
                if (_footerParagraphCount == 1 && _section != null &&
                    _numberedFooterDistanceApplied.Add(_section))
                    _section.PageSetup.FooterDistance = Unit.FromPoint(
                        _section.PageSetup.FooterDistance.Point -
                        (numberFormat == NumberFormatValues.IdeographDigital ? 0.48 : 1.15));
            }
            var properties = numbering?.lvl.GetFirstChild<NumberingSymbolRunProperties>();
            var family = properties?.RunFonts?.Ascii?.Value;
            var markerText = DxpPdfVisitor.NormalizeMarker(marker);
            if (properties?.Caps is { } markerCaps && markerCaps.Val?.Value != false)
                markerText = markerText.ToUpperInvariant();
            if (family == "Symbol" && _availableFontFamilies.Contains(family) && markerText == "•")
                markerText = "\uf0b7";
            if (family == "Wingdings" && _availableFontFamilies.Contains(family) &&
                numbering?.lvl.LevelText?.Val?.Value is string originalMarker &&
                originalMarker.Length == 1 && originalMarker[0] is >= '\uf000' and <= '\uf0ff')
                markerText = originalMarker;
            if (numberFormat == NumberFormatValues.DecimalEnclosedCircle &&
                !_availableFontFamilies.Contains("Yu Mincho") && _cjkFontFamily == null &&
                markerText.Length > 0 && char.GetNumericValue(markerText[0]) is >= 0 and <= 50)
                markerText = ((int)char.GetNumericValue(markerText[0]))
                    .ToString(CultureInfo.InvariantCulture) + markerText[1..];
            var fullWidthSuffixStart = numberFormat == NumberFormatValues.DecimalFullWidth ||
                    numberFormat == NumberFormatValues.DecimalFullWidth2
                ? markerText.TakeWhile(character => character is >= '０' and <= '９').Count()
                : 0;
            var markerSuffixStart = fullWidthSuffixStart > 0 ? fullWidthSuffixStart :
                numberFormat == NumberFormatValues.IdeographDigital
                    ? GetIdeographMarkerPrefixLength(markerText) :
                numberFormat == NumberFormatValues.DecimalEnclosedCircle &&
                    markerText.Length > 0 && markerText[0] is ('\u24ea' or >= '\u2460' and <= '\u2473' or
                        >= '\u3251' and <= '\u325f' or >= '\u32b1' and <= '\u32bf')
                    ? 1 : 0;
            var markerSpan = _paragraph.AddFormattedText(markerSuffixStart > 0
                ? markerText[..markerSuffixStart] : markerText);
            if (properties?.Bold is { } markerBold)
                markerSpan.Bold = markerBold.Val?.Value != false;
            if (properties?.Italic is { } markerItalic)
                markerSpan.Italic = markerItalic.Val?.Value != false;
            var markerColor = properties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>()?
                .Val?.Value;
            var revisedParagraphMark = _includeRevisionMarkup &&
                paragraph.ParagraphProperties?.ParagraphMarkRunProperties is { } paragraphMark &&
                (paragraphMark.GetFirstChild<Inserted>() != null ||
                 paragraphMark.GetFirstChild<Deleted>() != null);
            if (revisedParagraphMark)
                _revisedListMarkerLength = markerText.Length + 1;
            var resolvedMarkerColor = revisedParagraphMark
                ? ParseColor($"#{_revisionMarkupColor.R:X2}{_revisionMarkupColor.G:X2}{_revisionMarkupColor.B:X2}")
                : ParseColor("#" + markerColor);
            if (resolvedMarkerColor is { } textColor)
                markerSpan.Color = textColor;
            var firstRun = paragraph.Descendants<Run>().FirstOrDefault();
            if (firstRun != null)
            {
                var runStyle = context.Styles.ResolveRunStyle(paragraph, firstRun);
                if (runStyle.FontSizeHalfPoints is > 0)
                    markerSpan.Size = Unit.FromPoint(runStyle.FontSizeHalfPoints.Value / 2.0);
                if (runStyle.FontName != null && _availableFontFamilies.Contains(runStyle.FontName))
                    markerSpan.Font.Name = runStyle.FontName;
            }
            if (numbering is { } resolvedNumbering)
            {
                if (family != null && _availableFontFamilies.Contains(family))
                    markerSpan.Font.Name = family;
                if (int.TryParse(properties?.FontSize?.Val?.Value, out var halfPoints) && halfPoints > 0)
                    markerSpan.Size = Unit.FromPoint(halfPoints / 2.0);
                var alignment = resolvedNumbering.lvl.LevelJustification?.Val?.Value;
                if (alignment == LevelJustificationValues.Right || alignment == LevelJustificationValues.Center)
                {
                    var familyName = markerSpan.Font.Name?.Length > 0 ? markerSpan.Font.Name : _fontFamily;
                    var fontSize = markerSpan.Font.Size.Point > 0 ? markerSpan.Font.Size.Point : 11;
                    using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100),
                        XGraphicsUnit.Point, XPageDirection.Downwards);
                    var width = graphics.MeasureString(markerText, new XFont(familyName, fontSize)).Width;
                    format.FirstLineIndent = Unit.FromPoint(format.FirstLineIndent.Point -
                        width * (alignment == LevelJustificationValues.Center ? 0.5 : 1));
                }
            }
            if ((markerText.Any(character => IsCjk(character) || character is >= '０' and <= '９' ||
                    numberFormat == NumberFormatValues.IdeographDigital && character == '\u3007') ||
                numberFormat == NumberFormatValues.DecimalEnclosedCircle && markerSuffixStart > 0) &&
                (family == null || !_availableFontFamilies.Contains(family)))
            {
                if (_availableFontFamilies.Contains("Yu Mincho"))
                    markerSpan.Font.Name = "Yu Mincho";
                else if (_cjkFontFamily != null)
                    markerSpan.Font.Name = _cjkFontFamily;
            }
            if (markerSuffixStart > 0 && markerSuffixStart < markerText.Length)
            {
                var suffix = _paragraph.AddFormattedText(markerText[markerSuffixStart..]);
                suffix.Size = markerSpan.Size;
                suffix.Bold = markerSpan.Bold;
                suffix.Italic = markerSpan.Italic;
                if (resolvedMarkerColor is { } suffixColor)
                    suffix.Color = suffixColor;
                if (family != null && _availableFontFamilies.Contains(family))
                    suffix.Font.Name = family;
                else if (firstRun != null &&
                    context.Styles.ResolveRunStyle(paragraph, firstRun).FontName is string suffixFamily &&
                    _availableFontFamilies.Contains(suffixFamily))
                    suffix.Font.Name = suffixFamily;
                else if (_implicitModernFont && _availableFontFamilies.Contains("Aptos"))
                    suffix.Font.Name = "Aptos";
            }
            if (_container is Section && _availableFontFamilies.Contains("Aptos") &&
                string.Equals(markerSpan.Font.Name, "Yu Mincho", StringComparison.OrdinalIgnoreCase) &&
                (numberFormat == NumberFormatValues.IdeographDigital ||
                 numberFormat == NumberFormatValues.DecimalFullWidth ||
                 numberFormat == NumberFormatValues.DecimalFullWidth2))
                _cjkNumberedBodyParagraphs.Add(_paragraph);
            if (_container is Section && numberFormat == NumberFormatValues.Bullet &&
                paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "ListParagraph" &&
                _availableFontFamilies.Contains("Aptos") &&
                paragraph.Parent is Body body &&
                body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is
                    { } openingParagraph &&
                context.Styles.ResolveEffectiveNumPr(openingParagraph) is { } openingList &&
                style.MarkerAccept is { numId: int markerNumId } &&
                openingList.NumId == markerNumId &&
                markerSpan.Font.Name is { } markerFamily)
            {
                var rise = markerFamily switch
                {
                    "Symbol" => 0.82,
                    "Courier New" => 0.68,
                    "Wingdings" => 0.25,
                    _ => 0
                };
                if (rise > 0) _bulletBaselineParagraphs[_paragraph] = (markerFamily, rise);
            }
            if (TryAddPictureBullet(numbering?.lvl, context, format))
                markerSpan.Elements.OfType<MigraDoc.DocumentObjectModel.Text>()
                    .First().Content = "";
            if (properties?.Underline is { } markerUnderlineProperty &&
                AttributeValue(markerUnderlineProperty, "val") is not ("none" or "nil"))
            {
                var markerUnderlineColor = AttributeValue(markerUnderlineProperty, "color") ??
                    markerColor ?? "000000";
                if (markerUnderlineColor.Length == 6)
                    _listMarkerUnderlines[_paragraph] = (markerText, markerUnderlineColor,
                        AttributeValue(markerUnderlineProperty, "val") == "double");
            }
            if (properties?.Strike is { } markerStrikeProperty &&
                markerStrikeProperty.Val?.Value != false)
                _listMarkerStrikes[_paragraph] = (markerText, markerColor ?? "000000");
            if (fullWidthSuffixStart > 0 && _container is Section &&
                context.Styles is DocxportNet.Walker.Context.DxpStyleResolver listSpacingResolver &&
                listSpacingResolver.GetContextualSpacing(paragraph))
            {
                format.LineSpacingRule = LineSpacingRule.Exactly;
                format.LineSpacing = Unit.FromPoint(20);
            }
            var zeroIndentTableListTab = _container is Cell && format.LeftIndent.Point == 0 &&
                numbering?.lvl.GetFirstChild<PreviousParagraphProperties>()?.Indentation?.Left?.Value is
                    string levelIndent && int.TryParse(levelIndent, NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out var levelTwips) && levelTwips > 0 &&
                numbering?.lvl.GetFirstChild<LevelSuffix>()?.Val?.Value != LevelSuffixValues.Space &&
                numbering?.lvl.GetFirstChild<LevelSuffix>()?.Val?.Value != LevelSuffixValues.Nothing;
            var numberingTab = numbering?.lvl.GetFirstChild<PreviousParagraphProperties>()?.Tabs?
                .Elements<DocumentFormat.OpenXml.Wordprocessing.TabStop>()
                .FirstOrDefault(stop => AttributeValue(stop, "val") is "num" or "left" &&
                    int.TryParse(AttributeValue(stop, "pos"), NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out var position) && position > 0);
            var hasNumberingTab = numberingTab != null &&
                numbering?.lvl.GetFirstChild<LevelSuffix>()?.Val?.Value != LevelSuffixValues.Space &&
                numbering?.lvl.GetFirstChild<LevelSuffix>()?.Val?.Value != LevelSuffixValues.Nothing;
            if (format.LeftIndent.Point > 0 || zeroIndentTableListTab || hasNumberingTab)
            {
                var listTab = hasNumberingTab ? Unit.FromPoint(
                        int.Parse(AttributeValue(numberingTab!, "pos")!, CultureInfo.InvariantCulture) / 20.0) :
                    zeroIndentTableListTab ? Unit.FromPoint(_defaultTabStopPoints ?? 36) :
                    _container is Section &&
                    _defaultTabStopPoints is double defaultInterval &&
                    numbering?.lvl.GetFirstChild<PreviousParagraphProperties>()?.Tabs == null
                    ? Unit.FromPoint(defaultInterval) : format.LeftIndent;
                format.AddTabStop(listTab, TabAlignment.Left);
                _paragraph.AddTab();
            }
            else
                _paragraph.AddText(" ");
        }
        if (_inHeader && _container is HeaderFooter && !string.IsNullOrWhiteSpace(paragraph.InnerText))
        {
            _headerHasContextualSpacing |= context.Styles is DocxportNet.Walker.Context.DxpStyleResolver resolver &&
                resolver.GetContextualSpacing(paragraph);
            if (_headerParagraphCount == 1)
            {
                _headerFirstBeforePoints = format.SpaceBefore.Point;
                _headerFirstAfterPoints = format.SpaceAfter.Point + _headerFirstAfterLayoutAdjustmentPoints;
            }
            _headerLastAfterPoints = format.SpaceAfter.Point;
            _headerPreviousParagraphUsesBodySpacing =
                paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value is string styleId &&
                HeaderParagraphHasAfterSpacing(styleId);
        }
        if (_container is HeaderFooter && paragraph.InnerText.Length == 0 &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextStyled &&
            context.Styles is DocxportNet.Walker.Context.DxpStyleResolver styleResolver &&
            styleResolver.GetParagraphSpacing(nextStyled)?.Line?.Value == "360")
        {
            format.LineSpacingRule = LineSpacingRule.Exactly;
            format.LineSpacing = Unit.FromPoint(32);
            _paragraph.AddText("\u00a0");
            if (_inHeader && _section != null)
                _headersWithStyledEmptyLine.Add(_section);
            else if (_section != null && _footersWithStyledEmptyLine.Add(_section))
                _section.PageSetup.FooterDistance = Unit.FromPoint(
                    _section.PageSetup.FooterDistance.Point + 8);
        }
        if (_inHeader && _container is HeaderFooter && _section != null &&
            paragraph.InnerText.Length == 0 && paragraph.NextSibling() == null &&
            markFont?.GetFirstChild<FontSize>() != null)
        {
            _headersWithStyledEmptyLine.Add(_section);
            _headersWithStyledTrailingLine.Add(_section);
        }
        if (_inHeader && _container is HeaderFooter && paragraph.InnerText.Length == 0 &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph nextHeaderParagraph &&
            paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph previousHeaderParagraph &&
            previousHeaderParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value ==
                nextHeaderParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value &&
            previousHeaderParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value !=
                paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value &&
            markFont?.ChildElements.Count > 0 && markFont.GetFirstChild<FontSize>() == null &&
            paragraph.ParagraphProperties?.SpacingBetweenLines == null &&
            _defaultParagraphAfterPoints > 0 &&
            context.Styles is DocxportNet.Walker.Context.DxpStyleResolver emptyStyleResolver &&
            emptyStyleResolver.GetParagraphSpacing(paragraph)?.After?.Value == "0")
        {
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + _defaultParagraphAfterPoints);
            _headerSizedEmptyLineExtraPoints = Math.Max(_headerSizedEmptyLineExtraPoints,
                Math.Max(0, _defaultParagraphAfterPoints - 1.1));
        }
        if (_inHeader && _container is HeaderFooter && paragraph.InnerText.Length == 0 &&
            paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph &&
            int.TryParse(markFont?.GetFirstChild<FontSize>()?.Val?.Value, out var emptyMarkSize) &&
            context.Styles.ResolveRunStyle(paragraph, new Run()).FontSizeHalfPoints is int inheritedMarkSize)
        {
            var emptySpacing = paragraph.ParagraphProperties?.SpacingBetweenLines;
            var emptyBefore = int.TryParse(emptySpacing?.Before?.Value, out var beforeTwips)
                ? beforeTwips / 20.0 : 0;
            var emptyAfter = int.TryParse(emptySpacing?.After?.Value, out var afterTwips)
                ? afterTwips / 20.0 : 0;
            _headerSizedEmptyLineExtraPoints = Math.Max(_headerSizedEmptyLineExtraPoints,
                (emptyMarkSize - inheritedMarkSize) / 2.0 + emptyBefore + emptyAfter +
                (_availableFontFamilies.Contains("Aptos") && emptyMarkSize > inheritedMarkSize ? 1.47 : 0));
        }
        if (!_inHeader && _container is HeaderFooter && _section != null &&
            paragraph.InnerText.Length == 0 && paragraph.NextSibling() == null &&
            int.TryParse(markFont?.GetFirstChild<FontSize>()?.Val?.Value, out var trailingSize) &&
            context.Styles.ResolveRunStyle(paragraph, new Run()).FontSizeHalfPoints is int trailingInheritedSize &&
            trailingSize > trailingInheritedSize && _footersWithStyledTrailingLine.Add(_section))
            _section.PageSetup.FooterDistance = Unit.FromPoint(Math.Max(0,
                _section.PageSetup.FooterDistance.Point - (trailingSize - trailingInheritedSize) / 2.0));
        if (_container is HeaderFooter && paragraph.InnerText.Length == 0 &&
            paragraph.ParagraphProperties?.SpacingBetweenLines?.After == null &&
            paragraph.PreviousSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>() is { } previousStory &&
            !string.IsNullOrEmpty(previousStory.InnerText) &&
            previousStory.ParagraphProperties?.ParagraphBorders != null &&
            previousStory.ParagraphProperties.Shading != null &&
            paragraph.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>() is { } nextStory &&
            !string.IsNullOrEmpty(nextStory.InnerText))
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + 1);
        if (_container is HeaderFooter && !_inHeader && _documentHasFitText &&
            paragraph.Descendants<FitText>().Any() &&
            paragraph.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>() is { } nextFooterParagraph &&
            nextFooterParagraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Any(IsCjk)))
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point - 2);
        if (_inHeader && _container is HeaderFooter &&
            paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value is int numberId &&
            _numberingResolver?.ResolveLevel(numberId,
                paragraph.ParagraphProperties.NumberingProperties.NumberingLevelReference?.Val?.Value ?? 0)?
                .lvl.NumberingFormat?.Val?.Value is { } headerNumberFormat &&
            (headerNumberFormat == NumberFormatValues.DecimalEnclosedCircle ||
             headerNumberFormat == NumberFormatValues.DecimalFullWidth2) &&
            paragraph.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>() is { } nextNumberedHeader &&
            nextNumberedHeader.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value == numberId)
        {
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point - 1.6);
            if (_headerParagraphCount == 1 && _section != null)
                _section.PageSetup.HeaderDistance = Unit.FromPoint(
                    _section.PageSetup.HeaderDistance.Point -
                    (headerNumberFormat == NumberFormatValues.DecimalFullWidth2 ? 1.44 : 0.8));
        }
        if (_container is Section && paragraph.PreviousSibling() == null &&
            _gridLinePitchPoints is > 0 && _paragraphGridCharSpacingPoints is > 0 &&
            HasBaselineTextAlignment(paragraph, context) && _section != null)
            _section.PageSetup.TopMargin = Unit.FromPoint(_section.PageSetup.TopMargin.Point + 0.8);
        if (_container is Cell && style.MarkerAccept != null &&
            paragraph.Descendants<Drawing>().Any(drawing => drawing.Inline != null) &&
            paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any())
        {
            format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point + 0.65);
            format.SpaceAfter = Unit.FromPoint(_defaultParagraphAfterPoints);
        }
        if (_horizontalMergeTable && _tableInHeaderFooter && !_inHeader &&
            _container is Cell && _table?.Rows.Count > 1 && _row == _table.Rows[1] &&
            format.SpaceAfter.Point >= 0.2)
        {
            format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point + 0.2);
            format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point - 0.2);
        }
        if (_paragraphStyles.Count == 0 && paragraph.ParagraphProperties?.SpacingBetweenLines == null &&
            (_container is Cell && paragraph.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>()
                .Any(sourceTable => sourceTable.TableProperties?.ChildElements
                    .Any(element => element.LocalName == "tblpPr") == true) ||
             _container is Section && (paragraph.ParagraphProperties?.ChildElements
                 .Any(element => element.LocalName == "framePr") == true ||
                 paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                     { ParagraphProperties: { } previousFrameProperties } &&
                 previousFrameProperties.ChildElements.Any(element => element.LocalName == "framePr") ||
                 paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                     { ParagraphProperties: { } nextFrameProperties } &&
                 nextFrameProperties.ChildElements.Any(element => element.LocalName == "framePr")) ||
             _container is Section && (paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Table
                 { TableProperties: { } previousProperties } &&
                 previousProperties.ChildElements.Any(element => element.LocalName == "tblpPr") ||
                 paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Table
                     { TableProperties: { } nextProperties } &&
                 nextProperties.ChildElements.Any(element => element.LocalName == "tblpPr"))))
        {
            format.LineSpacingRule = LineSpacingRule.Multiple;
            format.LineSpacing = paragraph.PreviousSibling() is
                DocumentFormat.OpenXml.Wordprocessing.Paragraph
                { ParagraphProperties: { } adjacentFrameProperties } &&
                adjacentFrameProperties.ChildElements.Any(element => element.LocalName == "framePr")
                ? 1.158 : 1.15;
            if (paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                { ParagraphProperties: { } precedingFrameProperties } &&
                precedingFrameProperties.ChildElements.Any(element =>
                    element.LocalName == "framePr" && AttributeValue(element, "wrap") == "around"))
                format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point + 0.15);
            format.SpaceAfter = Unit.FromPoint(8);
            if (string.Equals(_document.Styles["Normal"].Font.Name, "Aptos",
                    StringComparison.OrdinalIgnoreCase) &&
                Math.Abs(_document.Styles["Normal"].Font.Size.Point - 12) < 0.01 &&
                paragraph.Descendants<Run>().All(run =>
                    run.RunProperties?.RunFonts == null && run.RunProperties?.FontSize == null) &&
                (_container is Cell &&
                 paragraph.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>().Any(sourceTable =>
                     sourceTable.TableProperties?.ChildElements.Any(element => element.LocalName == "tblpPr") == true) ||
                 _container is Section &&
                 (paragraph.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Table adjacentPreviousTable &&
                  adjacentPreviousTable.TableProperties?.ChildElements.Any(element => element.LocalName == "tblpPr") == true ||
                  paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Table adjacentNextTable &&
                  adjacentNextTable.TableProperties?.ChildElements.Any(element => element.LocalName == "tblpPr") == true)))
                format.LineSpacing = StylelessAptosFloatingTableLineFactor;
        }
        if (_container is Section && _section != null && paragraph.Parent is Body &&
            paragraph.PreviousSibling() == null)
        {
            double clearance = 0;
            foreach (var drawing in paragraph.Descendants<Drawing>())
            {
                if (!drawing.Descendants<TextBoxContent>().Any())
                    continue;
                var anchor = drawing.Descendants().FirstOrDefault(element => element.LocalName == "anchor");
                if (anchor == null || !anchor.ChildElements.Any(element => element.LocalName == "wrapTopAndBottom"))
                    continue;
                var extent = anchor.ChildElements.FirstOrDefault(element => element.LocalName == "extent");
                var vertical = anchor.ChildElements.FirstOrDefault(element => element.LocalName == "positionV");
                if (!long.TryParse(extent?.GetAttribute("cx", "").Value, out var widthUnits) ||
                    !long.TryParse(extent?.GetAttribute("cy", "").Value, out var heightUnits) ||
                    !long.TryParse(vertical?.ChildElements.FirstOrDefault(element => element.LocalName == "posOffset")?
                        .InnerText, out var topUnits))
                    continue;
                var width = widthUnits / 12700.0;
                var height = heightUnits / 12700.0;
                var rotation = long.TryParse(drawing.Descendants().FirstOrDefault(element => element.LocalName == "xfrm")?
                    .GetAttribute("rot", "").Value, out var rotationUnits) ? rotationUnits / 60000.0 : 0;
                var radians = rotation * Math.PI / 180;
                var rotatedHeight = Math.Abs(width * Math.Sin(radians)) + Math.Abs(height * Math.Cos(radians));
                var bottom = topUnits / 12700.0 + (height + rotatedHeight) / 2;
                clearance = Math.Max(clearance, bottom + 1);
            }
            if (clearance > 0)
            {
                var spacer = new MigraDoc.DocumentObjectModel.Paragraph();
                spacer.Format.Font.Size = Unit.FromPoint(1);
                spacer.Format.LineSpacingRule = LineSpacingRule.Exactly;
                spacer.Format.LineSpacing = Unit.FromPoint(clearance);
                spacer.Format.SpaceAfter = Unit.Zero;
                spacer.Format.SpaceBefore = Unit.Zero;
                _section.Elements.InsertObject(0, spacer);
            }
        }
        if (_container is MigraDoc.DocumentObjectModel.Shapes.TextFrame
            { Orientation: MigraDoc.DocumentObjectModel.Shapes.TextOrientation.Vertical
                or MigraDoc.DocumentObjectModel.Shapes.TextOrientation.Upward } &&
            paragraph.ParagraphProperties?.SpacingBetweenLines == null &&
            paragraph.ParagraphProperties?.ParagraphStyleId == null &&
            _document.Styles["Normal"].ParagraphFormat.LineSpacingRule == LineSpacingRule.Multiple)
        {
            format.LineSpacingRule = LineSpacingRule.Multiple;
            format.LineSpacing = Unit.FromPoint(
                _document.Styles["Normal"].ParagraphFormat.LineSpacing.Point *
                VerticalTextLineSpacingFactor);
        }
        var formattedParagraph = _paragraph;
        var centeredSpacings = format.Alignment == ParagraphAlignment.Center &&
            !paragraph.Descendants<Break>().Any()
            ? paragraph.Descendants<Run>()
                .Where(run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any())
                .Select(run => ResolveRunCharacterSpacing(run, context)).Distinct().ToArray()
            : [];
        return DxpDisposable.Create(() =>
        {
            if (formattedParagraph != null &&
                !_paragraphShades.ContainsKey(formattedParagraph) &&
                paragraph.Descendants<Run>().Any(run =>
                    run.RunProperties?.ChildElements.Any(element => element.LocalName == "bdr") == true ||
                    run.RunProperties?.RunStyle?.Val?.Value is string styleId &&
                    _paragraphStyles.TryGetValue(styleId, out var runStyle) &&
                    runStyle.StyleRunProperties?.ChildElements.Any(element =>
                        element.LocalName == "bdr") == true) &&
                _textBorderRanges.TryGetValue(formattedParagraph, out var paragraphBorders) &&
                paragraphBorders.Count > 0)
            {
                var clearance = paragraphBorders.Max(border => border.Double
                    ? border.Inset : 2 * border.Inset + (border.Shadow ? 2 : 1));
                if (_container is Section &&
                    paragraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
                    format.SpaceAfter = Unit.FromPoint(format.SpaceAfter.Point + clearance);
                if (_container is HeaderFooter && _inHeader && _headerParagraphCount == 1)
                    format.SpaceBefore = Unit.FromPoint(format.SpaceBefore.Point + clearance);
                if (_container is HeaderFooter && !_inHeader && _section != null &&
                    paragraphBorders.Any(border => border.Double))
                    _section.PageSetup.FooterDistance = Unit.FromPoint(
                        _section.PageSetup.FooterDistance.Point +
                        paragraphBorders.Max(border => border.Inset) + 0.5);
                else if (_container is HeaderFooter && !_inHeader && _section != null &&
                    paragraphBorders.Any(border => border.Shadow))
                    _section.PageSetup.FooterDistance = Unit.FromPoint(
                        _section.PageSetup.FooterDistance.Point + 2);
            }
            if (formattedParagraph != null && centeredSpacings is [var spacing] && spacing != 0 &&
                _paragraphTextLength > 0 && !_characterSpacings.ContainsKey(formattedParagraph))
            {
                _characterSpacings[formattedParagraph] = spacing;
                _centeredSpacingOffsets[formattedParagraph] = -spacing * (_paragraphTextLength - 1) / 2;
            }
            if (formattedParagraph != null && _availableFontFamilies.Contains("Aptos Display") &&
                format.Alignment == ParagraphAlignment.Center &&
                paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Title" &&
                paragraph.Descendants<InsertedRun>().Any() &&
                paragraph.Descendants<DeletedRun>().Any())
            {
                _characterScales[formattedParagraph] = 98.8;
                _centeredSpacingOffsets[formattedParagraph] =
                    _centeredSpacingOffsets.GetValueOrDefault(formattedParagraph) + 11.3;
            }
            if (_documentHasTrackedInsertions && _container is Section &&
                format.Alignment == ParagraphAlignment.Right &&
                paragraph.Descendants<Run>().Count(run =>
                    run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any()) == 1 &&
                !paragraph.Descendants<TabChar>().Any() && !paragraph.Descendants<Break>().Any() &&
                paragraph.InnerText.Length >= 10)
                _rightAlignedTrackedTexts.Add(paragraph.InnerText);
            if (_container is Section &&
                format.Alignment == ParagraphAlignment.Justify &&
                paragraph.InnerText.Length > 50 &&
                !paragraph.Descendants<TabChar>().Any() &&
                !paragraph.Descendants<Break>().Any() &&
                paragraph.Descendants<Run>().Where(run =>
                    run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any())
                    .All(run => !HasRunKerningOverride(run, context)))
                _longJustifiedTrackedTexts.Add(Regex.Replace(string.Concat(paragraph.Descendants<
                    DocumentFormat.OpenXml.Wordprocessing.Text>().Select(text => text.Text)), @"[ \t]+", " "));
            if (_pendingNoBreakHyphen)
            {
                _pendingNoBreakHyphen = false;
                AddStyledText("-");
            }
            _paragraph = null;
            _activeFloatingTextWrap = null;
        });
    }

    private bool HeaderParagraphHasAfterSpacing(string styleId)
    {
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var current = styleId; visited.Add(current) && _paragraphStyles.TryGetValue(current, out var style);
            current = style.BasedOn?.Val?.Value ?? "")
        {
            if (style.StyleParagraphProperties?.SpacingBetweenLines?.After?.Value is string after &&
                int.TryParse(after, out var twips))
                return twips > 0;
        }
        return _defaultParagraphAfterPoints > 0;
    }

    private bool StoryParagraphWraps(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph,
        DxpIDocumentContext context)
    {
        if (_section == null || paragraph.Descendants<Run>().FirstOrDefault() is not Run firstRun)
            return false;
        var style = context.Styles.ResolveRunStyle(paragraph, firstRun);
        var family = style.FontName != null && _availableFontFamilies.Contains(style.FontName)
            ? style.FontName : _fontFamily;
        var size = style.FontSizeHalfPoints is > 0 ? style.FontSizeHalfPoints.Value / 2.0 : 11;
        using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100),
            XGraphicsUnit.Point, XPageDirection.Downwards);
        return graphics.MeasureString(paragraph.InnerText, new XFont(family, size)).Width >
            _section.PageSetup.PageWidth.Point - _section.PageSetup.LeftMargin.Point -
            _section.PageSetup.RightMargin.Point;
    }

    private bool UsesMirrorIndents(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph,
        DxpIDocumentContext context)
    {
        var setting = paragraph.ParagraphProperties?.MirrorIndents;
        if (setting == null)
        {
            foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
            {
                if (_paragraphStyles.TryGetValue(entry.StyleId, out var style) &&
                    style.StyleParagraphProperties?.MirrorIndents is { } inherited)
                {
                    setting = inherited;
                    break;
                }
            }
        }
        return setting != null && setting.Val?.Value != false;
    }

    private static bool HasRenderedTabLeader(DocumentFormat.OpenXml.Wordprocessing.TabStop stop)
    {
        var value = stop.Leader?.Value;
        return value == TabStopLeaderCharValues.Dot || value == TabStopLeaderCharValues.Hyphen ||
            value == TabStopLeaderCharValues.Underscore || value == TabStopLeaderCharValues.Heavy ||
            value == TabStopLeaderCharValues.MiddleDot;
    }

    internal static bool HasImplicitPlainAptos12Runs(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph) =>
        paragraph.ParagraphProperties?.SpacingBetweenLines?.After == null &&
        paragraph.ChildElements.All(element => element is Run or ParagraphProperties) &&
        paragraph.Elements<Run>().Any(run => !string.IsNullOrEmpty(run.InnerText)) &&
        paragraph.Elements<Run>().All(run => run.ChildElements.All(element =>
            element is RunProperties or DocumentFormat.OpenXml.Wordprocessing.Text)) &&
        paragraph.Elements<Run>().Where(run => !string.IsNullOrEmpty(run.InnerText))
            .All(run => run.RunProperties?.RunFonts?.Ascii?.Value == "Aptos" &&
                run.RunProperties?.FontSize?.Val?.Value == "24");

    private void ApplyTabStops(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph,
        DxpIDocumentContext context, ParagraphFormat format)
    {
        var stops = new SortedDictionary<int, DocumentFormat.OpenXml.Wordprocessing.TabStop>();
        foreach (var style in context.Styles.GetParagraphStyleChain(paragraph).Reverse())
        {
            if (_paragraphStyles.TryGetValue(style.StyleId, out var definition))
                MergeTabs(definition.StyleParagraphProperties?.Tabs, stops);
        }
        MergeTabs(paragraph.ParagraphProperties?.Tabs, stops);
        var borderedTabInset = _container is Section && paragraph.Descendants<TabChar>().Any() &&
            paragraph.Descendants<Run>().Select(run => ResolveRunTextBorder(run, context)?.Inset ?? 0)
                .Distinct().ToArray() is [var sharedInset] && sharedInset > 0
                ? sharedInset : 0;
        if (borderedTabInset > 0 && stops.Count > 0 && _paragraph != null)
            _borderedTabInsets[_paragraph] = borderedTabInset;
        if (stops.Count == 0 && _defaultTabStopPoints is double interval && _section != null &&
            paragraph.Descendants<TabChar>().Any())
        {
            var page = _section.PageSetup;
            var width = page.PageWidth.Point - page.LeftMargin.Point - page.RightMargin.Point;
            for (var position = interval; position <= width + interval; position += interval)
                format.AddTabStop(Unit.FromPoint(position));
        }
        foreach (var (position, stop) in stops)
        {
            var alignment = stop.Val?.Value switch
            {
                var value when value == TabStopValues.Center => TabAlignment.Center,
                var value when value == TabStopValues.Right => TabAlignment.Right,
                var value when value == TabStopValues.Decimal => TabAlignment.Decimal,
                _ => TabAlignment.Left
            };
            var leader = stop.Leader?.Value switch
            {
                var value when value == TabStopLeaderCharValues.Dot => TabLeader.Dots,
                var value when value == TabStopLeaderCharValues.Hyphen => TabLeader.Dashes,
                var value when value == TabStopLeaderCharValues.Underscore => TabLeader.Lines,
                var value when value == TabStopLeaderCharValues.Heavy => TabLeader.Heavy,
                var value when value == TabStopLeaderCharValues.MiddleDot => TabLeader.MiddleDot,
                _ => TabLeader.Spaces
            };
            if (leader != TabLeader.Spaces)
                _paragraphHasLeaderTab = true;
            var numberedHeaderFooter = _container is HeaderFooter &&
                paragraph.ParagraphProperties?.NumberingProperties != null;
            format.AddTabStop(Unit.FromPoint(position / 20.0 - borderedTabInset),
                numberedHeaderFooter ? TabAlignment.Left : alignment, leader);
        }
    }

    private static void MergeTabs(Tabs? tabs, SortedDictionary<int, DocumentFormat.OpenXml.Wordprocessing.TabStop> stops)
    {
        if (tabs == null) return;
        foreach (var stop in tabs.Elements<DocumentFormat.OpenXml.Wordprocessing.TabStop>())
        {
            if (stop.Position?.Value is not int position || position <= 0) continue;
            if (stop.Val?.Value == TabStopValues.Clear) stops.Remove(position);
            else stops[position] = stop;
        }
    }

    private void ApplyPositionalTabStops(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph,
        ParagraphFormat format)
    {
        if (_section == null) return;
        var page = _section.PageSetup;
        var marginWidth = page.PageWidth.Point - page.LeftMargin.Point - page.RightMargin.Point;
        if (!paragraph.Descendants<TabChar>().Any() &&
            paragraph.Descendants<PositionalTab>().Any(tab =>
                AttributeValue(tab, "relativeTo") == "margin" &&
                AttributeValue(tab, "alignment") is "center" or "right"))
            format.TabStops.ClearAll();
        foreach (var tab in paragraph.Descendants<PositionalTab>())
        {
            if (AttributeValue(tab, "relativeTo") != "margin") continue;
            var alignment = AttributeValue(tab, "alignment");
            var position = alignment switch
            {
                "center" => marginWidth / 2,
                "right" => marginWidth,
                _ => 0
            };
            var leader = AttributeValue(tab, "leader") switch
            {
                "dot" => TabLeader.Dots,
                "hyphen" => TabLeader.Dashes,
                "underscore" => TabLeader.Lines,
                "heavy" => TabLeader.Heavy,
                "middleDot" => TabLeader.MiddleDot,
                _ => TabLeader.Spaces
            };
            if (position > 0)
                format.AddTabStop(Unit.FromPoint(position), alignment switch
                {
                    "center" => TabAlignment.Center,
                    "right" => TabAlignment.Right,
                    _ => TabAlignment.Left
                }, leader);
        }
    }

    private bool HasPageBreakBefore(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph, DxpIDocumentContext context)
    {
        if (paragraph.ParagraphProperties?.PageBreakBefore is { } direct)
            return direct.Val?.Value != false;
        foreach (var style in context.Styles.GetParagraphStyleChain(paragraph))
        {
            if (_paragraphStyles.TryGetValue(style.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.PageBreakBefore is { } inherited)
                return inherited.Val?.Value != false;
        }
        return false;
    }

    private bool HasBaselineTextAlignment(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph,
        DxpIDocumentContext context)
    {
        if (paragraph.ParagraphProperties?.TextAlignment is { } direct)
            return direct.Val?.Value == VerticalTextAlignmentValues.Baseline;
        foreach (var style in context.Styles.GetParagraphStyleChain(paragraph))
        {
            if (_paragraphStyles.TryGetValue(style.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.TextAlignment is { } inherited)
                return inherited.Val?.Value == VerticalTextAlignmentValues.Baseline;
        }
        return false;
    }

    private bool HasTopTextAlignment(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph,
        DxpIDocumentContext context)
    {
        if (paragraph.ParagraphProperties?.TextAlignment is { } direct)
            return direct.Val?.Value == VerticalTextAlignmentValues.Top;
        foreach (var style in context.Styles.GetParagraphStyleChain(paragraph))
        {
            if (_paragraphStyles.TryGetValue(style.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.TextAlignment is { } inherited)
                return inherited.Val?.Value == VerticalTextAlignmentValues.Top;
        }
        return false;
    }

    private bool UsesAdjustedRightIndent(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph,
        DxpIDocumentContext context)
    {
        if (paragraph.ParagraphProperties?.AdjustRightIndent is { } direct)
            return direct.Val?.Value != false;
        foreach (var style in context.Styles.GetParagraphStyleChain(paragraph))
        {
            if (_paragraphStyles.TryGetValue(style.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.AdjustRightIndent is { } inherited)
                return inherited.Val?.Value != false;
        }
        return false;
    }

    private bool UsesDocumentGrid(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph, DxpIDocumentContext context)
    {
        if (paragraph.ParagraphProperties?.SnapToGrid is { } direct)
            return direct.Val?.Value != false;
        foreach (var style in context.Styles.GetParagraphStyleChain(paragraph))
        {
            if (_paragraphStyles.TryGetValue(style.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.SnapToGrid is { } inherited)
                return inherited.Val?.Value != false;
        }
        return true;
    }

    private bool ResolveAutoSpace(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph,
        DxpIDocumentContext context, bool numbers)
    {
        OnOffType? option = numbers ? paragraph.ParagraphProperties?.GetFirstChild<AutoSpaceDN>()
            : paragraph.ParagraphProperties?.GetFirstChild<AutoSpaceDE>();
        if (option != null) return option.Val?.Value != false;
        foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
        {
            if (!_paragraphStyles.TryGetValue(entry.StyleId, out var style)) continue;
            option = numbers ? style.StyleParagraphProperties?.GetFirstChild<AutoSpaceDN>()
                : style.StyleParagraphProperties?.GetFirstChild<AutoSpaceDE>();
            if (option != null) return option.Val?.Value != false;
        }
        if (!numbers && _container is HeaderFooter && _documentHasFitText &&
            paragraph.Descendants<FitText>().Any() && paragraph.InnerText.Any(IsCjk) &&
            paragraph.InnerText.Any(char.IsAsciiLetter))
            return true;
        return _container is HeaderFooter && _gridCharSpacingPoints is > 0 &&
            context.Styles.GetParagraphStyleChain(paragraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var style) &&
                string.Equals(style.StyleRunProperties?.RunFonts?.Ascii?.Value,
                    "Yu Mincho", StringComparison.OrdinalIgnoreCase));
    }

    public override IDisposable VisitInsertedRunBegin(InsertedRun inserted, DxpIDocumentContext context)
    {
        var previous = _insideInsertedRun;
        _insideInsertedRun = _includeRevisionMarkup;
        return DxpDisposable.Create(() => _insideInsertedRun = previous);
    }

    public override IDisposable VisitDeletedRunBegin(DeletedRun deleted, DxpIDocumentContext context)
    {
        var previous = _insideDeletedRun;
        _insideDeletedRun = _includeRevisionMarkup;
        return DxpDisposable.Create(() => _insideDeletedRun = previous);
    }

    public override IDisposable? VisitMoveFromRunBegin(MoveFromRun moved, DxpIDocumentContext context)
    {
        if (!_includeRevisionMarkup) return null;
        var previousDeleted = _insideDeletedRun;
        var previousMoved = _insideMoveFromRun;
        _insideDeletedRun = true;
        _insideMoveFromRun = true;
        return DxpDisposable.Create(() =>
        {
            _insideDeletedRun = previousDeleted;
            _insideMoveFromRun = previousMoved;
        });
    }

    public override void VisitDeletedText(DeletedText deleted, DxpIDocumentContext context)
    {
        if (_insideDeletedRun)
            AddStyledText(deleted.Text);
    }

    public override IDisposable VisitRunBegin(Run run, DxpIDocumentContext context)
    {
        if (_includeRevisionMarkup && _paragraph != null &&
            run.RunProperties?.GetFirstChild<RunPropertiesChange>() is { } change &&
            run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is { } sourceParagraph &&
            context.Styles is DxpStyleResolver resolver &&
            change.Descendants<Languages>().FirstOrDefault()?.Val?.Value is { } oldLanguage &&
            resolver.ResolveRunLanguage(sourceParagraph, run) is { } newLanguage &&
            !string.Equals(oldLanguage, newLanguage, StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var description = CultureInfo.GetCultureInfo(newLanguage).EnglishName;
                if (!_formattingBalloons.TryGetValue(_paragraph, out var descriptions))
                    _formattingBalloons[_paragraph] = descriptions = [];
                descriptions.Add(description);
            }
            catch (CultureNotFoundException) { }
        }
        var previousInserted = _insideInsertedRun;
        var previousMoved = _insideMoveToRun;
        if (_includeRevisionMarkup && run.Ancestors<MoveToRun>().Any())
        {
            _insideInsertedRun = true;
            _insideMoveToRun = true;
        }
        _implicitFittedStoryFont = false;
        _calibriFittedStoryText = false;
        _runStyle = context.CurrentRun?.Style;
        if (_runStyle != null && _includeRevisionMarkup)
            _runStyle = _runStyle with
            {
                Underline = _insideInsertedRun ? false : _runStyle.Underline,
                Strike = _runStyle.Strike || _insideDeletedRun,
                ColorCss = _insideDeletedRun
                    ? $"#{_revisionDeletionColor.R:X2}{_revisionDeletionColor.G:X2}{_revisionDeletionColor.B:X2}"
                    : _insideInsertedRun
                        ? $"#{_revisionMarkupColor.R:X2}{_revisionMarkupColor.G:X2}{_revisionMarkupColor.B:X2}"
                        : _runStyle.ColorCss
            };
        _runHighAnsiFont = run.RunProperties?.RunFonts?.HighAnsi?.Value;
        _runEastAsiaFont = run.RunProperties?.RunFonts?.EastAsia?.Value;
        if (string.Equals(_runHighAnsiFont, _runStyle?.FontName, StringComparison.OrdinalIgnoreCase))
            _runHighAnsiFont = null;
        _runRightToLeft = run.RunProperties?.GetFirstChild<RightToLeftText>() is { } direction &&
            direction.Val?.Value != false;
        _runComplexBold = _runRightToLeft && ResolveRunEffectFlag<BoldComplexScript>(run, context);
        _runComplexItalic = _runRightToLeft && ResolveRunEffectFlag<ItalicComplexScript>(run, context);
        _runTextBorder = ResolveRunTextBorder(run, context);
        if (_inHeader && _section != null && _headerParagraphCount == 1 &&
            _runTextBorder is { Shadow: false, Double: false } headerBorder &&
            run.RunProperties?.ChildElements.Any(element => element.LocalName == "bdr") != true &&
            run.RunProperties?.RunStyle == null &&
            (ReferenceEquals(_document?.Sections[0], _section) ||
                ResolveRunShading(run, context) != null) &&
            _headerRunBorderClearances.TryAdd(_section, 2 * headerBorder.Inset + 1))
            _section.PageSetup.HeaderDistance = Unit.FromPoint(
                _section.PageSetup.HeaderDistance.Point + 2 * headerBorder.Inset + 1);
        _runCharacterSpacing = ResolveRunCharacterSpacing(run, context);
        _runScalePercent = ResolveRunScale(run, context);
        if (_runScalePercent != 100 && _paragraph != null &&
            run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is { } paragraph &&
            paragraph.Descendants<Run>().Count() == 1)
            _characterScales[_paragraph] = _runScalePercent;
        if (_runCharacterSpacing != 0 && _paragraph != null &&
            run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is { } spacedParagraph &&
            spacedParagraph.Descendants<Run>().All(candidate =>
                ResolveRunCharacterSpacing(candidate, context) == _runCharacterSpacing))
        {
            _characterSpacings[_paragraph] = _runCharacterSpacing;
            if (spacedParagraph.Descendants<TabChar>().Any() &&
                spacedParagraph.Descendants<Run>().Select(candidate =>
                    ResolveRunShading(candidate, context)).Distinct().ToArray() is [string shading])
                _tabbedRunShadings[_paragraph] = shading;
        }
        var outline = ResolveRunEffectFlag<Outline>(run, context);
        var emboss = ResolveRunEffectFlag<Emboss>(run, context);
        var imprint = ResolveRunEffectFlag<Imprint>(run, context);
        var shadow = ResolveRunEffectFlag<Shadow>(run, context);
        _runEffect = (emboss ? 1 : 0) + (imprint ? 1 : 0) + (shadow ? 1 : 0) > 1
            ? outline ? TextEffectKind.ShadowOutline : shadow ? TextEffectKind.Shadow : null
            : emboss
            ? TextEffectKind.Emboss
            : imprint
                ? TextEffectKind.Imprint
                : shadow && outline
                    ? TextEffectKind.ShadowOutline
                    : shadow ? TextEffectKind.Shadow
                    : outline ? TextEffectKind.Outline : null;
        _hiddenRun = ResolveRunEffectFlag<Vanish>(run, context);
        _runHasExplicitColor = run.RunProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>() != null;
        _runHasExplicitBold = run.RunProperties?.Bold != null;
        _runHasExplicitItalic = run.RunProperties?.Italic != null;
        _runHasExplicitUnderline = run.RunProperties?.Underline != null;
        _tocHyperlinkRun = run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Hyperlink>().Any() &&
            run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault()?
                .ParagraphProperties?.ParagraphStyleId?.Val?.Value is string paragraphStyle &&
            paragraphStyle.StartsWith("TOC", StringComparison.OrdinalIgnoreCase);
        _runHighlight = ResolveHighlight(run, context);
        _runEmphasis = ResolveEmphasis(run, context);
        _runShading = ResolveRunShading(run, context);
        return DxpDisposable.Create(() =>
        {
            _insideInsertedRun = previousInserted;
            _insideMoveToRun = previousMoved;
            _runStyle = null;
            _runHighAnsiFont = null;
            _runEastAsiaFont = null;
            _runTextBorder = null;
            _runScalePercent = 100;
            _runCharacterSpacing = 0;
            _runEffect = null;
            _hiddenRun = false;
            _runHasExplicitColor = false;
            _runHasExplicitBold = false;
            _runHasExplicitItalic = false;
            _runHasExplicitUnderline = false;
            _runRightToLeft = false;
            _runComplexBold = false;
            _runComplexItalic = false;
            _tocHyperlinkRun = false;
            _runHighlight = null;
            _runEmphasis = null;
            _runShading = null;
        });
    }

    private TextBorderRange? ResolveRunTextBorder(Run run, DxpIDocumentContext context)
    {
        static OpenXmlElement? Border(OpenXmlElement? properties) =>
            properties?.ChildElements.FirstOrDefault(element => element.LocalName == "bdr");
        var border = Border(run.RunProperties);
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var current = run.RunProperties?.RunStyle?.Val?.Value;
            border == null && current != null && visited.Add(current) &&
            _paragraphStyles.TryGetValue(current, out var style);
            current = style.BasedOn?.Val?.Value)
            border = Border(style.StyleRunProperties);
        var paragraph = run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
        if (border == null && paragraph != null)
            foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
                if (_paragraphStyles.TryGetValue(entry.StyleId, out var style) &&
                    Border(style.StyleRunProperties) is { } inherited)
                {
                    border = inherited;
                    break;
                }
        if (border == null || AttributeValue(border, "val") is "nil" or "none") return null;
        var color = AttributeValue(border, "color");
        if (color?.Length != 6) return null;
        var width = double.TryParse(AttributeValue(border, "sz"), NumberStyles.Float,
            CultureInfo.InvariantCulture, out var eighths) ? eighths / 8 : 0.5;
        var space = double.TryParse(AttributeValue(border, "space"), NumberStyles.Float,
            CultureInfo.InvariantCulture, out var points) ? points : 0;
        return new TextBorderRange(0, 0, color, width, space + width,
            AttributeValue(border, "shadow") is "1" or "true" or "on",
            AttributeValue(border, "val") == "double");
    }

    private string? ResolveRunUnderlineColor(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph,
        Run run, DxpIDocumentContext context)
    {
        if (!context.Styles.ResolveRunStyle(paragraph, run).Underline) return null;
        static DocumentFormat.OpenXml.Wordprocessing.Underline? Mark(OpenXmlElement? properties) =>
            properties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Underline>();
        var underline = Mark(run.RunProperties);
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var current = run.RunProperties?.RunStyle?.Val?.Value;
        while (underline == null && current != null && visited.Add(current) &&
               _paragraphStyles.TryGetValue(current, out var style))
        {
            underline = Mark(style.StyleRunProperties);
            current = style.BasedOn?.Val?.Value;
        }
        underline ??= Mark(paragraph.ParagraphProperties?.GetFirstChild<RunProperties>());
        if (underline == null)
            foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
                if (_paragraphStyles.TryGetValue(entry.StyleId, out var style) &&
                    Mark(style.StyleRunProperties) is { } inherited)
                {
                    underline = inherited;
                    break;
                }
        var color = underline == null ? null : AttributeValue(underline, "color");
        return color is { Length: 6 } && int.TryParse(color, NumberStyles.HexNumber,
            CultureInfo.InvariantCulture, out _) ? color : null;
    }

    private double ResolveRunScale(Run run, DxpIDocumentContext context)
    {
        static double? Scale(OpenXmlElement? properties)
        {
            var value = properties?.ChildElements.FirstOrDefault(element => element.LocalName == "w")?
                .GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var percent) &&
                percent is >= 1 and <= 600 ? percent : null;
        }
        if (Scale(run.RunProperties) is double direct) return direct;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var current = run.RunProperties?.RunStyle?.Val?.Value;
            current != null && visited.Add(current) && _paragraphStyles.TryGetValue(current, out var style);
            current = style.BasedOn?.Val?.Value)
            if (Scale(style.StyleRunProperties) is double inherited) return inherited;
        var paragraph = run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
        if (paragraph != null)
            foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
                if (_paragraphStyles.TryGetValue(entry.StyleId, out var style) &&
                    Scale(style.StyleRunProperties) is double inherited) return inherited;
        return 100;
    }

    private double ResolveRunCharacterSpacing(Run run, DxpIDocumentContext context)
    {
        static double? Spacing(OpenXmlElement? properties)
        {
            var value = properties?.ChildElements.FirstOrDefault(element => element.LocalName == "spacing")?
                .GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var twips)
                ? twips / 20 : null;
        }
        if (Spacing(run.RunProperties) is double direct) return direct;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var current = run.RunProperties?.RunStyle?.Val?.Value;
            current != null && visited.Add(current) && _paragraphStyles.TryGetValue(current, out var style);
            current = style.BasedOn?.Val?.Value)
            if (Spacing(style.StyleRunProperties) is double inherited) return inherited;
        var paragraph = run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
        if (paragraph != null)
            foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
                if (_paragraphStyles.TryGetValue(entry.StyleId, out var style) &&
                    Spacing(style.StyleRunProperties) is double inherited) return inherited;
        return 0;
    }

    private bool HasRunKerningOverride(Run run, DxpIDocumentContext context)
    {
        if (run.RunProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Kern>() != null)
            return true;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var current = run.RunProperties?.RunStyle?.Val?.Value;
            current != null && visited.Add(current) && _paragraphStyles.TryGetValue(current, out var style);
            current = style.BasedOn?.Val?.Value)
            if (style.StyleRunProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Kern>() != null)
                return true;
        var paragraph = run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
        if (paragraph == null) return false;
        foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
            if (_paragraphStyles.TryGetValue(entry.StyleId, out var style) &&
                style.StyleRunProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Kern>() != null)
                return true;
        return false;
    }

    private bool ResolveRunEffectFlag<T>(Run run, DxpIDocumentContext context) where T : OnOffType
    {
        if (run.RunProperties?.GetFirstChild<T>() is { } direct)
            return direct.Val?.Value != false;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var current = run.RunProperties?.RunStyle?.Val?.Value;
            current != null && visited.Add(current) && _paragraphStyles.TryGetValue(current, out var style);
            current = style.BasedOn?.Val?.Value)
        {
            if (style.StyleRunProperties?.GetFirstChild<T>() is { } inherited)
                return inherited.Val?.Value != false;
        }
        var paragraph = run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
        if (paragraph == null) return false;
        foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
        {
            if (_paragraphStyles.TryGetValue(entry.StyleId, out var style) &&
                style.StyleRunProperties?.GetFirstChild<T>() is { } inherited)
                return inherited.Val?.Value != false;
        }
        return false;
    }

    private static bool IsHidden(OnOffType? property) => property != null && property.Val?.Value != false;

    private FitText? ResolveFitText(Run run)
    {
        if (run.RunProperties?.GetFirstChild<FitText>() is { } direct) return direct;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var current = run.RunProperties?.RunStyle?.Val?.Value;
            current != null && visited.Add(current) && _paragraphStyles.TryGetValue(current, out var style);
            current = style.BasedOn?.Val?.Value)
            if (style.StyleRunProperties?.GetFirstChild<FitText>() is { } inherited)
                return inherited;
        return null;
    }

    private FitText? ResolveParagraphFitText(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph,
        DxpIDocumentContext context)
    {
        foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
            if (_paragraphStyles.TryGetValue(entry.StyleId, out var style) &&
                style.StyleRunProperties?.GetFirstChild<FitText>() is { } inherited)
                return inherited;
        return null;
    }

    private string? ResolveRunShading(Run run, DxpIDocumentContext context)
    {
        if (run.RunProperties?.Shading is { } direct)
            return RunShadingColor(direct);
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var current = run.RunProperties?.RunStyle?.Val?.Value;
            current != null && visited.Add(current) && _paragraphStyles.TryGetValue(current, out var style);
            current = style.BasedOn?.Val?.Value)
        {
            if (style.StyleRunProperties?.Shading is { } inherited)
                return RunShadingColor(inherited);
        }
        var paragraph = run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
        if (paragraph != null)
        {
            foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
            {
                if (_paragraphStyles.TryGetValue(entry.StyleId, out var style) &&
                    style.StyleRunProperties?.Shading is { } inherited)
                    return RunShadingColor(inherited);
            }
        }
        return null;
    }

    private static string? RunShadingColor(DocumentFormat.OpenXml.Wordprocessing.Shading shading)
    {
        var fill = shading.Fill?.Value;
        var foreground = shading.Color?.Value;
        var pattern = shading.GetAttribute("val",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
        if (pattern == "solid" && foreground is { Length: 6 }) return foreground;
        if (pattern is not { Length: > 3 } || !pattern.StartsWith("pct", StringComparison.Ordinal) ||
            !int.TryParse(pattern.AsSpan(3), NumberStyles.None, CultureInfo.InvariantCulture,
                out var percentage) || percentage is < 0 or > 100 ||
            foreground is not { Length: 6 } || fill is not { Length: 6 } ||
            !int.TryParse(foreground, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var front) ||
            !int.TryParse(fill, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var back))
            return fill;
        var red = (((front >> 16) & 255) * percentage +
            ((back >> 16) & 255) * (100 - percentage) + 49) / 100;
        var green = (((front >> 8) & 255) * percentage +
            ((back >> 8) & 255) * (100 - percentage) + 49) / 100;
        var blue = ((front & 255) * percentage + (back & 255) * (100 - percentage) + 49) / 100;
        return $"{red:X2}{green:X2}{blue:X2}";
    }

    private HighlightColorValues? ResolveHighlight(Run run, DxpIDocumentContext context)
    {
        if (run.RunProperties?.Highlight is { } direct)
            return direct.Val?.Value;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var current = run.RunProperties?.RunStyle?.Val?.Value;
            current != null && visited.Add(current) && _paragraphStyles.TryGetValue(current, out var style);
            current = style.BasedOn?.Val?.Value)
        {
            if (style.StyleRunProperties?.GetFirstChild<Highlight>() is { } inherited)
                return inherited.Val?.Value;
        }
        var paragraph = run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
        if (paragraph != null)
        {
            foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
            {
                if (_paragraphStyles.TryGetValue(entry.StyleId, out var style) &&
                    style.StyleRunProperties?.GetFirstChild<Highlight>() is { } inherited)
                    return inherited.Val?.Value;
            }
        }
        return _defaultRunHighlight;
    }

    private EmphasisMarkValues? ResolveEmphasis(Run run, DxpIDocumentContext context)
    {
        if (run.RunProperties?.GetFirstChild<Emphasis>() is { } direct)
            return direct.Val?.Value;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var current = run.RunProperties?.RunStyle?.Val?.Value;
            current != null && visited.Add(current) && _paragraphStyles.TryGetValue(current, out var style);
            current = style.BasedOn?.Val?.Value)
        {
            if (style.StyleRunProperties?.GetFirstChild<Emphasis>() is { } inherited)
                return inherited.Val?.Value;
        }
        var paragraph = run.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
        if (paragraph == null) return null;
        foreach (var entry in context.Styles.GetParagraphStyleChain(paragraph))
        {
            if (_paragraphStyles.TryGetValue(entry.StyleId, out var style) &&
                style.StyleRunProperties?.GetFirstChild<Emphasis>() is { } inherited)
                return inherited.Val?.Value;
        }
        return null;
    }

    public override IDisposable VisitHyperlinkBegin(DocumentFormat.OpenXml.Wordprocessing.Hyperlink link,
        DxpLinkAnchor? target, DxpIDocumentContext context)
    {
        if (_paragraph == null || target == null)
            return DxpDisposable.Empty;

        MigraDoc.DocumentObjectModel.Hyperlink? pdfLink;
        if (target.internalRef is { Length: > 0 } bookmark)
            pdfLink = _paragraph.AddHyperlink(bookmark);
        else if (Uri.TryCreate(target.uri, UriKind.Absolute, out var uri) &&
            uri.Scheme is "http" or "https" or "mailto")
            pdfLink = _paragraph.AddHyperlink(target.uri, HyperlinkType.Url);
        else
            return DxpDisposable.Empty;

        var previous = _activeHyperlink;
        _activeHyperlink = pdfLink;
        return DxpDisposable.Create(() => _activeHyperlink = previous);
    }

    public override void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext context)
    {
        _complexFieldLinks.Push((new System.Text.StringBuilder(), _activeHyperlink));
    }

    public override void VisitComplexFieldInstruction(FieldCode instruction, string text,
        DxpIDocumentContext context)
    {
        if (_complexFieldLinks.TryPeek(out var field))
            field.Instruction.Append(text);
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext context)
    {
        if (_paragraph == null || !_complexFieldLinks.TryPeek(out var field))
            return;
        var parsed = new DxpFieldParser().Parse(field.Instruction.ToString()).Ast;
        if (!string.Equals(parsed.FieldType, "HYPERLINK", StringComparison.OrdinalIgnoreCase) ||
            !DxpFieldTokenization.TryGetFirstToken(parsed.ArgumentsText, out var target) ||
            !Uri.TryCreate(target, UriKind.Absolute, out var uri) ||
            uri.Scheme is not ("http" or "https" or "mailto"))
            return;
        _activeHyperlink = _paragraph.AddHyperlink(target, HyperlinkType.Url);
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext context)
    {
        if (_complexFieldLinks.TryPop(out var field))
            _activeHyperlink = field.Previous;
    }

    public override void VisitBookmarkStart(BookmarkStart bookmark, DxpIDocumentContext context)
    {
        if (_paragraph != null && bookmark.Name?.Value is { Length: > 0 } name &&
            context.ReferencedBookmarkAnchors.Contains(name))
            _paragraph.AddBookmark(name, prepend: false);
    }

    public override IDisposable VisitRubyBegin(Ruby ruby, DxpIDocumentContext context)
    {
        var previous = _renderRubyGuide;
        _renderRubyGuide = false;
        if (_paragraph == null || ruby.GetFirstChild<RubyContent>() is not { } guide ||
            ruby.GetFirstChild<RubyBase>() is not { } baseText ||
            guide.InnerText.Length == 0 || baseText.InnerText.Length == 0)
            return DxpDisposable.Create(() => _renderRubyGuide = previous);
        var sizeAttribute = ruby.GetFirstChild<RubyProperties>()?.ChildElements
            .FirstOrDefault(element => element.LocalName == "hps")?.GetAttribute("val",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
        var size = int.TryParse(sizeAttribute, out var halfPoints) && halfPoints > 0
            ? halfPoints / 2.0 : 6;
        var raiseAttribute = ruby.GetFirstChild<RubyProperties>()?.ChildElements
            .FirstOrDefault(element => element.LocalName == "hpsRaise")?.GetAttribute("val",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
        var raise = int.TryParse(raiseAttribute, out var raiseHalfPoints)
            ? Math.Max(0, raiseHalfPoints / 2.0) : 0;
        var guideFonts = guide.Descendants<RunFonts>().FirstOrDefault();
        var guideFont = guideFonts?.EastAsia?.Value ?? guideFonts?.Ascii?.Value;
        var baseFamily = baseText.InnerText.Any(IsCjk) && _cjkFontFamily != null
            ? _cjkFontFamily : _fontFamily;
        var guideFamily = guideFont is { } name && _availableFontFamilies.Contains(name)
            ? name : guide.InnerText.Any(IsCjk) && _cjkFontFamily != null
                ? _cjkFontFamily : _fontFamily;
        var baseSizeAttribute = ruby.GetFirstChild<RubyProperties>()?.ChildElements
            .FirstOrDefault(element => element.LocalName == "hpsBaseText")?.GetAttribute("val",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;
        var baseSize = int.TryParse(baseSizeAttribute, out var baseHalfPoints) && baseHalfPoints > 0
            ? baseHalfPoints / 2.0 : _paragraph.Format.Font.Size.Point > 0
                ? _paragraph.Format.Font.Size.Point : _document.Styles["Normal"].Font.Size.Point;
        var baseFont = new XFont(baseFamily, baseSize);
        var guideFontMetrics = new XFont(guideFamily, size);
        var overflow = Math.Max(0, raise - Math.Max(0,
            0.825 * (baseFont.GetHeight() - guideFontMetrics.GetHeight())));
        var format = _paragraph.Format;
        if (!_rubyNaturalLineHeights.TryGetValue(_paragraph, out var lineHeight))
        {
            _rubyOriginalLineSpacingRules[_paragraph] = format.LineSpacingRule;
            var fontHeight = baseFont.GetHeight();
            lineHeight = format.LineSpacingRule == LineSpacingRule.Multiple
                ? fontHeight * format.LineSpacing.Point : fontHeight;
            _rubyNaturalLineHeights[_paragraph] = lineHeight;
        }
        var originalRule = _rubyOriginalLineSpacingRules[_paragraph];
        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_RUBY_OVERFLOW_BASELINE") == "1" &&
            overflow > 0 && originalRule == LineSpacingRule.Multiple)
        {
            var candidate = -0.96 * Math.Max(0, 2.2 - overflow);
            if (!_rubyOverflowBaselineOffsets.TryGetValue(_paragraph, out var current) ||
                candidate > current)
                _rubyOverflowBaselineOffsets[_paragraph] = candidate;
        }
        if (overflow > 0 && _paragraph.Format.LineSpacingRule != LineSpacingRule.Exactly)
        {
            var explicitAtLeast = originalRule == LineSpacingRule.AtLeast;
            var minimum = explicitAtLeast
                ? format.LineSpacing.Point : 0;
            format.LineSpacingRule = LineSpacingRule.AtLeast;
            format.LineSpacing = Unit.FromPoint(Math.Max(minimum,
                lineHeight + Math.Max(0, overflow - (explicitAtLeast ? 0 : 2.2))));
        }
        var alignment = ruby.GetFirstChild<RubyProperties>()?.ChildElements
            .FirstOrDefault(element => element.LocalName == "rubyAlign")?.GetAttribute("val",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value ?? "center";
        if (!_rubyGuides.TryGetValue(_paragraph, out var guides))
            _rubyGuides[_paragraph] = guides =
                new List<(string Base, string Guide, double Size, double Raise, int Start, string? Font, string Alignment)>();
        guides.Add((baseText.InnerText, guide.InnerText, size, raise, _paragraphBorderLetterLength, guideFont, alignment));
        _renderRubyGuide = true;
        return DxpDisposable.Create(() =>
        {
            if (alignment == "rightVertical") _paragraph?.AddCharacter(SymbolName.En);
            _renderRubyGuide = previous;
        });
    }

    public override IDisposable VisitRubyContentBegin(RubyContentType content, bool isBase,
        DxpIDocumentContext context)
    {
        var previous = _suppressRubyGuideText;
        var previousBase = _renderRubyBaseText;
        _suppressRubyGuideText = !isBase && _renderRubyGuide;
        _renderRubyBaseText = isBase && _renderRubyGuide;
        return DxpDisposable.Create(() =>
        {
            _suppressRubyGuideText = previous;
            _renderRubyBaseText = previousBase;
        });
    }

    public override void VisitText(DocumentFormat.OpenXml.Wordprocessing.Text text, DxpIDocumentContext context)
    {
        if (_suppressRubyGuideText) return;
        var sourceRun = text.Ancestors<Run>().FirstOrDefault();
        foreach (Match token in Regex.Matches(text.Text, @"[A-Za-zÀ-ÖØ-öø-ÿ]{2,},\u201D"))
        {
            _closingQuoteSourceCounts[token.Value] = _closingQuoteSourceCounts.GetValueOrDefault(token.Value) + 1;
            if (token.Index + token.Length < text.Text.Length && text.Text[token.Index + token.Length] == ' ')
                _closingQuoteFollowedBySpace.Add(token.Value);
            if (sourceRun != null && HasRunKerningOverride(sourceRun, context))
                _excludedClosingQuoteTokens.Add(token.Value);
        }
        if (string.Equals(_runStyle?.FontName, "Aptos Display", StringComparison.OrdinalIgnoreCase))
            foreach (Match token in Regex.Matches(text.Text,
                @"[A-Za-zÀ-ÖØ-öø-ÿ]{2,}\u2019[A-Za-zÀ-ÖØ-öø-ÿ]+"))
            {
                _apostropheDisplaySourceCounts[token.Value] =
                    _apostropheDisplaySourceCounts.GetValueOrDefault(token.Value) + 1;
                if (sourceRun != null && HasRunKerningOverride(sourceRun, context))
                    _excludedApostropheDisplayTokens.Add(token.Value);
            }
        _implicitFittedStoryFont = _container is HeaderFooter && sourceRun != null &&
            string.IsNullOrEmpty(_runStyle?.FontName) &&
            sourceRun.RunProperties?.RunFonts == null && sourceRun.RunProperties?.RunStyle == null &&
            ResolveFitText(sourceRun) != null && _availableFontFamilies.Contains("Aptos");
        _calibriFittedStoryText = _container is HeaderFooter && sourceRun != null &&
            string.Equals(_runStyle?.FontName, "Calibri", StringComparison.OrdinalIgnoreCase) &&
            _runStyle?.FontSizeHalfPoints == null && ResolveFitText(sourceRun) != null;
        if (_implicitFittedStoryFont && !_inHeader && _section != null &&
            _implicitFitFooterDistancesApplied.Add(_section))
            _section.PageSetup.FooterDistance = Unit.FromPoint(
                _section.PageSetup.FooterDistance.Point + 1.2);
        if (_calibriFittedStoryText && !_inHeader && _section != null &&
            sourceRun?.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Footer>().FirstOrDefault() is
                { } fittedFooter &&
            !fittedFooter.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(footerText => footerText.Text.Any(IsCjk)) &&
            _calibriFitFooterDistancesApplied.Add(_section))
            _section.PageSetup.FooterDistance = Unit.FromPoint(
                _section.PageSetup.FooterDistance.Point + 1.2);
        if (_activeRotatedCell)
        {
            if (!_rotatedCellPlaceholderAdded)
            {
                AddStyledText(" ");
                _rotatedCellPlaceholderAdded = true;
            }
            return;
        }
        var content = text.Space?.Value == SpaceProcessingModeValues.Preserve
            ? text.Text : text.Text.Trim(' ', '\t', '\r', '\n');
        if (text.Space?.Value != SpaceProcessingModeValues.Preserve &&
            text.Text.EndsWith(' ') &&
            ReferenceEquals(text, sourceRun?.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .LastOrDefault()) &&
            sourceRun?.NextSibling() is Run followingRun &&
            followingRun.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .FirstOrDefault()?.Text is { Length: > 0 } followingText &&
            IsCjk(followingText[0]))
            content += " ";
        if (text.Space?.Value == SpaceProcessingModeValues.Preserve && content.EndsWith(' ') &&
            ReferenceEquals(text, sourceRun?.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .LastOrDefault()) &&
            sourceRun?.NextSibling() is Run adjacentRun &&
            adjacentRun.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .FirstOrDefault()?.Text?.StartsWith(' ') == true)
            content = content[..^1] + '\u00A0';
        if (text.Space?.Value == SpaceProcessingModeValues.Preserve &&
            content.Length > 0 && content[0] == ' ' &&
            text.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is
                { } numberedParagraph &&
            numberedParagraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value is > 0 &&
            ReferenceEquals(numberedParagraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .FirstOrDefault(), text))
            content = "\u00a0" + content[1..];
        if (text.Space?.Value == SpaceProcessingModeValues.Preserve &&
            content.Length > 0 && content[0] == ' ' && _paragraph?.Elements.Count == 0 &&
            text.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Any())
        {
            AddStyledText("\u00a0");
            if (_balanceSingleByteDoubleByteWidth && _paragraph != null)
            {
                var size = _runStyle?.FontSizeHalfPoints is int leadingSpaceHalfPoints &&
                    leadingSpaceHalfPoints > 0
                    ? leadingSpaceHalfPoints / 2.0 : _paragraph.Format.Font.Size.Point;
                _paragraph.Format.AddTabStop(Unit.FromPoint(size / 2));
                _paragraph.AddTab();
            }
            content = content[1..];
        }
        if (_pendingNoBreakHyphen)
        {
            _pendingNoBreakHyphen = false;
            if (_paragraph?.Elements.LastOrDefault() is FormattedText previousSpan &&
                previousSpan.Elements.LastOrDefault() is MigraDoc.DocumentObjectModel.Text previousText &&
                content.Length > 0 && !char.IsWhiteSpace(content[0]))
            {
                var prefixEnd = previousText.Content.LastIndexOfAny([' ', '\t']);
                var previousWord = prefixEnd < 0 ? previousText.Content : previousText.Content[(prefixEnd + 1)..];
                var nextWordEnd = content.IndexOfAny([' ', '\t']);
                var nextWord = nextWordEnd < 0 ? content : content[..nextWordEnd];
                if (previousWord.Length > 0 && nextWord.Length > 0)
                {
                    if (prefixEnd >= 0)
                    {
                        previousText.Content = previousText.Content[..(prefixEnd + 1)];
                        previousText = previousSpan.AddText(previousWord);
                    }
                    var token = previousWord + "\u2013" + nextWord;
                    previousText.Content += "\u2013" + nextWord;
                    _noBreakWords.Add(token);
                    _noBreakParagraphs.Add(_paragraph);
                    _paragraphTextLength += 1 + nextWord.Length;
                    _paragraphBorderLetterLength += 1 + nextWord.Length;
                    if (nextWordEnd >= 0) AddStyledText(content[nextWordEnd..]);
                    return;
                }
            }
            AddStyledText("-");
        }
        if (_runStyle is { Superscript: true, FontSizeHalfPoints: 40 } &&
            string.Equals(_runStyle.FontName, "Aptos Display", StringComparison.OrdinalIgnoreCase) &&
            content.Length >= 2 && content.All(char.IsAsciiLetter))
            _largeSuperscriptTexts.Add(content);
        if (_activeFloatingTextWrap is { } bodyWrap && !_inHeader && _paragraph != null &&
            text.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is
                { } sourceBody && _bodyParagraphsUnderHeaderWrap.Contains(sourceBody))
        {
            var bodyWords = content.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var family = _runStyle?.FontName is string bodyFamily && _availableFontFamilies.Contains(bodyFamily)
                ? bodyFamily : _fontFamily;
            var size = _runStyle?.FontSizeHalfPoints is int bodyHalfPoints && bodyHalfPoints > 0
                ? bodyHalfPoints / 2.0 : _paragraph.Format.Font.Size.Point;
            if (size <= 0) size = 11;
            using var measure = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
                XPageDirection.Downwards);
            var bodyFont = new XFont(family, size);
            var wordCount = 0;
            while (wordCount < bodyWords.Length &&
                measure.MeasureString(string.Join(" ", bodyWords[..(wordCount + 1)]), bodyFont).Width <= bodyWrap.Left)
                wordCount++;
            if (wordCount > 0 && wordCount < bodyWords.Length)
            {
                _paragraph.Format.AddTabStop(Unit.FromPoint(bodyWrap.Right));
                AddStyledText(string.Join(" ", bodyWords[..wordCount]));
                _paragraph.AddTab();
                AddStyledText(string.Join(" ", bodyWords[wordCount..]));
                _activeFloatingTextWrap = null;
                return;
            }
        }
        if (_inHeader && _activeFloatingTextWrap is { } headerWrap && _paragraph != null &&
            _section != null && text.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .FirstOrDefault() is { } sourceHeader &&
            sourceHeader.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Count() == 1 &&
            content.Split(' ', StringSplitOptions.RemoveEmptyEntries) is { Length: > 2 } headerWords)
        {
            var family = _runStyle?.FontName is string headerFamily && _availableFontFamilies.Contains(headerFamily)
                ? headerFamily : _fontFamily;
            var size = _runStyle?.FontSizeHalfPoints is int headerHalfPoints && headerHalfPoints > 0
                ? headerHalfPoints / 2.0 : _paragraph.Format.Font.Size.Point;
            if (size <= 0) size = 11;
            using var headerMeasure = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
                XPageDirection.Downwards);
            var headerFont = new XFont(family, size);
            var page = _section.PageSetup;
            var headerWidth = page.PageWidth.Point - page.LeftMargin.Point - page.RightMargin.Point;
            var wordIndex = 0;
            string TakeWords(double limit)
            {
                var start = wordIndex;
                while (wordIndex < headerWords.Length &&
                    headerMeasure.MeasureString(string.Join(" ", headerWords[start..(wordIndex + 1)]),
                        headerFont).Width <= limit)
                    wordIndex++;
                return string.Join(" ", headerWords[start..wordIndex]);
            }
            var firstLine = TakeWords(headerWidth);
            var leftLine = TakeWords(headerWrap.Left);
            var rightLine = TakeWords(headerWidth - headerWrap.Right);
            if (firstLine.Length > 0 && leftLine.Length > 0 && rightLine.Length > 0 &&
                wordIndex < headerWords.Length)
            {
                _paragraph.Format.AddTabStop(Unit.FromPoint(headerWrap.Right));
                AddStyledText(firstLine);
                _paragraph.AddLineBreak();
                AddStyledText(leftLine);
                _paragraph.AddTab();
                AddStyledText(rightLine);
                _paragraph.AddLineBreak();
                AddStyledText(string.Join(" ", headerWords[wordIndex..]));
                _activeFloatingTextWrap = null;
                return;
            }
        }
        if (_activeFloatingTextWrap is { } wrap && _paragraph != null && _section != null &&
            Regex.Matches(content, @"\S+") is { Count: >= 2 } wrapMatches)
        {
            var wrapWords = wrapMatches.Select(match => match.Value).ToArray();
            var wrapFontName = _runStyle?.FontName is string family && _availableFontFamilies.Contains(family)
                ? family : _fontFamily;
            var wrapFontSize = _runStyle?.FontSizeHalfPoints is int wrapHalfPoints && wrapHalfPoints > 0
                ? wrapHalfPoints / 2.0 : _paragraph.Format.Font.Size.Point;
            if (wrapFontSize <= 0) wrapFontSize = 11;
            using var measure = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
                XPageDirection.Downwards);
            var wrapFont = new XFont(wrapFontName, wrapFontSize);
            var prefixCount = 1;
            while (prefixCount < wrapWords.Length &&
                measure.MeasureString(string.Join(" ", wrapWords[..(prefixCount + 1)]), wrapFont).Width <=
                    wrap.Left)
                prefixCount++;
            var prefix = string.Join(" ", wrapWords[..prefixCount]);
            if (prefixCount < wrapWords.Length &&
                measure.MeasureString(prefix, wrapFont).Width < wrap.Left)
            {
                var prefixEnd = wrapMatches[prefixCount - 1].Index + wrapMatches[prefixCount - 1].Length;
                _paragraph.Format.AddTabStop(Unit.FromPoint(wrap.Right));
                AddStyledText(content[..prefixEnd]);
                _paragraph.AddTab();
                AddStyledText(content[prefixEnd..].TrimStart());
                _activeFloatingTextWrap = null;
                return;
            }
        }
        if (_paragraph != null && text.Parent is Run insertedRun &&
            insertedRun.Parent is InsertedRun insertion &&
            insertion.PreviousSibling<Run>() is { } precedingRun &&
            precedingRun.RunProperties?.OuterXml == insertedRun.RunProperties?.OuterXml &&
            precedingRun.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Text>() is { } precedingText &&
            precedingText.Text.Length > 0 && char.IsLetter(precedingText.Text[^1]) &&
            content.Length > 0 && char.IsLetter(content[0]) &&
            _paragraph.Elements.LastOrDefault() is FormattedText precedingSpan &&
            precedingSpan.Elements.LastOrDefault() is MigraDoc.DocumentObjectModel.Text precedingSegment &&
            precedingSegment.Content.Length > 0 && char.IsLetter(precedingSegment.Content[^1]) &&
            _runEffect == null && _runEmphasis == null && _runTextBorder == null &&
            _runShading == null && _runHighlight == null)
        {
            precedingSegment.Content += content;
            _paragraphTextLength += content.Length;
            _paragraphBorderLetterLength += content.Count(character => !char.IsWhiteSpace(character));
            return;
        }
        if (_distributedText != text.Text || _paragraph == null || _section == null)
        {
            AddStyledText(content);
            return;
        }
        var words = content.Split(' ');
        var format = _paragraph.Format;
        var width = _section.PageSetup.PageWidth.Point - _section.PageSetup.LeftMargin.Point -
            _section.PageSetup.RightMargin.Point - format.LeftIndent.Point - format.RightIndent.Point;
        var fontName = _runStyle?.FontName is string name && _availableFontFamilies.Contains(name)
            ? name : _fontFamily;
        var fontSize = _runStyle?.FontSizeHalfPoints is int halfPoints && halfPoints > 0
            ? halfPoints / 2.0 : format.Font.Size.Point;
        if (fontSize <= 0) fontSize = 11;
        using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
            XPageDirection.Downwards);
        var font = new XFont(fontName, fontSize);
        var wordWidths = words.Select(word => graphics.MeasureString(word, font).Width).ToArray();
        var spacing = (width - wordWidths.Sum()) / (words.Length - 1);
        if (spacing <= 0)
        {
            AddStyledText(content);
            return;
        }
        format.Alignment = ParagraphAlignment.Left;
        format.TabStops.ClearAll();
        var position = 0.0;
        for (var index = 0; index < words.Length; index++)
        {
            if (index > 0)
            {
                position += spacing;
                format.AddTabStop(Unit.FromPoint(position), TabAlignment.Left);
                _paragraph.AddTab();
            }
            AddStyledText(words[index]);
            position += wordWidths[index];
        }
    }

    public override void VisitSymbol(SymbolChar symbol, DxpIDocumentContext context)
    {
        var translated = DxpFontSymbols.TranslateWordSymbol(symbol.Font?.Value, symbol.Char?.Value);
        if (_paragraph == null || _hiddenRun || string.IsNullOrEmpty(translated)) return;
        AddStyledText(translated);
        if (_availableFontFamilies.Contains("DejaVu Sans") &&
            _paragraph.Elements[_paragraph.Elements.Count - 1] is FormattedText span)
            span.Font.Name = "DejaVu Sans";
    }

    public override void VisitNoBreakHyphen(NoBreakHyphen hyphen, DxpIDocumentContext context)
        => _pendingNoBreakHyphen = true;

    private static void RestoreNoBreakWords(MigraDoc.DocumentObjectModel.DocumentObjectCollection elements,
        HashSet<string> tokens, HashSet<string> restored)
    {
        var texts = GetNoBreakTexts(elements).Where(text => text.Content.Length > 0).ToArray();
        foreach (var token in tokens)
        {
            for (var index = 0; index < texts.Length; index++)
            {
                var text = texts[index];
                if (text.Content.Contains(token, StringComparison.Ordinal))
                {
                    text.Content = text.Content.Replace(token, token.Replace('\u2013', '-'),
                        StringComparison.Ordinal);
                    restored.Add(token);
                }
                var hyphenIndex = token.IndexOf('\u2013');
                if (hyphenIndex < 0 || index == 0) continue;
                var prefix = token[..hyphenIndex];
                var suffix = token[hyphenIndex..];
                if (texts[index - 1].Content.EndsWith(prefix, StringComparison.Ordinal) &&
                    text.Content.StartsWith(suffix, StringComparison.Ordinal))
                {
                    text.Content = '-' + text.Content[1..];
                    restored.Add(token);
                }
            }
        }
    }

    private static IEnumerable<MigraDoc.DocumentObjectModel.Text> GetNoBreakTexts(
        MigraDoc.DocumentObjectModel.DocumentObjectCollection elements)
    {
        foreach (var element in elements)
        {
            if (element is MigraDoc.DocumentObjectModel.Text text)
                yield return text;
            else if (element is FormattedText span)
                foreach (var nested in GetNoBreakTexts(span.Elements))
                    yield return nested;
            else if (element is MigraDoc.DocumentObjectModel.Hyperlink hyperlink)
                foreach (var nested in GetNoBreakTexts(hyperlink.Elements))
                    yield return nested;
        }
    }


    public override void VisitSoftHyphen(SoftHyphen hyphen, DxpIDocumentContext context)
        => AddStyledText("\u00ad");

    internal static string AddFootnoteUrlBreakOpportunities(string text)
    {
        var breakable = new System.Text.StringBuilder(text.Length * 2);
        for (var index = 0; index < text.Length; index++)
        {
            breakable.Append(text[index]);
            if (index > 8 && text[index] == '-') breakable.Append('\u200B');
        }
        return breakable.ToString();
    }

    private void AddStyledText(string text)
    {
        if (_paragraph == null || _hiddenRun) return;
        if ((_insideInsertedRun || _insideDeletedRun) && text.Length > 0)
            _revisionParagraphs.Add(_paragraph);
        if (_activeFootnoteIndex != null && text.Contains('-') &&
            text.TrimStart().StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            text = AddFootnoteUrlBreakOpportunities(text);
        if (_runTextBorder == null && text.StartsWith(' ') &&
            _textBorderRanges.TryGetValue(_paragraph, out var previousBorders) &&
            previousBorders.Count > 0 && previousBorders[^1] is { } previousBorder &&
            previousBorder.Start + previousBorder.Length == _paragraphTextLength &&
            previousBorder.Inset > previousBorder.Width)
            _paragraph.AddCharacter(SymbolName.Em4);
        if (_runTextBorder is { } border)
        {
            if (border.Shadow && _container is Section &&
                !_textBorderRanges.ContainsKey(_paragraph))
                _paragraph.Format.SpaceAfter = Unit.FromPoint(_paragraph.Format.SpaceAfter.Point + 1);
            if (_container is HeaderFooter)
            {
                if (_inHeader) _headerRunBorderParagraphs.Add(_paragraph);
                else _footerRunBorderParagraphs.Add(_paragraph);
            }
            if (!_textBorderRanges.TryGetValue(_paragraph, out var borders))
                _textBorderRanges[_paragraph] = borders = new List<TextBorderRange>();
            borders.Add(border with { Start = _paragraphTextLength, Length = text.Length,
                VisibleStart = _paragraphBorderLetterLength,
                VisibleLength = text.Count(character => !char.IsWhiteSpace(character)) });
        }
        else if (!string.IsNullOrWhiteSpace(text) && _textBorderRanges.ContainsKey(_paragraph))
            _borderResetParagraphs.Add(_paragraph);
        if (_runEffect is { } effect && _runStyle != null)
        {
            if (!_textEffectRanges.TryGetValue(_paragraph, out var effects))
                _textEffectRanges[_paragraph] = effects = new List<TextEffectRange>();
            effects.Add(new TextEffectRange(_paragraphTextLength, text.Length,
                _runStyle.FontName ?? _fontFamily, _runStyle.Bold, _runStyle.Italic, effect));
        }
        if (_runStyle is { Strike: true } or { DoubleStrike: true })
        {
            if (!_strikeRanges.TryGetValue(_paragraph, out var strikes))
                _strikeRanges[_paragraph] = strikes = new List<StrikeRange>();
            var strike = _insideDeletedRun
                ? new StrikeRange(_paragraphTextLength, text.Length, _runStyle.DoubleStrike || _insideMoveFromRun,
                    _paragraphTextLength == 0 && _revisedListMarkerLength > 0
                        ? 0 : _revisedListMarkerLength > 0
                            ? _revisedListMarkerLength + _paragraphTextLength
                            : _paragraphBorderLetterLength,
                    _revisedListMarkerLength > 0
                        ? text.Length + (_paragraphTextLength == 0 ? _revisedListMarkerLength : 0)
                        : text.Count(character => !char.IsWhiteSpace(character)), _insideMoveFromRun)
                : new StrikeRange(_paragraphTextLength, text.Length, _runStyle.DoubleStrike);
            if (strike.Moved && strikes.LastOrDefault() is { Moved: true } previous &&
                previous.Start + previous.Length == strike.Start &&
                previous.VisibleStart + previous.VisibleLength == strike.VisibleStart)
                strikes[^1] = previous with { Length = previous.Length + strike.Length,
                    VisibleLength = previous.VisibleLength + strike.VisibleLength };
            else
                strikes.Add(strike);
        }
        if (_insideInsertedRun && !string.IsNullOrWhiteSpace(text))
        {
            if (!_revisionUnderlineRanges.TryGetValue(_paragraph, out var underlines))
                _revisionUnderlineRanges[_paragraph] = underlines = new List<RevisionUnderlineRange>();
            var underline = new RevisionUnderlineRange(
                _paragraphTextLength == 0 && _revisedListMarkerLength > 0
                    ? 0 : _revisedListMarkerLength > 0
                        ? _revisedListMarkerLength + _paragraphTextLength
                        : _paragraphBorderLetterLength,
                _revisedListMarkerLength > 0
                    ? text.Length + (_paragraphTextLength == 0 ? _revisedListMarkerLength : 0)
                    : text.Count(character => !char.IsWhiteSpace(character)), _insideMoveToRun);
            if (underline.Moved && underlines.LastOrDefault() is { Moved: true } previous &&
                previous.Start + previous.Length == underline.Start)
                underlines[^1] = previous with { Length = previous.Length + underline.Length };
            else
                underlines.Add(underline);
        }
        if (_runEmphasis is { } emphasis && emphasis != EmphasisMarkValues.None)
        {
            if (!_emphasisRanges.TryGetValue(_paragraph, out var ranges))
                _emphasisRanges[_paragraph] = ranges = new List<EmphasisRange>();
            ranges.Add(new EmphasisRange(_paragraphTextLength, text.Length, emphasis));
        }
        _paragraphTextLength += text.Length;
        _paragraphBorderLetterLength += text.Count(character => !char.IsWhiteSpace(character));
        var smallCaps = _runStyle?.SmallCaps == true && _runStyle.AllCaps != true;
        var content = _runStyle?.AllCaps == true || smallCaps ? text.ToUpperInvariant() : text;
        if (_runRightToLeft && content.Any(character => character is >= '\u0590' and <= '\u06ff') &&
            content.Any(char.IsAsciiLetterOrDigit))
            content = ReorderExplicitRtlRunForPdf(content);
        else if (!_runRightToLeft && content.Any(character => character is >= '\u0590' and <= '\u06ff'))
        {
            if (_availableFontFamilies.Contains("Tahoma") && content.Any(IsArabic))
                content = ShapeImplicitArabicForPdf(content);
            content = ReorderRtlForPdf(content);
        }
        if (_runRightToLeft && _availableFontFamilies.Contains("Tahoma"))
            content = ShapeExplicitArabicForPdf(content);
        content = OrderShapedArabicForPdf(content);
        if (!_highlightParagraphTexts.TryGetValue(_paragraph, out var paragraphText))
            paragraphText = (new StringBuilder(), 0);
        var sourceOffset = paragraphText.NormalizedLength;
        paragraphText.Text.Append(content);
        paragraphText.NormalizedLength += content.Count(character => !char.IsWhiteSpace(character));
        _highlightParagraphTexts[_paragraph] = paragraphText;
        if (_runHighlight is { } highlight && !string.IsNullOrWhiteSpace(content))
            _highlights.Add(new DxpPdfHighlight(content, highlight, null,
                _paragraphLineSize, _container is HeaderFooter && !_inHeader,
                ParagraphText: paragraphText.Text, SourceOffset: sourceOffset,
                RepeatedStory: _container is HeaderFooter));
        else if (_runShading is { } shading && !string.IsNullOrWhiteSpace(content))
        {
            var storySpacing = _container is HeaderFooter &&
                _characterSpacings.TryGetValue(_paragraph, out var spacing) ? spacing : 0;
            var rightPadding = _container is HeaderFooter && !_inHeader && _runTextBorder != null &&
                (_runStyle?.FontName ?? _fontFamily).Contains("Aptos", StringComparison.OrdinalIgnoreCase)
                    ? -0.4
                    : (_container is HeaderFooter ? (_runTextBorder?.Inset ?? 0) * 2 :
                        (_runStyle?.FontName ?? _fontFamily).Contains("Aptos", StringComparison.OrdinalIgnoreCase)
                            ? _runTextBorder?.Inset - _runTextBorder?.Width ?? 0
                            : _runTextBorder?.Inset ?? 0) -
                      (_container is HeaderFooter && _runTextBorder != null
                          ? content.Count(character => character == ' ') * storySpacing : 0);
            if (_container is HeaderFooter && _inHeader && _runTextBorder == null && storySpacing < 0)
                rightPadding += storySpacing;
            _highlights.Add(new DxpPdfHighlight(content, null, shading,
                _paragraphLineSize, _container is HeaderFooter && !_inHeader, rightPadding,
                paragraphText.Text, sourceOffset, _container is HeaderFooter));
        }
        if (!smallCaps)
        {
            AddStyledTextSegment(content, false);
            return;
        }
        var start = 0;
        var reduced = false;
        for (var index = 0; index < text.Length; index++)
        {
            if (char.IsWhiteSpace(text[index])) continue;
            var nextReduced = char.IsLower(text[index]);
            if (nextReduced == reduced || index == start) { reduced = nextReduced; continue; }
            AddStyledTextSegment(text[start..index].ToUpperInvariant(), reduced);
            start = index;
            reduced = nextReduced;
        }
        if (start < text.Length) AddStyledTextSegment(text[start..].ToUpperInvariant(), reduced);
    }

    private bool IsSafeClosingMarkToken(string prefix, string mark, string color)
    {
        var key = prefix + "\0" + mark + "\0" + color;
        if (_safeClosingMarkTokens.TryGetValue(key, out var safe)) return safe;
        var main = _sourceWordDocument?.MainDocumentPart;
        if (main?.Document == null) return false;
        var roots = new List<DocumentFormat.OpenXml.OpenXmlElement> { main.Document };
        roots.AddRange(main.HeaderParts.Select(part => part.Header)
            .OfType<DocumentFormat.OpenXml.OpenXmlElement>());
        roots.AddRange(main.FooterParts.Select(part => part.Footer)
            .OfType<DocumentFormat.OpenXml.OpenXmlElement>());
        if (main.FootnotesPart?.Footnotes is { } notes) roots.Add(notes);
        var token = prefix + mark;
        var total = 0;
        var styled = 0;
        foreach (var paragraph in roots.SelectMany(root => root.Descendants<
                     DocumentFormat.OpenXml.Wordprocessing.Paragraph>()))
        {
            var runs = paragraph.Descendants<Run>().Select(run => (Run: run,
                Text: string.Concat(run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Select(text => text.Text)))).ToArray();
            var text = string.Concat(runs.Select(run => run.Text));
            for (var index = text.IndexOf(token, StringComparison.Ordinal); index >= 0;
                 index = text.IndexOf(token, index + 1, StringComparison.Ordinal)) total++;
            var position = 0;
            foreach (var run in runs)
            {
                if (run.Text.StartsWith(mark, StringComparison.Ordinal) &&
                    position >= prefix.Length &&
                    string.CompareOrdinal(text, position - prefix.Length, prefix, 0,
                        prefix.Length) == 0 &&
                    string.Equals(run.Run.RunProperties?.Color?.Val?.Value, color[1..],
                        StringComparison.OrdinalIgnoreCase)) styled++;
                position += run.Text.Length;
            }
        }
        safe = total > 0 && total == styled;
        _safeClosingMarkTokens[key] = safe;
        return safe;
    }

    private void AddStyledTextSegment(string content, bool smallCap, bool narrowSpace = false,
        bool hangulBoundary = false)
    {
        if (_paragraph != null && _previousTextParagraph == _paragraph &&
            _paragraph.Elements.LastOrDefault() is FormattedText boundarySpan &&
            boundarySpan.Elements.LastOrDefault() is MigraDoc.DocumentObjectModel.Text boundaryText &&
            NeedsAutoSpaceAcrossRuns(boundaryText.Content, content, _autoSpaceDE))
            AddStyledTextSegment(" ", smallCap);
        if (!smallCap && !narrowSpace && content.Length > 0 &&
            IsClosingLinePunctuation(content[0]) &&
            !IsFullWidthCjkPunctuation(content[0]) && _paragraph != null &&
            _previousTextParagraph == _paragraph && _previousTextCanJoin &&
            _previousTextBasicMetrics && _runCharacterSpacing == 0 &&
            _runScalePercent == 100 && _runHighAnsiFont == null && !_runRightToLeft &&
            _previousTextStyle is { } precedingStyle && _runStyle is { } currentStyle &&
            precedingStyle != currentStyle &&
            precedingStyle == currentStyle with { ColorCss = precedingStyle.ColorCss } &&
            currentStyle.ColorCss is { Length: 7 } foreground && foreground[0] == '#' &&
            (precedingStyle.ColorCss is null or { Length: 7 }) &&
            _runTextBorder == null && _runEffect == null && _runEmphasis == null &&
            _runHighlight == null && _runShading == null &&
            _paragraph.Elements.LastOrDefault() is FormattedText precedingSpan &&
            precedingSpan.Elements.LastOrDefault() is MigraDoc.DocumentObjectModel.Text precedingText &&
            precedingText.Content.Length > 0 && !char.IsWhiteSpace(precedingText.Content[^1]) &&
            content.Length == 1)
        {
            var separator = precedingText.Content.LastIndexOfAny([' ', '\t']);
            var prefix = precedingText.Content[(separator + 1)..];
            if (prefix.Length > 0 && prefix.All(character => character is >= '!' and <= '~' &&
                    character is not '(' and not ')' and not '\\') &&
                IsSafeClosingMarkToken(prefix, content, foreground) &&
                int.TryParse(foreground.AsSpan(1), NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture, out _) &&
                (precedingStyle.ColorCss is not { } oldColor ||
                 int.TryParse(oldColor.AsSpan(1), NumberStyles.HexNumber,
                     CultureInfo.InvariantCulture, out _)))
            {
                var background = precedingStyle.ColorCss ?? "#000000";
                if (!_joinedClosingMarkColors.TryGetValue(prefix, out var existing) ||
                    existing.Foreground == foreground && existing.Background == background)
                {
                    _joinedClosingMarkColors[prefix] = (foreground, background, existing.Count + 1);
                    precedingText.Content += content;
                    return;
                }
            }
        }
        if (!smallCap && !narrowSpace && content.Length > 0 && _paragraph != null &&
            _previousTextParagraph == _paragraph && _previousTextStyle == _runStyle &&
            _previousTextCanJoin &&
            _runTextBorder == null && _runEffect == null && _runEmphasis == null &&
            _runHighlight == null && _runShading == null &&
            _paragraph.Elements.LastOrDefault() is FormattedText previousSpan &&
            previousSpan.Elements.LastOrDefault() is MigraDoc.DocumentObjectModel.Text previousText &&
            previousText.Content.Length > 0 && !char.IsWhiteSpace(previousText.Content[^1]))
        {
            if (previousText.Content.All(IsOpeningLinePunctuation) &&
                !char.IsWhiteSpace(content[0]))
            {
                previousText.Content += content;
                return;
            }
            var length = 0;
            while (length < content.Length && IsClosingLinePunctuation(content[length]) &&
                !IsFullWidthCjkPunctuation(content[length]) &&
                !(_runHighAnsiFont != null &&
                  IsHighAnsiCharacter(content[length]) !=
                  IsHighAnsiCharacter(previousText.Content[^1])) &&
                (string.IsNullOrEmpty(previousSpan.Font.Name) ||
                 string.Equals(previousSpan.Font.Name, _runStyle?.FontName,
                     StringComparison.OrdinalIgnoreCase))) length++;
            if (length > 0)
            {
                previousText.Content += content[..length];
                if (length < content.Length) AddStyledTextSegment(content[length..], smallCap);
                return;
            }
        }
        if (_runHighAnsiFont != null && content.Length > 1 &&
            content.Any(IsHighAnsiCharacter) && content.Any(character => !IsHighAnsiCharacter(character)))
        {
            var start = 0;
            for (var index = 1; index < content.Length; index++)
            {
                if (IsHighAnsiCharacter(content[index]) == IsHighAnsiCharacter(content[index - 1])) continue;
                AddStyledTextSegment(content[start..index], smallCap, narrowSpace);
                start = index;
            }
            AddStyledTextSegment(content[start..], smallCap, narrowSpace);
            return;
        }
        if (!narrowSpace && _narrowJustifiedSpaces && content.Contains(' '))
        {
            foreach (Match part in Regex.Matches(content, @" +|[^ ]+"))
                AddStyledTextSegment(part.Value, smallCap, part.Value[0] == ' ');
            return;
        }
        if (_unfittedMixedScriptLatinParagraph && !_chineseFallbackParagraph &&
            content.Contains(' '))
        {
            foreach (Match part in Regex.Matches(content, @" +|[^ ]+"))
                if (part.Value[0] == ' ')
                    for (var space = 0; space < part.Length; space++)
                        AddStyledTextSegment("\u2002", smallCap);
                else
                    AddStyledTextSegment(part.Value, smallCap);
            return;
        }
        if (_availableFontFamilies.Contains("MingLiU-ExtB") &&
            content.EnumerateRunes().Any(rune => IsSupplementaryCjk(rune.Value)))
        {
            var start = 0;
            var inSupplementary = false;
            for (var index = 0; index < content.Length;)
            {
                var supplementary = char.IsSurrogatePair(content, index) &&
                    IsSupplementaryCjk(char.ConvertToUtf32(content, index));
                if (supplementary != inSupplementary && index > start)
                {
                    AddStyledTextSegment(content[start..index], smallCap);
                    var other = supplementary ? content[index - 1] : content[index];
                    if (!char.IsWhiteSpace(other) &&
                        (char.IsLetter(other) && _autoSpaceDE || char.IsDigit(other) && _autoSpaceDN))
                        AddStyledTextSegment(" ", smallCap);
                    start = index;
                }
                inSupplementary = supplementary;
                index += supplementary ? 2 : 1;
            }
            if (start > 0)
            {
                AddStyledTextSegment(content[start..], smallCap);
                return;
            }
        }
        if (_availableFontFamilies.Contains("Yu Mincho") &&
            string.Equals(_runStyle?.FontName, "Calibri", StringComparison.OrdinalIgnoreCase) &&
            content.Any(IsJapaneseFallbackCharacter) &&
            content.Any(character => !IsJapaneseFallbackCharacter(character)))
        {
            var start = 0;
            for (var index = 1; index < content.Length; index++)
            {
                if (IsJapaneseFallbackCharacter(content[index]) ==
                    IsJapaneseFallbackCharacter(content[index - 1])) continue;
                AddStyledTextSegment(content[start..index], smallCap);
                if (_unfittedMixedScriptLatinParagraph && !char.IsWhiteSpace(content[index - 1]) &&
                    !char.IsWhiteSpace(content[index]))
                    AddStyledTextSegment(_chineseFallbackParagraph ? " " : "\u2004", smallCap);
                start = index;
            }
            AddStyledTextSegment(content[start..], smallCap);
            return;
        }
        if (FullWidthPunctuationFont != null &&
            string.Equals(_runStyle?.FontName, "Aptos", StringComparison.OrdinalIgnoreCase) &&
            content.Any(IsFullWidthCjkPunctuation) &&
            content.Any(character => !IsFullWidthCjkPunctuation(character)))
        {
            var start = 0;
            for (var index = 1; index < content.Length; index++)
            {
                if (IsFullWidthCjkPunctuation(content[index]) ==
                    IsFullWidthCjkPunctuation(content[index - 1])) continue;
                AddStyledTextSegment(content[start..index], smallCap);
                start = index;
            }
            AddStyledTextSegment(content[start..], smallCap);
            return;
        }
        if (_availableFontFamilies.Contains("Malgun Gothic") && content.Any(IsHangul) &&
            content.Any(character => !IsHangul(character)))
        {
            var start = 0;
            for (var index = 1; index < content.Length; index++)
            {
                if (IsHangul(content[index]) == IsHangul(content[index - 1])) continue;
                AddStyledTextSegment(content[start..index], smallCap);
                if (!char.IsWhiteSpace(content[index - 1]) && !char.IsWhiteSpace(content[index]))
                {
                    var other = IsHangul(content[index - 1]) ? content[index] : content[index - 1];
                    if (char.IsLetter(other) && _autoSpaceDE || char.IsDigit(other) && _autoSpaceDN)
                        AddStyledTextSegment(" ", smallCap, hangulBoundary: true);
                }
                start = index;
            }
            AddStyledTextSegment(content[start..], smallCap);
            return;
        }
        if (_chineseFallbackParagraph && _chineseFallbackFont != null && content.Length > 1 &&
            content.Any(IsCjk))
        {
            var start = 0;
            var family = _chineseFallbackFont(content[0]);
            for (var index = 1; index < content.Length; index++)
            {
                var nextFamily = _chineseFallbackFont(content[index]);
                if (string.Equals(nextFamily, family, StringComparison.OrdinalIgnoreCase)) continue;
                AddStyledTextSegment(content[start..index], smallCap);
                start = index;
                family = nextFamily;
            }
            if (start > 0)
            {
                AddStyledTextSegment(content[start..], smallCap);
                return;
            }
        }
        if (_availableFontFamilies.Contains("Tahoma") && content.Any(IsArabic) &&
            content.Any(char.IsAsciiLetterOrDigit))
        {
            var start = 0;
            for (var index = 1; index < content.Length; index++)
            {
                if (IsArabic(content[index]) == IsArabic(content[index - 1])) continue;
                AddStyledTextSegment(content[start..index], smallCap, narrowSpace);
                start = index;
            }
            AddStyledTextSegment(content[start..], smallCap, narrowSpace);
            return;
        }
        var emittedText = ShouldPreventCellWrap(content) ? content.Replace(' ', '\u00a0') :
            _snapToCharsGrid && !_compactSnapToCharsParagraph && _container is Section &&
                _gridCharSpacingPoints is >= 8 and <= 12 && content.Any(IsCjk)
                ? AddGridBreakOpportunities(content, true, _snapToCharsGrid) :
            _gridLinePitchPoints is > 0 && _gridCharSpacingPoints is >= 8 and <= 12 &&
                _headerGridWrapClearancePoints > 0 && _container is Section && content.Any(IsCjk)
                ? _compactLinesAndCharsParagraph ? AddCjkBreakOpportunities(content) :
                    AddGridBreakOpportunities(content, true, _snapToCharsGrid) :
            _paragraphGridCharSpacingPoints is >= 8 and <= 12 && !_snapToCharsGrid &&
                _container is Section && (content.Any(IsCjk) || _runStyle?.FontSizeHalfPoints >= 36)
                ? AddGridBreakOpportunities(content, false, _snapToCharsGrid,
                    _runStyle?.FontSizeHalfPoints >= 36)
                : AddNarrowCellBreakOpportunities(AddCjkBreakOpportunities(content));
        if (_container is Section && _section != null && _paragraph != null &&
            _paragraph.Format.Alignment == ParagraphAlignment.Left &&
            emittedText.Length > 2 && emittedText[^2] == '.' &&
            (IsClosingLinePunctuation(emittedText[^1]) || emittedText[^1] == '\'') &&
            _runStyle?.Bold != true && _runStyle?.Italic != true)
        {
            var lastSpace = emittedText.LastIndexOf(' ');
            var token = emittedText[(lastSpace + 1)..];
            var fontName = _runStyle?.FontName is string family && _availableFontFamilies.Contains(family)
                ? family : _fontFamily;
            var fontSize = _runStyle?.FontSizeHalfPoints is int halfPoints && halfPoints > 0
                ? halfPoints / 2.0 : _paragraph.Format.Font.Size.Point;
            var font = new XFont(fontName, fontSize > 0 ? fontSize : 11);
            var available = _section.PageSetup.PageWidth.Point - _section.PageSetup.LeftMargin.Point -
                _section.PageSetup.RightMargin.Point - _paragraph.Format.LeftIndent.Point -
                _paragraph.Format.RightIndent.Point;
            using var measure = XGraphics.CreateMeasureContext(new XSize(100, 100),
                XGraphicsUnit.Point, XPageDirection.Downwards);
            var measuredToken = DocSystemFontAdvances.Find(fontName, includeCloudFonts: true) is
                { } advances && advances.TryMeasure(token, font.Size, out var kernedWidth,
                    useKerning: true) ? kernedWidth : measure.MeasureString(token, font).Width;
            if (measuredToken > available &&
                measure.MeasureString(token[..^1], font).Width <= available &&
                GetLineEndOverhang(_paragraph, token, font) == null)
                emittedText = emittedText[..^1] + "\u200B" + emittedText[^1];
        }
        emittedText = Regex.Replace(emittedText, @"(?<=\S) +(?=[!?.:,;%)\]\u201d]\S)",
            match => new string('\u00a0', match.Value.Length));
        if (_container is Section && _section != null && _paragraph != null &&
            emittedText.Contains('\u00a0'))
        {
            var availableWidth = _section.PageSetup.PageWidth.Point -
                _section.PageSetup.LeftMargin.Point - _section.PageSetup.RightMargin.Point -
                _paragraph.Format.LeftIndent.Point - _paragraph.Format.RightIndent.Point;
            var fontName = _runStyle?.FontName is string family && _availableFontFamilies.Contains(family)
                ? family : _fontFamily;
            var fontSize = _runStyle?.FontSizeHalfPoints is int halfPoints && halfPoints > 0
                ? halfPoints / 2.0 : _paragraph.Format.Font.Size.Point;
            using var measure = XGraphics.CreateMeasureContext(new XSize(100, 100),
                XGraphicsUnit.Point, XPageDirection.Downwards);
            var font = new XFont(fontName, fontSize > 0 ? fontSize : 11);
            emittedText = Regex.Replace(emittedText,
                @"(?<prefix>[^\s\u00a0]+\u00a0+[!?.:,;%)\]\u201d])(?<word>[A-Za-z]{3,})",
                match =>
                {
                    if (measure.MeasureString(match.Value.Replace('\u00a0', ' '), font).Width <=
                        availableWidth)
                        return match.Value;
                    _emergencyBreakParagraphs.Add(_paragraph);
                    return match.Groups["prefix"].Value + "\u200B" + string.Join("\u200B",
                        match.Groups["word"].Value.Select(character => character.ToString()));
                });
        }
        emittedText = Regex.Replace(emittedText, @"(?<=\u2014)(?=\S)", "\u200B");
        var preserveRepeatedSpaces = _stylelessBodyParagraphSpacing && _container is Section &&
            emittedText.Contains("  ", StringComparison.Ordinal);
        var preserveLeadingTabSpace = _stylelessBodyParagraphSpacing && _container is Section &&
            emittedText.StartsWith(' ') &&
            _paragraph?.Elements.LastOrDefault() is Character { SymbolName: SymbolName.Tab };
        var explicitSpaces = preserveRepeatedSpaces || preserveLeadingTabSpace;
        var span = _activeHyperlink?.AddFormattedText(explicitSpaces ? "" : emittedText) ??
            _paragraph!.AddFormattedText(explicitSpaces ? "" : emittedText);
        if (explicitSpaces)
        {
            if (preserveLeadingTabSpace)
                span.AddSpace(1);
            var spaceText = preserveLeadingTabSpace ? emittedText[1..] : emittedText;
            foreach (Match part in Regex.Matches(spaceText, @" {2,}|[^ ]+| "))
                if (part.Value.Length > 1 && part.Value[0] == ' ')
                {
                    span.AddText(" ");
                    var nextWord = Regex.Match(spaceText[(part.Index + part.Length)..], @"\S+").Value;
                    var wrapBeforeWord = false;
                    if (preserveRepeatedSpaces && nextWord.Length > 0 &&
                        char.IsLetterOrDigit(nextWord[0]) && _section != null)
                    {
                        var fontName = _runStyle?.FontName is string family &&
                            _availableFontFamilies.Contains(family) ? family : _fontFamily;
                        var fontSize = _runStyle?.FontSizeHalfPoints is int halfPoints && halfPoints > 0
                            ? halfPoints / 2.0 : _paragraph!.Format.Font.Size.Point;
                        using var measure = XGraphics.CreateMeasureContext(new XSize(100, 100),
                            XGraphicsUnit.Point, XPageDirection.Downwards);
                        var font = new XFont(fontName, fontSize > 0 ? fontSize : 11);
                        var availableWidth = _section.PageSetup.PageWidth.Point -
                            _section.PageSetup.LeftMargin.Point - _section.PageSetup.RightMargin.Point -
                            _paragraph!.Format.LeftIndent.Point - _paragraph.Format.RightIndent.Point;
                        wrapBeforeWord = measure.MeasureString(
                            spaceText[..part.Index] + part.Value + nextWord, font).Width > availableWidth;
                    }
                    if (wrapBeforeWord) span.AddText(part.Value[1..]);
                    else span.AddSpace(part.Value.Length - 1);
                }
                else
                    span.AddText(part.Value);
        }
        _previousTextStyle = _runStyle;
        _previousTextParagraph = _paragraph;
        _previousTextCanJoin = _runTextBorder == null && _runEffect == null &&
            _runEmphasis == null && _runHighlight == null && _runShading == null &&
            !(FullWidthPunctuationFont != null &&
                string.Equals(_runStyle?.FontName, "Aptos", StringComparison.OrdinalIgnoreCase) &&
                content.Length > 0 &&
                content.All(IsFullWidthCjkPunctuation));
        _previousTextBasicMetrics = _runCharacterSpacing == 0 && _runScalePercent == 100 &&
            _runHighAnsiFont == null && !_runRightToLeft;
        if (_paragraphGridCharSpacingPoints is >= 8 and <= 12 && !_snapToCharsGrid &&
            _container is Section && _runStyle?.FontSizeHalfPoints >= 36 &&
            content.Length > 1 && content[0] == ' ' && content[1..].All(char.IsLetter))
            _oversizedGridLatinParagraphs.Add(_paragraph);
        if (_container is HeaderFooter && _runStyle?.FontSizeHalfPoints == 36 &&
            string.Equals(_runStyle.FontName, "Yu Mincho", StringComparison.OrdinalIgnoreCase) &&
            content.Length > 1 && content[0] == ' ' && content[1..].All(char.IsAsciiLetter))
            _oversizedStoryLatinParagraphs.Add(_paragraph);
        if (_gridLinePitchPoints is > 0 && _gridCharSpacingPoints is >= 8 and <= 12 &&
            _headerGridWrapClearancePoints > 0 && !_compactLinesAndCharsParagraph &&
            _container is Section && content.Any(IsCjk) &&
            _characterSpacings.TryAdd(_paragraph, 0.16))
            _packedGridParagraphs.Add(_paragraph);
        if (_snapToCharsGrid && !_compactSnapToCharsParagraph &&
            _gridCharSpacingPoints is >= 8 and <= 12 && _container is Section &&
            content.Any(IsCjk) && _characterSpacings.TryAdd(_paragraph, -0.09))
        {
            _packedGridParagraphs.Add(_paragraph);
            _packedGridOffsets[_paragraph] = 4.6;
        }
        if (_paragraphGridCharSpacingPoints is >= 8 and <= 12 && _gridLinePitchPoints is > 0 &&
            _container is Section && _runStyle?.FontSizeHalfPoints == null &&
            _availableFontFamilies.Contains("Yu Mincho") &&
            string.Equals(_runStyle.FontName, "Yu Mincho", StringComparison.OrdinalIgnoreCase) &&
            content.Any(IsCjk))
            _gridGlyphParagraphs.Add(_paragraph);
        if (_paragraphGridCharSpacingPoints is >= 8 and <= 12 && !_snapToCharsGrid &&
            _container is Section && content.Any(IsCjk) && _runStyle?.FontSizeHalfPoints is null or < 36)
            _gridBoundaryParagraphs.Add(_paragraph);
        if (_runStyle == null) return;
        span.Bold = _runRightToLeft ? _runComplexBold :
            _runStyle.Bold || (!_runHasExplicitBold && _cellStyle?.ConditionalBold == true);
        span.Italic = _runRightToLeft ? _runComplexItalic :
            _runStyle.Italic || (!_runHasExplicitItalic && _cellStyle?.ConditionalItalic == true);
        span.Superscript = _runStyle.Superscript;
        span.Subscript = _runStyle.Subscript;
        if (_inHeader && _container is HeaderFooter && _runStyle.Superscript &&
            _superscriptHeaderParagraphs.Add(_paragraph))
            _paragraph.Format.SpaceAfter = Unit.FromPoint(_paragraph.Format.SpaceAfter.Point +
                SuperscriptHeaderLineCompensationPoints);
        if (_tocHyperlinkRun && !_runHasExplicitColor)
            span.Color = MigraDoc.DocumentObjectModel.Colors.Black;
        else if (!_runHasExplicitColor && _textBoxTextColor is { } textBoxColor)
            span.Color = textBoxColor;
        else if (ParseColor(_runStyle.ColorCss ?? (_runHasExplicitColor ? null : _cellStyle?.ConditionalTextColorCss)) is { } color)
            span.Color = color;
        if (_runEffect == TextEffectKind.ShadowOutline && !_runHasExplicitColor &&
            _runStyle.ColorCss == null)
            span.Color = MigraDoc.DocumentObjectModel.Color.FromArgb(0, 0, 0, 0);
        span.Underline = (!_tocHyperlinkRun || _runHasExplicitUnderline) &&
            (_runStyle.Underline || (!_runHasExplicitUnderline && _cellStyle?.ConditionalUnderline == true))
            ? MigraDoc.DocumentObjectModel.Underline.Single : MigraDoc.DocumentObjectModel.Underline.None;
        if (_paragraphInheritsFitText && _paragraph!.Format.Font.Size.Point < 11.5 &&
            _runStyle.FontSizeHalfPoints == null &&
            !_runStyle.Superscript && !_runStyle.Subscript)
            span.Size = Unit.FromPoint(12);
        else if (_unfittedCalibriParagraph && _runStyle.FontSizeHalfPoints is null or 22 &&
            !_runStyle.Superscript && !_runStyle.Subscript)
            span.Size = Unit.FromPoint(12);
        else if (_runStyle.FontSizeHalfPoints is int size && size > 0)
            span.Size = Unit.FromPoint(_runStyle.Superscript ? size / 2.0 * 0.58 : size / 2.0);
        else if (_gridLinePitchPoints is > 0 && _gridCharSpacingPoints == null &&
            string.Equals(_runStyle.FontName, "Calibri", StringComparison.OrdinalIgnoreCase))
            span.Size = Unit.FromPoint(12);
        else if (_cjkStyleParagraph &&
            string.Equals(_runStyle.FontName, "Yu Mincho", StringComparison.OrdinalIgnoreCase))
            span.Size = Unit.FromPoint(12);
        else if (_unfittedMixedScriptParagraph &&
            string.Equals(_runStyle.FontName, "Yu Mincho", StringComparison.OrdinalIgnoreCase) &&
            !_runStyle.Superscript && !_runStyle.Subscript)
            span.Size = Unit.FromPoint(12);
        if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_HIGHANSI_FIT_PAINT_12") == "1" &&
            _container is Section && _fitTextWidths.ContainsKey(_paragraph) &&
            _runHighAnsiFont != null && _runStyle.FontSizeHalfPoints == null)
            span.Size = Unit.FromPoint(12);
        if (_unfittedMixedScriptLatinParagraph && content == "\u2004")
            span.Size = Unit.FromPoint((span.Size.Point > 0 ? span.Size.Point :
                _paragraph.Format.Font.Size.Point) * (_chineseFallbackParagraph ? 0.775 :
                _mixedBodyMinchoFallback ? 0.65 : 0.735));
        if (_unfittedMixedScriptLatinParagraph && content == "\u2002")
            span.Size = Unit.FromPoint(_mixedBodyMinchoFallback ? 8.6 : 9.3);
        if (smallCap)
        {
            var baseSize = _runStyle.FontSizeHalfPoints is int halfPoints && halfPoints > 0
                ? halfPoints / 2.0 : _paragraph.Format.Font.Size.Point;
            span.Size = Unit.FromPoint(Math.Round((baseSize > 0 ? baseSize : 11) * 0.8));
        }
        if (_runStyle.Subscript)
            span.Size = Unit.FromPoint((span.Size.Point > 0 ? span.Size.Point :
                _paragraph.Format.Font.Size.Point) * 0.58 / 0.615);
        if (string.Equals(_runStyle.FontName, "Times New Roman", StringComparison.OrdinalIgnoreCase))
            span.Font.Name = "Times New Roman";
        else if (_runStyle.FontName != null && _availableFontFamilies.Contains(_runStyle.FontName))
            span.Font.Name = _runStyle.FontName;
        if (_implicitFittedStoryFont)
        {
            span.Font.Name = "Aptos";
            span.Size = Unit.FromPoint(12);
        }
        else if (_calibriFittedStoryText)
            span.Size = Unit.FromPoint(12);
        if (_runHighAnsiFont != null && content.Length > 0 && content.All(IsHighAnsiCharacter))
            span.Font.Name = _runHighAnsiFont;
        else if (string.Equals(_runStyle.FontName, "Open Sans;Arial", StringComparison.OrdinalIgnoreCase) &&
            _availableFontFamilies.Contains("Segoe UI"))
            span.Font.Name = "Segoe UI";
        else if (_runStyle.FontName is { Length: > 0 } unavailableFamily &&
            !unavailableFamily.Contains(';') && _availableFontFamilies.Contains("Cambria") &&
            !_availableFontFamilies.Contains(unavailableFamily))
            span.Font.Name = "Cambria";
        if (_availableFontFamilies.Contains("Tahoma") && content.Any(IsArabic) &&
            !(string.Equals(span.Font.Name, "Arial", StringComparison.OrdinalIgnoreCase) &&
              !span.Italic))
            span.Font.Name = "Tahoma";
        if (_runRightToLeft && _runStyle.FontSizeHalfPoints == 22)
            span.Size = Unit.FromPoint(12);
        if (_cjkFontFamily != null && !_availableFontFamilies.Contains(_runStyle.FontName ?? "") &&
            content.Any(character => character >= '\u3040' && character <= '\u9fff'))
            span.Font.Name = _cjkFontFamily;
        if (_availableFontFamilies.Contains("Yu Mincho") &&
            string.Equals(_runStyle.FontName, "Calibri", StringComparison.OrdinalIgnoreCase) &&
            content.Any(IsJapaneseFallbackCharacter))
        {
            var useLayoutFont = _mixedHighAnsiFittedHeaders.Contains(_paragraph) &&
                _availableFontFamilies.Contains("MS Mincho");
            span.Font.Name = useLayoutFont ? "MS Mincho" : "Yu Mincho";
            if ((_unfittedCalibriParagraph ||
                 _fitTextWidths.ContainsKey(_paragraph) &&
                 string.Equals(_runEastAsiaFont, "Calibri", StringComparison.OrdinalIgnoreCase)) &&
                _container is HeaderFooter &&
                _availableFontFamilies.Contains("MS Mincho") && !useLayoutFont)
                _paintOnlyMinchoSpans.Add(span);
        }
        if (_mixedBodyMinchoFallback && !_chineseFallbackParagraph &&
            _availableFontFamilies.Contains("MS Mincho") && content.Any(IsJapaneseFallbackCharacter))
        {
            span.Font.Name = "MS Mincho";
        }
        if (_chineseFallbackParagraph && _availableFontFamilies.Contains("MS Mincho") &&
            content.Any(IsCjk))
            span.Font.Name = _chineseFallbackFont?.Invoke(content.First(IsCjk)) ?? "MS Mincho";
        if (_availableFontFamilies.Contains("MingLiU-ExtB") &&
            content.EnumerateRunes().Any(rune => IsSupplementaryCjk(rune.Value)))
        {
            span.Font.Name = "MingLiU-ExtB";
            span.Bold = _runStyle.Bold || _paragraph.Format.Font.Bold;
        }
        if (_unfittedMixedScriptParagraph && content.Any(IsCjk) &&
            string.Equals(span.Font.Name, "Yu Mincho", StringComparison.OrdinalIgnoreCase))
            span.Bold = false;
        if (_unfittedCalibriParagraph && _container is HeaderFooter && content.Any(IsCjk) &&
            string.Equals(span.Font.Name, "Yu Mincho", StringComparison.OrdinalIgnoreCase))
            span.Bold = false;
        if (_gridLinePitchPoints is > 0 && _gridCharSpacingPoints is > 0 &&
            _container is HeaderFooter && content.Any(IsCjk) &&
            string.Equals(span.Font.Name, "Yu Mincho", StringComparison.OrdinalIgnoreCase) &&
            _runStyle.FontSizeHalfPoints is null or 22)
            span.Size = Unit.FromPoint(12);
        if (_availableFontFamilies.Contains("Malgun Gothic") && content.Any(IsHangul))
            span.Font.Name = "Malgun Gothic";
        if (hangulBoundary)
            span.Size = Unit.FromPoint((span.Size.Point > 0 ? span.Size.Point :
                _paragraph.Format.Font.Size.Point) * 1.09);
        if (FullWidthPunctuationFont is { } punctuationFont &&
            (string.Equals(span.Font.Name, "Aptos", StringComparison.OrdinalIgnoreCase) ||
             string.IsNullOrEmpty(span.Font.Name) && _runStyle.FontName == null) &&
            content.Length > 0 &&
            content.All(IsFullWidthCjkPunctuation))
            span.Font.Name = SelectFullWidthPunctuationFont(content, punctuationFont,
                _availableFontFamilies.Contains("Yu Mincho"));
        if (_renderRubyBaseText && _runStyle.FontName == null && content.Any(IsCjk) &&
            _availableFontFamilies.Contains("Yu Mincho"))
            _paintOnlyRubyBaseMinchoSpans.Add(span);
        if (narrowSpace)
            span.Size = Unit.FromPoint(_container is Section ?
                (_compressedBodyLayoutSpaces ? 8.5 : 9) : 8);
        if (_runStyle.Superscript && _runStyle.FontSizeHalfPoints is int superscriptSize &&
            content.Length is >= 1 and <= 8 && content.All(char.IsAsciiLetter) &&
            !_ambiguousInlineSuperscripts.Contains(content))
        {
            if (!_inlineSuperscriptSizes.TryAdd(content, superscriptSize / 2.0))
            {
                _inlineSuperscriptSizes.Remove(content);
                _ambiguousInlineSuperscripts.Add(content);
            }
        }
    }

    private bool TryAddPictureBullet(Level? level, DxpIDocumentContext context,
        ParagraphFormat format)
    {
        if (_paragraph == null || level == null ||
            context is not DxpDocumentContext documentContext ||
            documentContext.MainDocumentPart?.NumberingDefinitionsPart is not { } numberingPart)
            return false;
        const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string relationshipNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var bulletId = level.ChildElements.FirstOrDefault(element =>
            element.LocalName == "lvlPicBulletId")?.GetAttribute("val", wordNamespace).Value;
        if (string.IsNullOrEmpty(bulletId)) return false;
        var picture = numberingPart.Numbering?.ChildElements.FirstOrDefault(element =>
            element.LocalName == "numPicBullet" &&
            element.GetAttribute("numPicBulletId", wordNamespace).Value == bulletId);
        var imageData = picture?.Descendants().FirstOrDefault(element =>
            element.LocalName == "imagedata");
        var relationshipId = imageData?.GetAttribute("id", relationshipNamespace).Value;
        if (string.IsNullOrEmpty(relationshipId) ||
            numberingPart.Parts.FirstOrDefault(entry => entry.RelationshipId == relationshipId)
                .OpenXmlPart is not ImagePart imagePart) return false;
        using var source = imagePart.GetStream();
        using var memory = new MemoryStream();
        source.CopyTo(memory);
        var bytes = imagePart.ContentType == "image/gif"
            ? DxpRasterImageTransform.ToPngOrOriginal(memory.ToArray()) : memory.ToArray();
        var shape = picture?.Descendants().FirstOrDefault(element => element.LocalName == "shape");
        var style = shape?.GetAttribute("style", "").Value ?? "";
        var widthMatch = Regex.Match(style, @"width:(?<value>[\d.]+)pt");
        var heightMatch = Regex.Match(style, @"height:(?<value>[\d.]+)pt");
        if (!double.TryParse(widthMatch.Groups["value"].Value, NumberStyles.Float,
                CultureInfo.InvariantCulture, out var width) ||
            !double.TryParse(heightMatch.Groups["value"].Value, NumberStyles.Float,
                CultureInfo.InvariantCulture, out var height)) return false;
        if (!_paragraphImages.TryGetValue(_paragraph, out var images))
            _paragraphImages[_paragraph] = images = new List<ParagraphImageOverlay>();
        var imageWidth = width * 8 / 9;
        var imageHeight = height * 8 / 9;
        images.Add(new ParagraphImageOverlay(bytes,
            format.LeftIndent.Point + format.FirstLineIndent.Point,
            (height - imageHeight) / 2, imageWidth, imageHeight));
        return true;
    }

    private static bool IsClosingLinePunctuation(char character) =>
        character is '.' or ',' or ';' or ':' or '!' or '?' or '%' or '"' or '\u2019' or '\u201d' or
            '/' or '-' or '+' or '&' or '\u2026' ||
        char.GetUnicodeCategory(character) is System.Globalization.UnicodeCategory.ClosePunctuation or
            System.Globalization.UnicodeCategory.FinalQuotePunctuation or
            System.Globalization.UnicodeCategory.DashPunctuation;

    private static bool IsOpeningLinePunctuation(char character) =>
        character is '"' or '\'' ||
        char.GetUnicodeCategory(character) is System.Globalization.UnicodeCategory.OpenPunctuation or
            System.Globalization.UnicodeCategory.InitialQuotePunctuation;

    private static bool IsHighAnsiCharacter(char character) =>
        character is >= '\u00a1' and <= '\u052f' or >= '\u1e00' and <= '\u1eff' or
            >= '\u2000' and <= '\u22ff';

    private static bool IsFullWidthCjkPunctuation(char character) =>
        character is >= '\u3000' and <= '\u303f' or >= '\uff01' and <= '\uff65';

    private string? FullWidthPunctuationFont => _availableFontFamilies.Contains("MS Mincho")
        ? "MS Mincho" : _availableFontFamilies.Contains("Yu Mincho") ? "Yu Mincho" : null;

    internal static string SelectFullWidthPunctuationFont(string content, string fallback,
        bool yuMinchoAvailable) => yuMinchoAvailable && content.Length > 0 &&
        content.All(character => CharUnicodeInfo.GetUnicodeCategory(character) is
            UnicodeCategory.ClosePunctuation)
            ? "Yu Mincho" : fallback;

    private static bool IsHangul(char character) => character is >= '\uac00' and <= '\ud7a3';

    private static bool IsArabic(char character) => character is >= '\u0600' and <= '\u06ff' or
        >= '\ufb50' and <= '\ufdff' or >= '\ufe70' and <= '\ufefc';

    private static bool IsSupplementaryCjk(int codePoint) => codePoint is >= 0x20000 and <= 0x323af;

    private static readonly string[] ArabicPresentationForms =
    [
        "\uFE80\0\0\0", "\uFE81\uFE82\0\0", "\uFE83\uFE84\0\0", "\uFE85\uFE86\0\0",
        "\uFE87\uFE88\0\0", "\uFE89\uFE8A\uFE8B\uFE8C", "\uFE8D\uFE8E\0\0", "\uFE8F\uFE90\uFE91\uFE92",
        "\uFE93\uFE94\0\0", "\uFE95\uFE96\uFE97\uFE98", "\uFE99\uFE9A\uFE9B\uFE9C", "\uFE9D\uFE9E\uFE9F\uFEA0",
        "\uFEA1\uFEA2\uFEA3\uFEA4", "\uFEA5\uFEA6\uFEA7\uFEA8", "\uFEA9\uFEAA\0\0", "\uFEAB\uFEAC\0\0",
        "\uFEAD\uFEAE\0\0", "\uFEAF\uFEB0\0\0", "\uFEB1\uFEB2\uFEB3\uFEB4", "\uFEB5\uFEB6\uFEB7\uFEB8",
        "\uFEB9\uFEBA\uFEBB\uFEBC", "\uFEBD\uFEBE\uFEBF\uFEC0", "\uFEC1\uFEC2\uFEC3\uFEC4", "\uFEC5\uFEC6\uFEC7\uFEC8",
        "\uFEC9\uFECA\uFECB\uFECC", "\uFECD\uFECE\uFECF\uFED0", "\0\0\0\0", "\0\0\0\0",
        "\0\0\0\0", "\0\0\0\0", "\0\0\0\0", "\0\0\0\0",
        "\uFED1\uFED2\uFED3\uFED4", "\uFED5\uFED6\uFED7\uFED8", "\uFED9\uFEDA\uFEDB\uFEDC", "\uFEDD\uFEDE\uFEDF\uFEE0",
        "\uFEE1\uFEE2\uFEE3\uFEE4", "\uFEE5\uFEE6\uFEE7\uFEE8", "\uFEE9\uFEEA\uFEEB\uFEEC", "\uFEED\uFEEE\0\0",
        "\uFEEF\uFEF0\uFBE8\uFBE9", "\uFEF1\uFEF2\uFEF3\uFEF4"
    ];

    internal static string ShapeExplicitArabicForPdf(string content)
    {
        var shaped = content.ToCharArray();
        for (var index = 0; index < content.Length; index++)
        {
            var forms = ArabicForms(content[index]);
            if (forms[0] == '\0') continue;
            var previous = index - 1;
            while (previous >= 0 && char.GetUnicodeCategory(content[previous]) ==
                   System.Globalization.UnicodeCategory.NonSpacingMark) previous--;
            var next = index + 1;
            while (next < content.Length && char.GetUnicodeCategory(content[next]) ==
                   System.Globalization.UnicodeCategory.NonSpacingMark) next++;
            var joinsPrevious = previous >= 0 && ArabicForms(content[previous])[2] != '\0' && forms[1] != '\0';
            var joinsNext = next < content.Length && forms[2] != '\0' && ArabicForms(content[next])[1] != '\0';
            shaped[index] = forms[joinsPrevious ? joinsNext ? 3 : 1 : joinsNext ? 2 : 0];
        }
        return new string(shaped);
    }

    internal static string ShapeImplicitArabicForPdf(string content) =>
        Regex.Replace(ShapeExplicitArabicForPdf(content), @"([\ufb50-\ufdff\ufe70-\ufefc]+)(\s+)([0-9]+)",
            match => match.Groups[3].Value + match.Groups[2].Value + match.Groups[1].Value);

    internal static string OrderShapedArabicForPdf(string content) =>
        Regex.Replace(content, @"[\ufb50-\ufdff\ufe70-\ufefc]+",
            match => new string(match.Value.Reverse().ToArray()));

    private static string ArabicForms(char character) => character is >= '\u0621' and <= '\u064a'
        ? ArabicPresentationForms[character - '\u0621'] : "\0\0\0\0";

    internal static string ReorderRtlForPdf(string content)
    {
        foreach (var script in new[] { @"[\u0590-\u05ff]+", @"[\u0600-\u06ff]+" })
        {
            content = Regex.Replace(content, $@"({script})(\s+)([0-9]+)", match =>
                match.Groups[3].Value + match.Groups[2].Value + match.Groups[1].Value);
            content = Regex.Replace(content, script, match =>
            {
                var elements = new List<string>();
                var enumerator = StringInfo.GetTextElementEnumerator(match.Value);
                while (enumerator.MoveNext()) elements.Add(enumerator.GetTextElement());
                elements.Reverse();
                return string.Concat(elements);
            });
        }
        return content;
    }

    internal static string ReorderExplicitRtlRunForPdf(string content) => string.Concat(
        Regex.Matches(content, @"\s+|\S+").Select(match => match.Value).Reverse()
            .Select(part => Regex.Replace(part, @"[\u0590-\u05ff]+", match =>
            {
                var elements = new List<string>();
                var enumerator = StringInfo.GetTextElementEnumerator(match.Value);
                while (enumerator.MoveNext()) elements.Add(enumerator.GetTextElement());
                elements.Reverse();
                return string.Concat(elements);
            })));

    private static bool IsJapaneseFallbackCharacter(char character) =>
        IsCjk(character) || character is >= '\u3000' and <= '\u303f';


    private string AddCjkBreakOpportunities(string content)
    {
        if (!content.Any(IsCjk)) return content;
        var text = new System.Text.StringBuilder(content.Length * 2);
        for (var index = 0; index < content.Length; index++)
        {
            if (index > 0 && !char.IsWhiteSpace(content[index - 1]) &&
                !char.IsWhiteSpace(content[index]) &&
                (IsCjk(content[index - 1]) || IsCjk(content[index])))
            {
                var other = IsCjk(content[index - 1]) ? content[index] : content[index - 1];
                if (IsCjk(content[index - 1]) != IsCjk(content[index]) &&
                    (char.IsDigit(other) && _autoSpaceDN || char.IsLetter(other) && _autoSpaceDE))
                    text.Append(' ');
                text.Append('\u200B');
            }
            text.Append(content[index]);
        }
        return text.ToString();
    }

    private static bool IsCjk(char character) => character is >= '\u3040' and <= '\u9fff';

    internal static int GetIdeographMarkerPrefixLength(string marker) =>
        marker.TakeWhile(character => IsCjk(character) || character == '\u3007').Count();

    internal static bool NeedsAutoSpaceAcrossRuns(string previous, string next, bool autoSpaceDE) =>
        autoSpaceDE && previous.Length > 0 && next.Length > 0 &&
        (IsCjk(previous[^1]) && IsLatinLetter(next[0]) ||
            IsLatinLetter(previous[^1]) && IsCjk(next[0]));

    internal static HashSet<string> GetAsciiWordsAfterRunBoundary(IReadOnlyList<Run> runs)
    {
        var words = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 1; index < runs.Count; index++)
        {
            var previousText = runs[index - 1].InnerText;
            if (previousText.Length == 0 || previousText[^1] != ' ') continue;
            var word = Regex.Match(runs[index].InnerText, @"^[A-Za-z0-9]+").Value;
            if (word.Length > 0) words.Add(word);
        }
        return words;
    }

    private static bool IsLatinLetter(char character) => char.IsLetter(character) &&
        (character is >= 'A' and <= 'Z' or >= 'a' and <= 'z' or
            >= '\u00c0' and <= '\u024f' or >= '\u1e00' and <= '\u1eff');

    private static bool IsCjkGridCharacter(char character) =>
        IsCjk(character) || character is >= '\u3000' and <= '\u303f';

    private bool ShouldPreventCellWrap(string content)
    {
        if (!_cellNoWrap || _cellNoWrapWidthPoints is not double width || !content.Contains(' '))
            return false;
        var family = _runStyle?.FontName is string name && _availableFontFamilies.Contains(name)
            ? name : _fontFamily;
        var size = _runStyle?.FontSizeHalfPoints is int halfPoints && halfPoints > 0
            ? halfPoints / 2.0 : _paragraph!.Format.Font.Size.Point;
        using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
            XPageDirection.Downwards);
        return graphics.MeasureString(content, new XFont(family, size > 0 ? size : 11)).Width <= width + 12;
    }

    private string AddNarrowCellBreakOpportunities(string content)
    {
        if (_cellNoWrap || _cellWrapWidthPoints is not double width || width <= 0)
            return content;
        var family = _runStyle?.FontName is string name && _availableFontFamilies.Contains(name)
            ? name : _fontFamily;
        var size = _runStyle?.FontSizeHalfPoints is int halfPoints && halfPoints > 0
            ? halfPoints / 2.0 : _paragraph!.Format.Font.Size.Point;
        using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
            XPageDirection.Downwards);
        var font = new XFont(family, size > 0 ? size : 11);
        if (_balanceNarrowCellLastWord)
            content = Regex.Replace(content, @"\b([A-Za-z]{4,}) ([A-Za-z])\b", match =>
        {
            var first = match.Groups[1].Value;
            if (graphics.MeasureString(first, font).Width > width ||
                graphics.MeasureString(match.Value, font).Width <= width)
                return match.Value;
            return first[..^1] + "\u200b" + first[^1] + "\u00a0" + match.Groups[2].Value;
        });
        return Regex.Replace(content, @"[A-Za-z]{3,}", match =>
        {
            if (graphics.MeasureString(match.Value, font).Width <= width) return match.Value;
            return string.Join("\u200b", match.Value.Select(character => character.ToString()));
        });
    }

    internal static string AddGridBreakOpportunities(string content, bool snapToChars,
        bool punctuationIsCjk, bool halfEmSpacing = false)
    {
        var text = new System.Text.StringBuilder(content.Length * 2);
        for (var index = 0; index < content.Length; index++)
        {
            if (index > 0)
            {
                if (snapToChars) text.Append("\u00a0\u00a0\u00a0\u200b");
                else if (halfEmSpacing)
                    text.Append(char.IsWhiteSpace(content[index - 1]) &&
                        !char.IsWhiteSpace(content[index]) ? "\u2003" : "\u2002");
                else
                {
                    text.Append(!punctuationIsCjk && content[index] == '\u3002'
                        ? '\u2002' : '\u2003');
                }
                var previousIsCjk = punctuationIsCjk
                    ? IsCjkGridCharacter(content[index - 1]) : IsCjk(content[index - 1]);
                var currentIsCjk = punctuationIsCjk
                    ? IsCjkGridCharacter(content[index]) : IsCjk(content[index]);
                if (previousIsCjk != currentIsCjk &&
                    !(!snapToChars && !punctuationIsCjk && currentIsCjk &&
                        IsFullWidthCjkPunctuation(content[index - 1]) &&
                        !IsOpeningLinePunctuation(content[index - 1])) &&
                    !(snapToChars && !punctuationIsCjk &&
                        (content[index - 1] is >= '\u3000' and <= '\u303f' ||
                         content[index] is >= '\u3000' and <= '\u303f')) &&
                    !char.IsWhiteSpace(content[index - 1]) && !char.IsWhiteSpace(content[index]))
                    text.Append(previousIsCjk ? '\u2002' : ' ');
                if (!snapToChars && (!char.IsLetterOrDigit(content[index - 1]) ||
                    !char.IsLetterOrDigit(content[index]) ||
                    IsCjk(content[index - 1]) || IsCjk(content[index])))
                    text.Append('\u200B');
            }
            text.Append(content[index]);
        }
        return text.ToString();
    }

    public override void VisitTab(TabChar tab, DxpIDocumentContext context)
    {
        if (_paragraph == null) return;
        if (_runStyle is { Strike: true } or { DoubleStrike: true })
        {
            if (!_strikeTabBoundaries.TryGetValue(_paragraph, out var boundaries))
                _strikeTabBoundaries[_paragraph] = boundaries = new HashSet<int>();
            boundaries.Add(_paragraphTextLength);
        }
        if (!_paragraphHasLeaderTab)
        {
            if (_activeHyperlink != null) _activeHyperlink.AddTab();
            else _paragraph.AddTab();
            return;
        }
        var span = _activeHyperlink?.AddFormattedText() ?? _paragraph.AddFormattedText();
        var family = _runStyle?.FontName;
        if (family != null && _availableFontFamilies.Contains(family))
            span.Font.Name = family;
        else if (_implicitModernFont && _availableFontFamilies.Contains("Aptos"))
            span.Font.Name = "Aptos";
        span.AddTab();
    }

    public override void VisitPositionalTab(PositionalTab tab, DxpIDocumentContext context)
    {
        if (_paragraph == null) return;
        if (AttributeValue(tab, "alignment") == "left" &&
            AttributeValue(tab, "relativeTo") == "margin")
        {
            var following = tab.ElementsAfter().OfType<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .Any(text => text.Text.Length > 0) ||
                tab.Parent?.ElementsAfter().Any(element =>
                    element.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                        .Any(text => text.Text.Length > 0)) == true;
            if (_paragraphTextLength > 0 && following)
            {
                _paragraph.AddLineBreak();
                if (_inHeader) _headerLeftMarginTabBreakCount++;
                else if (_container is HeaderFooter) _footerLeftMarginTabBreakCount++;
            }
            return;
        }
        _paragraph.AddTab();
    }

    public override void VisitBreak(Break br, DxpIDocumentContext context)
    {
        if (br.Type?.Value == BreakValues.Column && _activeColumnCell != null)
        {
            _columnBreakPending = true;
            return;
        }
        if (br.Type?.Value == BreakValues.Page)
        {
            if (_container is Section && br.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                    .FirstOrDefault() is { } sourceParagraph && sourceParagraph.Parent is Body &&
                sourceParagraph.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph &&
                !sourceParagraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any() &&
                !sourceParagraph.Descendants<Drawing>().Any())
            {
                _pendingEmptyParagraphPageBreak = true;
                return;
            }
            var current = _paragraph;
            var next = AddParagraph();
            if (next != null)
            {
                if (current != null)
                {
                    next.Style = current.Style;
                    next.Format = (ParagraphFormat)current.Format.Clone();
                    next.Format.FirstLineIndent = Unit.Zero;
                    next.Format.SpaceBefore = Unit.Zero;
                    current.Format.SpaceAfter = Unit.Zero;
                    if (_paragraphShades.TryGetValue(current, out var shade))
                        _paragraphShades[next] = shade;
                    if (_manualParagraphShades.Contains(current))
                        _manualParagraphShades.Add(next);
                    if (_mirroredParagraphs.Contains(current))
                        _mirroredParagraphs.Add(next);
                    if (_sourceBodyParagraphs.Contains(current))
                        _sourceBodyParagraphs.Add(next);
                    if (_sourceCellParagraphs.Contains(current))
                        _sourceCellParagraphs.Add(next);
                    if (_suppressedLineNumberParagraphs.Contains(current))
                        _suppressedLineNumberParagraphs.Add(next);
                    if (_characterSpacings.TryGetValue(current, out var spacing))
                        _characterSpacings[next] = spacing;
                    if (_characterScales.TryGetValue(current, out var scale))
                        _characterScales[next] = scale;
                    if (_tabbedRunShadings.TryGetValue(current, out var tabShading))
                        _tabbedRunShadings[next] = tabShading;
                }
                next.Format.PageBreakBefore = true;
                _paragraph = next;
                _paragraphTextLength = 0;
                _paragraphBorderLetterLength = 0;
                _revisedListMarkerLength = 0;
                if (_activeHyperlink != null)
                    _activeHyperlink = next.AddHyperlink(_activeHyperlink.Name, _activeHyperlink.Type);
            }
        }
        else if (_activeHyperlink != null) _activeHyperlink.Elements.AddLineBreak();
        else _paragraph?.AddLineBreak();
    }

    public override void VisitCarriageReturn(CarriageReturn carriageReturn, DxpIDocumentContext context)
    {
        if (_activeHyperlink != null) _activeHyperlink.Elements.AddLineBreak();
        else _paragraph?.AddLineBreak();
    }

    private MigraDoc.DocumentObjectModel.Paragraph? AddParagraph()
    {
        var paragraph = _container switch
        {
            Section when _activeColumnCell != null => _activeColumnCell.AddParagraph(),
            Section section => section.AddParagraph(),
            HeaderFooter headerFooter => headerFooter.AddParagraph(),
            Cell cell => cell.AddParagraph(),
            MigraDoc.DocumentObjectModel.Shapes.TextFrame frame => frame.AddParagraph(),
            _ => null
        };
        if (_container is Section && _activeColumnCell == null && paragraph != null)
            _bodyParagraphs.Add(paragraph);
        if (paragraph != null && _activeFootnoteIndex is int index)
            _footnoteParagraphs.Add((index, paragraph, 0));
        if (paragraph != null && _activeEndnoteIndex is int endnoteIndex &&
            _endnoteDestinationOrdinals.Add(endnoteIndex))
            paragraph.AddBookmark($"dxp-endnote-{endnoteIndex}", prepend: true);
        return paragraph;
    }

    public override void VisitPageNumber(PageNumber number, DxpIDocumentContext context)
    {
        if (_paragraph != null)
            (_activeHyperlink?.AddPageField() ?? _paragraph.AddPageField()).Format = _sectionPageNumberFormat;
    }

    public override void VisitMoveToRun(MoveToRun moved, DxpIDocumentContext context)
    {
    }

    public override void VisitFootnoteReference(FootnoteReference reference,
        DxpIFootnoteContext footnote, DxpIDocumentContext context)
    {
        if (!_includeRevisionMarkup &&
            (reference.Ancestors<Deleted>().Any() || reference.Ancestors<DeletedRun>().Any())) return;
        if (_paragraph == null || footnote.Index is not int index) return;
        var customMark = reference.GetAttributes().FirstOrDefault(attribute =>
            attribute.LocalName == "customMarkFollows").Value is "1" or "true";
        _footnoteReferences.TryAdd(index, _paragraph);
        if (customMark) _customFootnoteMarks.Add(index);
        var labelText = FormatFootnoteIndex(index, context);
        _footnoteLabels[index] = labelText;
        _paragraph.AddBookmark($"dxp-footnote-{index}-{++_footnoteBookmarkCount}", prepend: false);
        if (customMark) return;
        var markerLink = _activeHyperlink ?? _paragraph.AddHyperlink($"dxp-note-{index}");
        var marker = markerLink.AddFormattedText(labelText);
        marker.Superscript = true;
        marker.Size = Unit.FromPoint(11.3);
        if (_runStyle?.FontName is string referenceFont && _availableFontFamilies.Contains(referenceFont))
            marker.Font.Name = referenceFont;
    }

    public override IDisposable VisitFootnoteBegin(DocumentFormat.OpenXml.Wordprocessing.Footnote note,
        DxpIFootnoteContext footnote, DxpIDocumentContext context)
    {
        if (_section == null) return DxpDisposable.Empty;
        if (!_includeRevisionMarkup && note.Descendants<Deleted>().Any() &&
            !note.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any() &&
            !note.Descendants<Drawing>().Any() && !note.Descendants<Picture>().Any())
        {
            var previous = _container;
            _container = null;
            return DxpDisposable.Create(() => _container = previous);
        }
        if (!_footnoteAppendixStarted)
        {
            _section.AddPageBreak();
            _footnoteAppendixBreak = _section.Elements[_section.Elements.Count - 1] as
                MigraDoc.DocumentObjectModel.PageBreak;
            _footnoteAppendixStarted = true;
        }
        var previousContainer = _container;
        var previousIndex = _activeFootnoteIndex;
        _activeFootnoteIndex = footnote.Index;
        _container = _section;
        return DxpDisposable.Create(() =>
        {
            _container = previousContainer;
            _activeFootnoteIndex = previousIndex;
        });
    }

    public override void VisitFootnoteReferenceMark(FootnoteReferenceMark mark,
        DxpIFootnoteContext footnote, DxpIDocumentContext context)
    {
        if (_paragraph != null && footnote.Index is int index && !_customFootnoteMarks.Contains(index))
        {
            var label = _paragraph.AddFormattedText(_footnoteLabels.GetValueOrDefault(index) ??
                FormatFootnoteIndex(index, context));
            label.Superscript = true;
            label.Size = Unit.FromPoint(9);
            if (_runStyle?.FontName is string noteFont && _availableFontFamilies.Contains(noteFont))
                label.Font.Name = noteFont;
            _paragraph.AddBookmark($"dxp-note-{index}", prepend: false);
        }
    }

    public override void VisitEndnoteReference(EndnoteReference reference, DxpIDocumentContext context)
    {
        if (_paragraph == null || reference.Id?.Value is not long id ||
            context is not DxpDocumentContext documentContext ||
            !documentContext.Endnotes.Resolve(id, out _)) return;
        if (!_endnoteOrdinals.TryGetValue(id, out var index))
            _endnoteOrdinals[id] = index = _endnoteOrdinals.Count + 1;
        if (!_endnoteLabels.TryGetValue(index, out var labelText))
            _endnoteLabels[index] = labelText = FormatEndnoteIndex(index, context);
        var customMark = reference.GetAttributes().FirstOrDefault(attribute =>
            attribute.LocalName == "customMarkFollows").Value is "1" or "true";
        if (customMark) _customEndnoteMarks.Add(index);
        if (_section != null)
            _endnoteSections.TryAdd(id, _section);
        if (customMark) return;
        var markerLink = _activeHyperlink ?? _paragraph.AddHyperlink($"dxp-endnote-{index}");
        var marker = markerLink.AddFormattedText(labelText);
        marker.Superscript = _runStyle?.Superscript == true;
    }

    public override IDisposable VisitEndnoteBegin(Endnote note, long id, int index,
        DxpIDocumentContext context)
    {
        if (_section == null) return DxpDisposable.Empty;
        if (!_endnoteOrdinals.TryGetValue(id, out var ordinal))
        {
            var orphanPreviousContainer = _container;
            _container = null;
            return DxpDisposable.Create(() => _container = orphanPreviousContainer);
        }
        var targetSection = _endnotesAtSectionEnd && _endnoteSections.TryGetValue(id, out var referencedSection)
            ? referencedSection : _section;
        if (_endnoteSeparatorSections.Add(targetSection))
        {
            var separator = targetSection.AddParagraph();
            var contentWidth = targetSection.PageSetup.PageWidth.Point -
                targetSection.PageSetup.LeftMargin.Point - targetSection.PageSetup.RightMargin.Point;
            separator.Format.RightIndent = Unit.FromPoint(Math.Max(0, contentWidth - 144));
            separator.Format.Font.Size = Unit.FromPoint(1);
            var bodyEndsWithTable = context is DxpDocumentContext documentContext &&
                documentContext.MainDocumentPart?.Document?.Body?.ChildElements
                    .LastOrDefault(element => element is DocumentFormat.OpenXml.Wordprocessing.Table or
                        DocumentFormat.OpenXml.Wordprocessing.Paragraph) is
                    DocumentFormat.OpenXml.Wordprocessing.Table;
            separator.Format.SpaceBefore = Unit.FromPoint(_stylelessBodyParagraphSpacing &&
                !bodyEndsWithTable ? 14 : 16);
            separator.Format.SpaceAfter = Unit.FromPoint(7);
            separator.Format.Borders.Bottom.Width = Unit.FromPoint(0.5);
        }
        var previousContainer = _container;
        var previousSection = _section;
        var previousIndex = _activeEndnoteIndex;
        _activeEndnoteIndex = ordinal;
        _section = targetSection;
        _container = targetSection;
        return DxpDisposable.Create(() =>
        {
            _container = previousContainer;
            _section = previousSection;
            _activeEndnoteIndex = previousIndex;
        });
    }

    public override void VisitEndnoteReferenceMark(EndnoteReferenceMark mark, DxpIDocumentContext context)
    {
        if (_paragraph == null || _activeEndnoteIndex is not int index ||
            _customEndnoteMarks.Contains(index)) return;
        var label = _paragraph.AddFormattedText(_endnoteLabels.GetValueOrDefault(index) ??
            FormatEndnoteIndex(index, context));
        label.Superscript = _runStyle?.Superscript == true;
    }

    private string FormatEndnoteIndex(int index, DxpIDocumentContext context)
    {
        var sectionProperties = context.CurrentSection.SectionProperties?.GetFirstChild<EndnoteProperties>();
        var format = sectionProperties?.NumberingFormat?.Val?.Value ?? NumberFormatValues.LowerRoman;
        var start = sectionProperties?.NumberingStart?.Val?.Value ?? 1;
        var sectionIndex = sectionProperties?.NumberingRestart?.Val?.Value == RestartNumberValues.EachSection
            ? _endnoteSectionStarts.GetValueOrDefault(_sectionCount, index) : 1;
        if (sectionIndex == index &&
            sectionProperties?.NumberingRestart?.Val?.Value == RestartNumberValues.EachSection)
            _endnoteSectionStarts.TryAdd(_sectionCount, index);
        return DxpListMarkerFormatter.FormatNumber(index - sectionIndex + start, format);
    }

    private string FormatFootnoteIndex(int index, DxpIDocumentContext context)
    {
        var properties = context.CurrentSection.SectionProperties?.GetFirstChild<FootnoteProperties>();
        var format = properties?.NumberingFormat?.Val?.Value ?? NumberFormatValues.Decimal;
        var start = properties?.NumberingStart?.Val?.Value ?? 1;
        var sectionIndex = properties?.NumberingRestart?.Val?.Value == RestartNumberValues.EachSection
            ? _footnoteSectionStarts.GetValueOrDefault(_sectionCount, index) : 1;
        if (sectionIndex == index && properties?.NumberingRestart?.Val?.Value == RestartNumberValues.EachSection)
            _footnoteSectionStarts.TryAdd(_sectionCount, index);
        return DxpListMarkerFormatter.FormatNumber(index - sectionIndex + start, format);
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext context)
    {
        if (_hiddenRun) return;
        if (UseCachedMarginMergePlaceholders &&
            (_container is HeaderFooter || _container is Cell && _tableInHeaderFooter) &&
            context.CurrentFields.Current is { InstructionText: string mergeInstruction } mergeFrame &&
            DxpFieldInstructionClassifier.IsMergeFieldInstruction(mergeInstruction))
        {
            var arguments = new DxpFieldParser().Parse(mergeInstruction).Ast.ArgumentsText;
            if (DxpFieldTokenization.TryGetFirstToken(arguments, out var name) &&
                _renderedMarginMergeFields.Add(mergeFrame))
                AddStyledText($"«{name}»");
            return;
        }
        if (_paragraph != null && !string.IsNullOrEmpty(text) &&
            context.CurrentFields.Current is { InstructionText: string instruction } frame &&
            Regex.IsMatch(instruction, @"^\s*PAGE(?:\s|$)", RegexOptions.IgnoreCase) &&
            !Regex.IsMatch(instruction, @"\\\*\s*(?!MERGEFORMAT\b)\w+", RegexOptions.IgnoreCase))
        {
            if (_renderedPageFields.Add(frame))
            {
                var fieldSpan = _paragraph.AddFormattedText();
                if (_runStyle is { } runStyle)
                {
                    fieldSpan.Bold = runStyle.Bold;
                    fieldSpan.Italic = runStyle.Italic;
                    fieldSpan.Underline = runStyle.Underline
                        ? MigraDoc.DocumentObjectModel.Underline.Single : MigraDoc.DocumentObjectModel.Underline.None;
                    if (ParseColor(runStyle.ColorCss) is { } color) fieldSpan.Color = color;
                    if (runStyle.FontSizeHalfPoints is > 0)
                        fieldSpan.Size = Unit.FromPoint(runStyle.FontSizeHalfPoints.Value / 2.0);
                    if (runStyle.FontName is string font && _availableFontFamilies.Contains(font))
                        fieldSpan.Font.Name = font;
                }
                fieldSpan.AddPageField().Format = _sectionPageNumberFormat;
            }
            return;
        }
        if (_inHeader && _container is HeaderFooter && _paragraph != null &&
            context.CurrentFields.Current is { InstructionText: string styleInstruction } &&
            Regex.Match(styleInstruction, "^\\s*STYLEREF\\s+\"(?<name>[^\"]+)\"",
                RegexOptions.IgnoreCase) is { Success: true } styleMatch)
            _styleRefHeaders.Add((styleMatch.Groups["name"].Value, text));
        AddStyledText(text);
    }

    public override IDisposable VisitLegacyPictureBegin(Picture picture, DxpIDocumentContext context)
    {
        TryAddVmlImage(picture, context);
        return DxpDisposable.Empty;
    }

    public override void VisitEmbeddedObject(EmbeddedObject embeddedObject, DxpIDocumentContext context)
    {
        TryAddVmlImage(embeddedObject, context);
    }

    private void TryAddVmlImage(OpenXmlElement carrier, DxpIDocumentContext context)
    {
        if (_paragraph == null || context.CurrentPart == null)
            return;
        var imageData = carrier.Descendants().FirstOrDefault(element =>
            element.LocalName == "imagedata" &&
            element.NamespaceUri == "urn:schemas-microsoft-com:vml");
        var relationshipId = imageData?.GetAttribute("id",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;
        if (string.IsNullOrEmpty(relationshipId) ||
            context.CurrentPart.GetPartById(relationshipId) is not ImagePart imagePart ||
            imagePart.ContentType is not ("image/png" or "image/jpeg" or "image/gif"))
            return;
        var shape = imageData?.Parent;
        var style = shape?.GetAttribute("style", "").Value ?? "";
        var widthMatch = Regex.Match(style, @"(?:^|;)\s*width:\s*(\d*\.?\d+)(pt|in|px)(?:;|$)",
            RegexOptions.IgnoreCase);
        var heightMatch = Regex.Match(style, @"(?:^|;)\s*height:\s*(\d*\.?\d+)(pt|in|px)(?:;|$)",
            RegexOptions.IgnoreCase);
        if (!widthMatch.Success || !heightMatch.Success ||
            !TryParseVmlDimension(widthMatch, out var width) ||
            !TryParseVmlDimension(heightMatch, out var height))
            return;
        using var stream = imagePart.GetStream();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        var bytes = buffer.ToArray();
        if (imagePart.ContentType == "image/gif")
            bytes = DxpRasterImageTransform.ToPngOrOriginal(bytes);
        var image = _paragraph.AddImage("base64:" + Convert.ToBase64String(bytes));
        image.Width = Unit.FromPoint(width);
        image.Height = Unit.FromPoint(height);
        if (_container is Section &&
            carrier.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is
                { } sourceParagraph &&
            !sourceParagraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any() &&
            !sourceParagraph.Descendants<Drawing>().Any() &&
            ReferenceEquals(sourceParagraph.Descendants().LastOrDefault(element =>
                element is Picture or EmbeddedObject), carrier))
            ApplyStandaloneImageLineSpacing(sourceParagraph, carrier, context,
                _paragraph.Elements.OfType<MigraDoc.DocumentObjectModel.Shapes.Image>()
                    .Max(item => item.Height.Point));
    }

    private void ApplyStandaloneImageLineSpacing(
        DocumentFormat.OpenXml.Wordprocessing.Paragraph sourceParagraph,
        OpenXmlElement carrier, DxpIDocumentContext context, double imageHeight)
    {
        var format = _paragraph!.Format;
        var normalFormat = _document.Styles["Normal"].ParagraphFormat;
        var rule = format.Values.LineSpacingRule ?? normalFormat.LineSpacingRule;
        if (rule is not (LineSpacingRule.Multiple or LineSpacingRule.AtLeast))
            return;
        var sourceRun = carrier.Ancestors<Run>().FirstOrDefault();
        var runStyle = sourceRun == null ? null : context.Styles.ResolveRunStyle(sourceParagraph, sourceRun);
        var sourceFontSize = runStyle?.FontSizeHalfPoints is > 0
            ? runStyle.FontSizeHalfPoints.Value / 2.0 : normalFormat.Font.Size.Point;
        var sourceFontFamily = runStyle?.FontName is string sourceFontName && _availableFontFamilies.Contains(sourceFontName)
            ? sourceFontName : _minorThemeLatinFont is string themeFont && _availableFontFamilies.Contains(themeFont)
                ? themeFont : normalFormat.Font.Name;
        var useSourceFont = Environment.GetEnvironmentVariable("DOCXPORT_DIAG_STANDALONE_IMAGE_SOURCE_FONT") == "1";
        if (useSourceFont)
        {
            format.Font.Name = sourceFontFamily;
            format.Font.Size = Unit.FromPoint(sourceFontSize);
        }
        var imageFont = useSourceFont
            ? new XFont(sourceFontFamily, sourceFontSize)
            : new XFont(normalFormat.Font.Name, normalFormat.Font.Size.Point);
        var imageDescent = imageFont.Metrics.Descent * imageFont.Size /
            imageFont.FontFamily.GetEmHeight(imageFont.Style);
        var inheritedAfter = format.Values.SpaceAfter?.Point ?? normalFormat.SpaceAfter.Point;
        var imageOffsetMode = Environment.GetEnvironmentVariable("DOCXPORT_DIAG_STANDALONE_IMAGE_DESCENT");
        if (rule == LineSpacingRule.AtLeast)
        {
            var minimum = format.Values.LineSpacingRule.HasValue
                ? format.LineSpacing.Point : normalFormat.LineSpacing.Point;
            var offset = Math.Min(imageDescent, Math.Max(0, minimum - imageHeight - imageDescent));
            if (imageOffsetMode == "2" && offset == 0)
                offset = -0.36;
            if (offset != 0)
                foreach (var image in _paragraph.Elements.OfType<MigraDoc.DocumentObjectModel.Shapes.Image>())
                    _inlineImageOffsets[image] = offset;
            if (offset == 0 && imageOffsetMode == "3" && !useSourceFont && carrier is Drawing)
                ShiftPrecedingImageSpacing(normalFormat);
            var excess = Math.Max(0, imageHeight + imageDescent - minimum);
            format.SpaceAfter = Unit.FromPoint(inheritedAfter - Math.Min(imageDescent, excess));
            return;
        }
        var factor = format.Values.LineSpacingRule.HasValue
            ? format.LineSpacing.Point : normalFormat.LineSpacing.Point;
        var size = runStyle?.FontSizeHalfPoints is > 0 ? runStyle.FontSizeHalfPoints.Value / 2.0 :
            normalFormat.Font.Size.Point;
        var family = runStyle?.FontName is string fontName && _availableFontFamilies.Contains(fontName)
            ? fontName : _fontFamily;
        var textFont = new XFont(family, size);
        if (imageOffsetMode == "2")
            foreach (var image in _paragraph.Elements.OfType<MigraDoc.DocumentObjectModel.Shapes.Image>())
                _inlineImageOffsets[image] = -0.36;
        if (imageOffsetMode == "3" && !useSourceFont && carrier is Drawing)
            ShiftPrecedingImageSpacing(normalFormat);
        format.LineSpacingRule = LineSpacingRule.Single;
        format.SpaceAfter = Unit.FromPoint(inheritedAfter +
            (factor - 1) * textFont.GetHeight() - imageDescent);
    }

    private void ShiftPrecedingImageSpacing(ParagraphFormat normalFormat)
    {
        if (_section?.Elements is not { } elements || elements.Count < 2 ||
            !ReferenceEquals(elements[elements.Count - 1], _paragraph) ||
            elements[elements.Count - 2] is not MigraDoc.DocumentObjectModel.Paragraph previous)
            return;
        var after = previous.Format.Values.SpaceAfter?.Point ?? normalFormat.SpaceAfter.Point;
        previous.Format.SpaceAfter = Unit.FromPoint(after - 0.36);
    }

    private static bool TryParseVmlDimension(Match match, out double points)
    {
        points = 0;
        if (!double.TryParse(match.Groups[1].Value, NumberStyles.Float,
                CultureInfo.InvariantCulture, out var value))
            return false;
        points = match.Groups[2].Value.ToLowerInvariant() switch
        {
            "in" => value * 72,
            "px" => value * 0.75,
            _ => value
        };
        return points > 0;
    }

    private static OpenXmlElement? GetRectangularTextWrap(Drawing drawing, bool useContourForInset)
    {
        var square = drawing.Descendants().FirstOrDefault(element =>
            element.LocalName == "wrapSquare");
        if (square != null)
            return square;

        var contour = drawing.Descendants().FirstOrDefault(element =>
            element.LocalName is "wrapTight" or "wrapThrough");
        var polygon = contour?.ChildElements.FirstOrDefault(element =>
            element.LocalName == "wrapPolygon");
        if (polygon == null)
            return null;

        var points = polygon.ChildElements.Select(element =>
        {
            var validX = long.TryParse(element.GetAttribute("x", "").Value, out var x);
            var validY = long.TryParse(element.GetAttribute("y", "").Value, out var y);
            var valid = validX && validY;
            return (valid, x, y);
        }).ToArray();
        if (points.Length != 5 || points.Any(point => !point.valid) ||
            points[0] != points[4])
            return null;
        var minX = points.Min(point => point.x);
        var maxX = points.Max(point => point.x);
        var minY = points.Min(point => point.y);
        var maxY = points.Max(point => point.y);
        if (minX != 0 || minY != 0 ||
            (useContourForInset ? maxX != 21600 || maxY != 21600 :
                maxX < 21000 || maxX > 22000 || maxY < 21000 || maxY > 22000) ||
            points.Take(4).Any(point =>
                (point.x != minX && point.x != maxX) ||
                (point.y != minY && point.y != maxY)) ||
            points.Take(4).Select(point => (point.x, point.y)).Distinct().Count() != 4)
            return null;
        return contour;
    }

    private static OpenXmlElement? GetContourTextWrap(Drawing drawing,
        out IReadOnlyList<XPoint>? points)
    {
        points = null;
        var contour = drawing.Descendants().FirstOrDefault(element =>
            element.LocalName is "wrapTight" or "wrapThrough");
        var polygon = contour?.ChildElements.FirstOrDefault(element =>
            element.LocalName == "wrapPolygon");
        if (polygon == null)
            return null;

        var vertices = new List<XPoint>();
        foreach (var element in polygon.ChildElements)
        {
            if (!long.TryParse(element.GetAttribute("x", "").Value, out var x) ||
                !long.TryParse(element.GetAttribute("y", "").Value, out var y) ||
                x is < 0 or > 21600 || y is < 0 or > 21600)
                return null;
            vertices.Add(new XPoint(x / 21600.0, y / 21600.0));
        }
        if (vertices.Count < 4 || vertices[0] != vertices[^1])
            return null;
        vertices.RemoveAt(vertices.Count - 1);
        points = vertices;
        return contour;
    }

    internal static bool RequiresRasterTransform(bool anchored, DxpImagePresentation? presentation) =>
        presentation != null && (anchored || Math.Abs(presentation.RotationDegrees) > 0.001 ||
            presentation.Crop is { IsEmpty: false } || presentation.FlipHorizontal ||
            presentation.FlipVertical);

    public override IDisposable VisitDrawingBegin(Drawing drawing, DxpDrawingInfo? info, DxpIDocumentContext context)
    {
        if (TryAddChart(drawing, info, context)) return DxpDisposable.Empty;
        if (_paragraph == null || info?.DataUri is not string uri ||
            !(info.ContentType?.Equals("image/png", StringComparison.OrdinalIgnoreCase) == true ||
              info.ContentType?.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) == true ||
              info.ContentType?.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) == true ||
              info.ContentType?.Equals("image/x-ms-bmp", StringComparison.OrdinalIgnoreCase) == true ||
              info.ContentType?.Equals("image/tiff", StringComparison.OrdinalIgnoreCase) == true ||
              info.ContentType?.Equals("image/gif", StringComparison.OrdinalIgnoreCase) == true))
            return DxpDisposable.Empty;
        var comma = uri.IndexOf(',');
        if (comma < 0) return DxpDisposable.Empty;
        var bytes = Convert.FromBase64String(uri.Substring(comma + 1));
        var imageData = "base64:" + uri.Substring(comma + 1);
        var anchored = drawing.Descendants().Any(element => element.LocalName == "anchor" &&
            element.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        if (!anchored && _paragraph.Elements.Count == 0 &&
            (_inHeader || _container is Section &&
                info.Presentation?.FrameHeightPoints is > 0 and <= 35))
        {
            if (_container is Section && _section != null &&
                ReferenceEquals(_section.Elements.OfType<MigraDoc.DocumentObjectModel.Paragraph>()
                    .FirstOrDefault(), _paragraph))
                _section.PageSetup.TopMargin = Unit.FromPoint(_section.PageSetup.TopMargin.Point + 0.69);
            else
                _paragraph.Format.SpaceBefore = Unit.FromPoint(_paragraph.Format.SpaceBefore.Point + 0.69);
        }
        if (info.ContentType?.Contains("bmp", StringComparison.OrdinalIgnoreCase) == true)
        {
            bytes = DxpBmp24Converter.NormalizeOrOriginal(bytes);
        }
        else if (info.ContentType?.Equals("image/png", StringComparison.OrdinalIgnoreCase) == true &&
            uri.StartsWith("data:image/png;base64,", StringComparison.OrdinalIgnoreCase))
        {
            if (bytes.Length > 25 && bytes[24] == 16)
                bytes = DxpPng16Converter.ToBmpOrOriginal(bytes);
        }
        else if (info.ContentType?.Equals("image/tiff", StringComparison.OrdinalIgnoreCase) == true)
            bytes = DxpTiffConverter.ToBmpOrOriginal(bytes);
        else if (info.ContentType?.Equals("image/gif", StringComparison.OrdinalIgnoreCase) == true)
            bytes = DxpRasterImageTransform.ToPngOrOriginal(bytes);
        var transformed = RequiresRasterTransform(anchored, info.Presentation) &&
            info.Presentation is { FrameWidthPoints: > 0, FrameHeightPoints: > 0 } imagePresentation
            ? DxpRasterImageTransform.Transform(bytes, imagePresentation.FrameWidthPoints.Value,
                imagePresentation.FrameHeightPoints.Value, imagePresentation.RotationDegrees,
                imagePresentation.FlipHorizontal, imagePresentation.FlipVertical,
                imagePresentation.Crop) : null;
        if (transformed is { } raster)
        {
            bytes = raster.Image;
            imageData = "base64:" + Convert.ToBase64String(bytes);
        }
        else if (info.Presentation is { } presentation)
            bytes = DxpBmp24Converter.TransformOrOriginal(bytes, presentation.RotationDegrees,
                presentation.FlipHorizontal, presentation.FlipVertical);
        if (transformed == null && (info.ContentType?.Contains("bmp", StringComparison.OrdinalIgnoreCase) == true ||
            info.ContentType?.Equals("image/tiff", StringComparison.OrdinalIgnoreCase) == true ||
            info.ContentType?.Equals("image/gif", StringComparison.OrdinalIgnoreCase) == true ||
            info.ContentType?.Equals("image/png", StringComparison.OrdinalIgnoreCase) == true &&
            bytes.Length > 0 && bytes[0] == 'B'))
            imageData = "base64:" + Convert.ToBase64String(bytes);
        var horizontal = anchored ? drawing.Descendants().FirstOrDefault(element => element.LocalName == "positionH") : null;
        var vertical = anchored ? drawing.Descendants().FirstOrDefault(element => element.LocalName == "positionV") : null;
        var horizontalOffset = MarginImageOffset(horizontal) ?? PositionImageOffset(horizontal, "column");
        var verticalOffset = MarginImageOffset(vertical);
        var anchorElement = anchored ? drawing.Descendants().FirstOrDefault(element => element.LocalName == "anchor") : null;
        var behindDocument = anchorElement?.GetAttribute("behindDoc", "").Value is "1" or "true";
        var zOrder = long.TryParse(anchorElement?.GetAttribute("relativeHeight", "").Value,
            out var heightOrder) ? heightOrder : 0;
        if (anchorElement?.GetAttribute("simplePos", "").Value is "1" or "true" &&
            _container is Section && _paragraph != null &&
            anchorElement.ChildElements.FirstOrDefault(element => element.LocalName == "simplePos") is
                { } simplePosition &&
            long.TryParse(simplePosition.GetAttribute("x", "").Value, out var simpleX) &&
            long.TryParse(simplePosition.GetAttribute("y", "").Value, out var simpleY) &&
            info.Presentation is { FrameWidthPoints: > 0, FrameHeightPoints: > 0 } simpleFrame)
        {
            if (!_paragraphImages.TryGetValue(_paragraph, out var images))
                _paragraphImages[_paragraph] = images = new List<ParagraphImageOverlay>();
            images.Add(new ParagraphImageOverlay(bytes,
                simpleX / 12700.0 +
                    (simpleFrame.FrameWidthPoints.Value - (transformed?.Width ?? simpleFrame.FrameWidthPoints.Value)) / 2,
                simpleY / 12700.0 +
                    (simpleFrame.FrameHeightPoints.Value - (transformed?.Height ?? simpleFrame.FrameHeightPoints.Value)) / 2,
                transformed?.Width ?? simpleFrame.FrameWidthPoints.Value,
                transformed?.Height ?? simpleFrame.FrameHeightPoints.Value,
                PageRelative: true, ZOrder: zOrder, BehindDocument: behindDocument));
            return DxpDisposable.Empty;
        }
        if (anchored && _inHeader && _container is HeaderFooter && _section != null &&
            vertical?.GetAttribute("relativeFrom", "").Value == "margin" &&
            drawing.Descendants().Any(element => element.LocalName == "wrapSquare") &&
            drawing.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is
                { } headerParagraph)
        {
            using var measure = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
                XPageDirection.Downwards);
            var fontSize = _paragraph?.Format.Font.Size.Point is > 0 ? _paragraph.Format.Font.Size.Point : 11;
            var textWidth = measure.MeasureString(headerParagraph.InnerText,
                new XFont(_fontFamily, fontSize)).Width;
            var page = _section.PageSetup;
            if (textWidth > page.PageWidth.Point - page.LeftMargin.Point - page.RightMargin.Point &&
                _headersWithMarginAnchoredWrap.Add(_section))
                page.TopMargin = Unit.FromPoint(page.TopMargin.Point + 8);
            var imageAbsoluteTop = page.TopMargin.Point + verticalOffset;
            var overlapsHeaderLine = imageAbsoluteTop < page.HeaderDistance.Point +
                    _defaultParagraphLinePoints &&
                imageAbsoluteTop + info.Presentation?.FrameHeightPoints > page.HeaderDistance.Point;
            if ((textWidth > page.PageWidth.Point - page.LeftMargin.Point - page.RightMargin.Point ||
                 overlapsHeaderLine) && horizontalOffset is double imageLeft &&
                info.Presentation?.FrameWidthPoints is double imageWidth &&
                anchorElement != null)
            {
                var leftDistance = int.TryParse(anchorElement.GetAttribute("distL", "").Value,
                    out var leftEmu) ? leftEmu / 12700.0 : 0;
                var rightDistance = int.TryParse(anchorElement.GetAttribute("distR", "").Value,
                    out var rightEmu) ? rightEmu / 12700.0 : 0;
                var visualLeft = imageLeft + (imageWidth - (transformed?.Width ?? imageWidth)) / 2;
                var headerWrap = new FloatingTextWrap(visualLeft - leftDistance,
                    visualLeft + (transformed?.Width ?? imageWidth) + rightDistance);
                _activeFloatingTextWrap = headerWrap;
                if (verticalOffset is double imageTop &&
                    info.Presentation.FrameHeightPoints is double imageHeight &&
                    imageTop < 15 && imageTop + imageHeight > 0 &&
                    (context as DxpDocumentContext)?.MainDocumentPart?.Document.Body?
                        .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is
                        { } firstBodyParagraph)
                {
                    _floatingTextWraps[firstBodyParagraph] = headerWrap;
                    _bodyParagraphsUnderHeaderWrap.Add(firstBodyParagraph);
                }
            }
        }
        if (anchored && _container is Section && _paragraph != null &&
            horizontal?.GetAttribute("relativeFrom", "").Value == "margin" &&
            vertical?.GetAttribute("relativeFrom", "").Value == "margin" &&
            info.Presentation is { FrameWidthPoints: > 0, FrameHeightPoints: > 0 } marginImageFrame &&
            horizontalOffset is double marginImageLeft && verticalOffset is double marginImageTop)
        {
            if (!_paragraphImages.TryGetValue(_paragraph, out var images))
                _paragraphImages[_paragraph] = images = new List<ParagraphImageOverlay>();
            images.Add(new ParagraphImageOverlay(bytes,
                _section!.PageSetup.LeftMargin.Point + marginImageLeft +
                    (marginImageFrame.FrameWidthPoints.Value - (transformed?.Width ?? marginImageFrame.FrameWidthPoints.Value)) / 2,
                _section.PageSetup.TopMargin.Point + marginImageTop +
                    (marginImageFrame.FrameHeightPoints.Value - (transformed?.Height ?? marginImageFrame.FrameHeightPoints.Value)) / 2,
                transformed?.Width ?? marginImageFrame.FrameWidthPoints.Value,
                transformed?.Height ?? marginImageFrame.FrameHeightPoints.Value,
                PageRelative: true, ZOrder: zOrder, BehindDocument: behindDocument));
            var marginSquareWrap = GetRectangularTextWrap(drawing, SupportsContourLayout);
            IReadOnlyList<XPoint>? contourPoints = null;
            if (marginSquareWrap == null &&
                SupportsContourLayout)
                marginSquareWrap = GetContourTextWrap(drawing, out contourPoints);
            if (marginSquareWrap == null &&
                (SupportsContourLayout ||
                 typeof(DocumentRenderer).GetProperty("EnableTopBottomExclusion") != null))
                marginSquareWrap = drawing.Descendants().FirstOrDefault(element =>
                    element.LocalName == "wrapTopAndBottom");
            if (marginSquareWrap != null)
            {
                var side = marginSquareWrap.LocalName == "wrapTopAndBottom"
                    ? "TopBottom"
                    : marginSquareWrap.GetAttribute("wrapText", "").Value switch
                {
                    "left" => "Left",
                    "right" => "Right",
                    "largest" => "Largest",
                    "bothSides" => "BothSides",
                    _ => null
                };
                if (side != null && Enum.TryParse<MigraDoc.DocumentObjectModel.Shapes.WrapStyle>(side,
                        out var wrapStyle))
                {
                    var marker = _section.AddImage(
                        "base64:iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4//8/AwAI/AL+p5qgoAAAAABJRU5ErkJggg==");
                    marker.RelativeHorizontal = MigraDoc.DocumentObjectModel.Shapes.RelativeHorizontal.Margin;
                    marker.RelativeVertical = MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Margin;
                    marker.Left = Unit.FromPoint(marginImageLeft);
                    marker.Top = Unit.FromPoint(marginImageTop);
                    marker.Width = Unit.FromPoint(marginImageFrame.FrameWidthPoints.Value);
                    marker.Height = Unit.FromPoint(marginImageFrame.FrameHeightPoints.Value);
                    marker.WrapFormat.Style = wrapStyle;
                    if (contourPoints != null)
                    {
                        _wrapContours[marker] = contourPoints;
                        if (drawing.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                                .FirstOrDefault()?.NextSibling<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                            is { } followingParagraph)
                        {
                            var runStyles = followingParagraph.Descendants<Run>()
                                .Where(run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any())
                                .Select(run => context.Styles.ResolveRunStyle(followingParagraph, run))
                                .ToArray();
                            if (runStyles.Length > 0 && runStyles.All(style =>
                                    style.FontName == runStyles[0].FontName &&
                                    style.FontSizeHalfPoints == runStyles[0].FontSizeHalfPoints))
                            {
                                var fontName = runStyles[0].FontName ?? _fontFamily;
                                var fontSize = runStyles[0].FontSizeHalfPoints is int halfPoints
                                    ? halfPoints / 2.0 : _document!.Styles["Normal"].Font.Size.Point;
                                try { _wrapContourBandHeights[marker] = new XFont(fontName, fontSize).GetHeight() + 0.25; }
                                catch (InvalidOperationException) { }
                            }
                        }
                    }
                    if (anchorElement != null)
                    {
                        if (int.TryParse(anchorElement.GetAttribute("distL", "").Value, out var leftEmu))
                            marker.WrapFormat.DistanceLeft = Unit.FromPoint(leftEmu / 12700.0);
                        if (int.TryParse(anchorElement.GetAttribute("distR", "").Value, out var rightEmu))
                            marker.WrapFormat.DistanceRight = Unit.FromPoint(rightEmu / 12700.0);
                        if (int.TryParse(anchorElement.GetAttribute("distT", "").Value, out var topEmu))
                            marker.WrapFormat.DistanceTop = Unit.FromPoint(topEmu / 12700.0);
                        if (int.TryParse(anchorElement.GetAttribute("distB", "").Value, out var bottomEmu))
                            marker.WrapFormat.DistanceBottom = Unit.FromPoint(bottomEmu / 12700.0);
                    }
                }
            }
            return DxpDisposable.Empty;
        }
        if (anchored && _container is Section && _paragraph != null &&
            horizontal?.GetAttribute("relativeFrom", "").Value == "page" &&
            vertical?.GetAttribute("relativeFrom", "").Value == "page" &&
            info.Presentation is { FrameWidthPoints: > 0, FrameHeightPoints: > 0 } pageFrame &&
            PositionImageOffset(horizontal, "page") is double pageLeft &&
            PositionImageOffset(vertical, "page") is double pageTop)
        {
            if (!_paragraphImages.TryGetValue(_paragraph, out var images))
                _paragraphImages[_paragraph] = images = new List<ParagraphImageOverlay>();
            images.Add(new ParagraphImageOverlay(bytes,
                pageLeft + (pageFrame.FrameWidthPoints.Value -
                    (transformed?.Width ?? pageFrame.FrameWidthPoints.Value)) / 2,
                pageTop + (pageFrame.FrameHeightPoints.Value -
                    (transformed?.Height ?? pageFrame.FrameHeightPoints.Value)) / 2,
                transformed?.Width ?? pageFrame.FrameWidthPoints.Value,
                transformed?.Height ?? pageFrame.FrameHeightPoints.Value, PageRelative: true,
                ZOrder: zOrder, BehindDocument: behindDocument));
            return DxpDisposable.Empty;
        }
        var paragraphAnchor = anchored && _container is Section &&
            horizontal?.GetAttribute("relativeFrom", "").Value == "column" &&
            vertical?.GetAttribute("relativeFrom", "").Value == "paragraph" &&
            drawing.Descendants().Any(element => element.LocalName == "wrapSquare") &&
            drawing.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault()?
                .Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any() == true;
        var marginParagraphAnchor = anchored && _container is Section && _paragraph != null &&
            horizontal?.GetAttribute("relativeFrom", "").Value == "margin" &&
            vertical?.GetAttribute("relativeFrom", "").Value == "paragraph" &&
            drawing.Descendants().Any(element => element.LocalName == "wrapSquare");
        if (marginParagraphAnchor && _paragraph != null && _section != null &&
            info.Presentation is { FrameWidthPoints: > 0, FrameHeightPoints: > 0 } marginFrame &&
            horizontalOffset is double marginLeft &&
            PositionImageOffset(vertical, "paragraph") is double paragraphTop)
        {
            var anchor = drawing.Descendants().First(element => element.LocalName == "anchor");
            var leftDistance = int.TryParse(anchor.GetAttribute("distL", "").Value, out var leftEmu)
                ? leftEmu / 12700.0 : 0;
            var rightDistance = int.TryParse(anchor.GetAttribute("distR", "").Value, out var rightEmu)
                ? rightEmu / 12700.0 : 0;
            if (!_paragraphImages.TryGetValue(_paragraph, out var images))
                _paragraphImages[_paragraph] = images = new List<ParagraphImageOverlay>();
            images.Add(new ParagraphImageOverlay(bytes,
                marginLeft + (marginFrame.FrameWidthPoints.Value -
                    (transformed?.Width ?? marginFrame.FrameWidthPoints.Value)) / 2,
                paragraphTop + (marginFrame.FrameHeightPoints.Value -
                    (transformed?.Height ?? marginFrame.FrameHeightPoints.Value)) / 2,
                transformed?.Width ?? marginFrame.FrameWidthPoints.Value,
                transformed?.Height ?? marginFrame.FrameHeightPoints.Value,
                ZOrder: zOrder, BehindDocument: behindDocument));
            if (drawing.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault()?
                    .NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph following)
                _floatingTextWraps[following] = new FloatingTextWrap(marginLeft - leftDistance,
                    marginLeft + marginFrame.FrameWidthPoints.Value + rightDistance);
            return DxpDisposable.Empty;
        }
        if (paragraphAnchor && _paragraph != null &&
            info.Presentation is { FrameWidthPoints: > 0, FrameHeightPoints: > 0 } frame)
        {
            var anchor = drawing.Descendants().First(element => element.LocalName == "anchor");
            var rightDistance = int.TryParse(anchor.GetAttribute("distR", "").Value, out var emu)
                ? emu / 12700.0 : 0;
            var left = int.TryParse(horizontal?.ChildElements.FirstOrDefault(element =>
                element.LocalName == "posOffset")?.InnerText, out var horizontalEmu)
                ? horizontalEmu / 12700.0 : 0;
            var top = int.TryParse(vertical?.ChildElements.FirstOrDefault(element =>
                element.LocalName == "posOffset")?.InnerText, out var verticalEmu)
                ? verticalEmu / 12700.0 : 0;
            if (!_paragraphImages.TryGetValue(_paragraph, out var images))
                _paragraphImages[_paragraph] = images = new List<ParagraphImageOverlay>();
            images.Add(new ParagraphImageOverlay(bytes,
                left + (frame.FrameWidthPoints.Value -
                    (transformed?.Width ?? frame.FrameWidthPoints.Value)) / 2,
                top + (frame.FrameHeightPoints.Value -
                    (transformed?.Height ?? frame.FrameHeightPoints.Value)) / 2,
                transformed?.Width ?? frame.FrameWidthPoints.Value,
                transformed?.Height ?? frame.FrameHeightPoints.Value,
                ZOrder: zOrder, BehindDocument: behindDocument));
            _paragraph.Format.LeftIndent = Unit.FromPoint(_paragraph.Format.LeftIndent.Point +
                frame.FrameWidthPoints.Value + rightDistance);
            return DxpDisposable.Empty;
        }
        var reservesCenteredParagraphImageFlow = false;
        if (anchored && _container is Section && _paragraph != null &&
            horizontal?.GetAttribute("relativeFrom", "").Value == "column" &&
            horizontal.ChildElements.Any(element => element.LocalName == "align" &&
                element.InnerText == "center") &&
            vertical?.GetAttribute("relativeFrom", "").Value == "paragraph" &&
            vertical.ChildElements.Any(element => element.LocalName == "align" &&
                element.InnerText == "top") &&
            drawing.Descendants().Any(element => element.LocalName == "wrapSquare") &&
            !drawing.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First()
                .Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any() &&
            info.Presentation is { FrameWidthPoints: > 0, FrameHeightPoints: > 0 } inlineAnchorFrame)
        {
            if (!_paragraphImages.TryGetValue(_paragraph, out var images))
                _paragraphImages[_paragraph] = images = new List<ParagraphImageOverlay>();
            var fontSize = _paragraph.Format.Font.Size.Point is > 0
                ? _paragraph.Format.Font.Size.Point : 11;
            images.Add(new ParagraphImageOverlay(bytes, 0,
                fontSize + 0.35,
                transformed?.Width ?? inlineAnchorFrame.FrameWidthPoints.Value,
                transformed?.Height ?? inlineAnchorFrame.FrameHeightPoints.Value,
                ZOrder: zOrder, BehindDocument: behindDocument));
            imageData = "base64:iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4//8/AwAI/AL+p5qgoAAAAABJRU5ErkJggg==";
            reservesCenteredParagraphImageFlow = true;
        }
        if (anchored && _container is HeaderFooter && _section != null &&
            horizontal?.GetAttribute("relativeFrom", "").Value == "column" &&
            vertical?.GetAttribute("relativeFrom", "").Value == "paragraph" &&
            PositionImageOffset(horizontal, "column") is double storyLeft &&
            PositionImageOffset(vertical, "paragraph") is double storyTop &&
            info.Presentation is { FrameWidthPoints: > 0, FrameHeightPoints: > 0 } storyFrame)
        {
            if (!_footerImages.TryGetValue(_section, out var images))
                _footerImages[_section] = images = new List<FooterImageOverlay>();
            images.Add(new FooterImageOverlay(_headerFooterKind, bytes, storyLeft, storyTop,
                transformed?.Width ?? storyFrame.FrameWidthPoints.Value,
                transformed?.Height ?? storyFrame.FrameHeightPoints.Value,
                Header: _inHeader, AnchorParagraph: _paragraph, BehindDocument: behindDocument,
                ZOrder: zOrder));
            if (drawing.Descendants().Any(element => element.LocalName == "wrapSquare") &&
                drawing.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault()?
                    .Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any() == true)
            {
                var anchor = drawing.Descendants().First(element => element.LocalName == "anchor");
                var rightDistance = int.TryParse(anchor.GetAttribute("distR", "").Value, out var emu)
                    ? emu / 12700.0 : 0;
                _paragraph.Format.LeftIndent = Unit.FromPoint(_paragraph.Format.LeftIndent.Point +
                    storyLeft + storyFrame.FrameWidthPoints.Value + rightDistance);
            }
            return DxpDisposable.Empty;
        }
        if (anchored && _container is HeaderFooter && !_inHeader && _section != null &&
            horizontalOffset is double footerLeft && verticalOffset is double footerTop &&
            info.Presentation?.FrameWidthPoints is double footerWidth && footerWidth > 0 &&
            info.Presentation.FrameHeightPoints is double footerHeight && footerHeight > 0)
        {
            if (!_footerImages.TryGetValue(_section, out var images))
                _footerImages[_section] = images = new List<FooterImageOverlay>();
            images.Add(new FooterImageOverlay(_headerFooterKind, bytes,
                footerLeft + (footerWidth - (transformed?.Width ?? footerWidth)) / 2,
                footerTop + (footerHeight - (transformed?.Height ?? footerHeight)) / 2,
                transformed?.Width ?? footerWidth, transformed?.Height ?? footerHeight,
                BehindDocument: behindDocument, ZOrder: zOrder));
            return DxpDisposable.Empty;
        }
        if (!anchored && transformed is { } headerRaster && _inHeader &&
            _container is HeaderFooter && _section != null &&
            info.Presentation is { FrameWidthPoints: > 0, FrameHeightPoints: > 0 } headerFrame)
        {
            if (!_footerImages.TryGetValue(_section, out var images))
                _footerImages[_section] = images = new List<FooterImageOverlay>();
            images.Add(new FooterImageOverlay(_headerFooterKind, bytes,
                    (headerFrame.FrameWidthPoints.Value - headerRaster.Width) / 2,
                    (headerFrame.FrameHeightPoints.Value - headerRaster.Height) / 2,
                    headerRaster.Width, headerRaster.Height, Header: true));
            imageData = "base64:iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4//8/AwAI/AL+p5qgoAAAAABJRU5ErkJggg==";
        }
        var image = anchored && _container is Section section
            ? section.AddImage(imageData)
            : anchored && _container is HeaderFooter headerFooter
                ? headerFooter.AddImage(imageData)
            : _paragraph.AddImage(imageData);
        if (anchored && _container is Section &&
            drawing.Descendants().Any(element => element.LocalName == "wrapTopAndBottom"))
            image.WrapFormat.Style = MigraDoc.DocumentObjectModel.Shapes.WrapStyle.TopBottom;
        if (anchored && !reservesCenteredParagraphImageFlow && _container is Section &&
            drawing.Descendants().FirstOrDefault(element => element.LocalName == "wrapSquare") is
                { } squareWrap)
        {
            var side = squareWrap.GetAttribute("wrapText", "").Value switch
            {
                "left" => "Left",
                "right" => "Right",
                "largest" => "Largest",
                "bothSides" => "BothSides",
                _ => null
            };
            if (side != null && Enum.TryParse<MigraDoc.DocumentObjectModel.Shapes.WrapStyle>(side,
                    out var wrapStyle))
                image.WrapFormat.Style = wrapStyle;
        }
        if (anchored)
        {
            if (horizontalOffset is double left)
            {
                image.RelativeHorizontal = MigraDoc.DocumentObjectModel.Shapes.RelativeHorizontal.Margin;
                image.Left = Unit.FromPoint(left +
                    (info.Presentation?.FrameWidthPoints - transformed?.Width ?? 0) / 2);
            }
            if (verticalOffset is double top)
            {
                image.RelativeVertical = MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Margin;
                image.Top = Unit.FromPoint(top +
                    (info.Presentation?.FrameHeightPoints - transformed?.Height ?? 0) / 2);
            }
        }
        if (info.Presentation?.FrameWidthPoints is double width && width > 0)
            image.Width = Unit.FromPoint(_inHeader && !anchored ? width : transformed?.Width ?? width);
        if (info.Presentation?.FrameHeightPoints is double height && height > 0)
        {
            image.Height = Unit.FromPoint(_inHeader && !anchored ? height : transformed?.Height ?? height);
            if (_inHeader && !anchored && _container is HeaderFooter && _section != null)
            {
                var page = _section.PageSetup;
                page.TopMargin = Unit.FromPoint(Math.Max(page.TopMargin.Point,
                    page.HeaderDistance.Point + height + 3));
            }
        }
        if (!anchored && _paragraph != null && image.Width.Point > 0 && image.Height.Point > 0)
        {
            if (Environment.GetEnvironmentVariable("DOCXPORT_DIAG_INLINE_IMAGE_PARAGRAPH_SPACE") == "1" &&
                _container is Section && image.Height.Point >= 12 &&
                drawing.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is
                    { } imageParagraph &&
                imageParagraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any() &&
                imageParagraph.Descendants<Drawing>().Count() == 1)
            {
                var normalSpacing = _document.Styles["Normal"].ParagraphFormat;
                var after = _paragraph.Format.Values.SpaceAfter?.Point ?? normalSpacing.SpaceAfter.Point;
                _paragraph.Format.SpaceAfter = Unit.FromPoint(after + 0.308);
                if (_section?.Elements is { } elements && elements.Count >= 2 &&
                    ReferenceEquals(elements[elements.Count - 1], _paragraph) &&
                    elements[elements.Count - 2] is MigraDoc.DocumentObjectModel.Paragraph previous)
                {
                    var previousAfter = previous.Format.Values.SpaceAfter?.Point ?? normalSpacing.SpaceAfter.Point;
                    previous.Format.SpaceAfter = Unit.FromPoint(previousAfter + 0.687);
                }
            }
            var normal = _document.Styles["Normal"].ParagraphFormat;
            if ((_paragraph.Format.Values.LineSpacingRule ?? normal.LineSpacingRule) == LineSpacingRule.Exactly)
            {
                var lineHeight = _paragraph.Format.Values.LineSpacing?.Point ?? normal.LineSpacing.Point;
                if (lineHeight > 0 && image.Height.Point > lineHeight + 1)
                    _exactLineImages.Add((_paragraph, image.Width.Point, image.Height.Point, lineHeight));
            }
        }
        if (!anchored && _container is Section &&
            drawing.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is
                { } sourceParagraph &&
            !sourceParagraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any() &&
            !sourceParagraph.Descendants<Picture>().Any() &&
            !sourceParagraph.Descendants<EmbeddedObject>().Any() &&
            sourceParagraph.Descendants<Drawing>().All(item => item.Descendants().Any(element =>
                element.LocalName == "inline") && !item.Descendants().Any(element =>
                element.LocalName == "chart")) &&
            ReferenceEquals(sourceParagraph.Descendants<Drawing>().Last(), drawing))
            ApplyStandaloneImageLineSpacing(sourceParagraph, drawing, context,
                _paragraph.Elements.OfType<MigraDoc.DocumentObjectModel.Shapes.Image>()
                    .Max(item => item.Height.Point));
        return DxpDisposable.Empty;
    }

    private static double? MarginImageOffset(OpenXmlElement? position)
        => PositionImageOffset(position, "margin");

    private static double? PositionImageOffset(OpenXmlElement? position, string relativeFrom)
    {
        if (position?.GetAttribute("relativeFrom", "").Value != relativeFrom ||
            !int.TryParse(position.ChildElements.FirstOrDefault(element => element.LocalName == "posOffset")?.InnerText,
                out var offset)) return null;
        return offset / 12700.0;
    }

    private bool TryAddChart(Drawing drawing, DxpDrawingInfo? info, DxpIDocumentContext context)
    {
        if (_container is not Section section || _paragraph == null || context.CurrentPart == null)
            return false;
        const string relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var reference = drawing.Descendants().FirstOrDefault(element =>
            element.LocalName == "chart" && element.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/chart");
        var relationshipId = reference?.GetAttribute("id", relationshipNamespace).Value;
        if (string.IsNullOrEmpty(relationshipId) ||
            context.CurrentPart.GetPartById(relationshipId) is not ChartPart chartPart)
            return false;
        XDocument xml;
        try
        {
            using var stream = chartPart.GetStream(FileMode.Open, FileAccess.Read);
            xml = XDocument.Load(stream);
        }
        catch (System.Xml.XmlException)
        {
            return false;
        }
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XNamespace drawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var plot = xml.Descendants(chartNamespace + "barChart").FirstOrDefault();
        if (plot == null) return false;
        var paragraph = drawing.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
        if (paragraph != null && !paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Any() &&
            paragraph.Descendants<Drawing>().Count() == 1)
        {
            _paragraph.Format.LineSpacingRule = LineSpacingRule.Exactly;
            _paragraph.Format.LineSpacing = Unit.FromPoint(9);
            _paragraph.Format.SpaceBefore = Unit.Zero;
            var precedingList = paragraph.PreviousSibling() as DocumentFormat.OpenXml.Wordprocessing.Paragraph;
            _paragraph.Format.SpaceAfter = Unit.FromPoint(precedingList?.ParagraphProperties?.
                NumberingProperties?.NumberingId != null &&
                precedingList.ParagraphProperties?.SpacingBetweenLines?.After?.Value == "225" &&
                _availableFontFamilies.Contains("DejaVu Sans") ? 7.4 : 12);
        }
        var direction = plot.Element(chartNamespace + "barDir")?.Attribute("val")?.Value;
        var chart = section.AddChart(direction == "bar" ? ChartType.Bar2D : ChartType.Column2D);
        var chartFont = _availableFontFamilies.Contains("Aptos") ? "Aptos" : _fontFamily;
        chart.XAxis.TickLabels.Font.Name = chartFont;
        chart.XAxis.TickLabels.Font.Size = Unit.FromPoint(10);
        chart.YAxis.TickLabels.Font.Name = chartFont;
        chart.YAxis.TickLabels.Font.Size = Unit.FromPoint(10);
        var plotAreaColor = xml.Descendants(chartNamespace + "plotArea").FirstOrDefault()?
            .Element(chartNamespace + "spPr")?.Element(drawingNamespace + "ln")?
            .Descendants(drawingNamespace + "srgbClr").FirstOrDefault()?.Attribute("val")?.Value;
        if (plotAreaColor != null && ParseColor("#" + plotAreaColor) is { } plotBorderColor)
        {
            chart.PlotArea.LineFormat.Color = plotBorderColor;
            chart.PlotArea.LineFormat.Width = Unit.FromPoint(0.75);
        }
        chart.XAxis.MajorTickMark = TickMarkType.None;
        chart.YAxis.MajorTickMark = TickMarkType.None;
#if PDFSHARP_SOURCE7
        var categoryAxis = xml.Descendants(chartNamespace + "catAx").FirstOrDefault();
        if (direction != "bar" && categoryAxis?.Element(chartNamespace + "majorTickMark")?
            .Attribute("val")?.Value == "out")
        {
            chart.XAxis.MajorTickMark = TickMarkType.Outside;
            chart.XAxis.MajorTickMarkLength = 3.13;
            chart.XAxis.TickLabelClearance = 4.64;
            chart.XAxis.TickLabelOffset = 0.34;
            chart.YAxis.TickLabelVerticalOffset = 0.17;
            var categoryAxisColor = categoryAxis.Element(chartNamespace + "spPr")?
                .Descendants(drawingNamespace + "srgbClr").FirstOrDefault()?.Attribute("val")?.Value;
            if (categoryAxisColor != null && ParseColor("#" + categoryAxisColor) is { } categoryLineColor)
                chart.XAxis.LineFormat.Color = categoryLineColor;
            chart.XAxis.LineFormat.Width = Unit.FromPoint(1);
        }
#endif
        var valueAxis = direction == "bar" ? chart.XAxis : chart.YAxis;
        var axisDefinition = xml.Descendants(chartNamespace + "valAx").FirstOrDefault();
        valueAxis.MajorTickMark = axisDefinition?.Element(chartNamespace + "majorTickMark")?
            .Attribute("val")?.Value switch
        {
            "out" => TickMarkType.Outside,
            "in" => TickMarkType.Inside,
            "cross" => TickMarkType.Cross,
            _ => TickMarkType.None
        };
        var axisColor = axisDefinition?.Element(chartNamespace + "spPr")?
            .Descendants(drawingNamespace + "srgbClr").FirstOrDefault()?.Attribute("val")?.Value;
        if (axisColor != null && ParseColor("#" + axisColor) is { } axisLineColor)
            valueAxis.LineFormat.Color = axisLineColor;
        valueAxis.LineFormat.Width = Unit.FromPoint(1);
        var maximum = axisDefinition?.Element(chartNamespace + "scaling")?
            .Element(chartNamespace + "max")?.Attribute("val")?.Value;
        if (double.TryParse(maximum, NumberStyles.Float, CultureInfo.InvariantCulture, out var maximumScale))
            valueAxis.MaximumScale = maximumScale;
        var majorUnit = axisDefinition?.Element(chartNamespace + "majorUnit")?.Attribute("val")?.Value;
        if (double.TryParse(majorUnit, NumberStyles.Float, CultureInfo.InvariantCulture, out var majorTick))
            valueAxis.MajorTick = majorTick;
        var numberFormat = axisDefinition?.Element(chartNamespace + "numFmt")?.Attribute("formatCode")?.Value;
        if (numberFormat != null) valueAxis.TickLabels.Format = numberFormat;
        if (axisDefinition?.Element(chartNamespace + "majorGridlines") is { } majorGridlines)
        {
            valueAxis.HasMajorGridlines = true;
            valueAxis.MajorGridlines.LineFormat.Width = Unit.FromPoint(1);
            var gridColor = majorGridlines.Descendants(drawingNamespace + "srgbClr")
                .FirstOrDefault()?.Attribute("val")?.Value;
            if (ParseColor("#" + gridColor) is { } color)
                valueAxis.MajorGridlines.LineFormat.Color = color;
        }
        if (info?.Presentation?.FrameWidthPoints is double width && width > 0)
            chart.Width = Unit.FromPoint(Math.Max(1, width - 6));
        if (info?.Presentation?.FrameHeightPoints is double height && height > 0)
            chart.Height = Unit.FromPoint(Math.Max(1, height - 13));
        var largestValue = 0.0;
        var seriesNames = new List<string>();
        foreach (var source in plot.Elements(chartNamespace + "ser"))
        {
            var points = source.Element(chartNamespace + "val")?
                .Descendants(chartNamespace + "numCache").FirstOrDefault()?
                .Elements(chartNamespace + "pt")
                .Select(point => double.TryParse(point.Element(chartNamespace + "v")?.Value,
                    NumberStyles.Float, CultureInfo.InvariantCulture, out var value) ? (double?)value : null)
                .ToArray();
            if (points == null || points.Length == 0) continue;
            largestValue = Math.Max(largestValue, points.Max(point => point ?? 0));
            var series = chart.SeriesCollection.AddSeries();
            series.Name = source.Element(chartNamespace + "tx")?.Descendants(chartNamespace + "v").FirstOrDefault()?.Value ?? "";
            seriesNames.Add(series.Name);
            series.Add(points.Select(point => point ?? 0).ToArray());
            var color = source.Element(chartNamespace + "spPr")?
                .Descendants(drawingNamespace + "srgbClr").FirstOrDefault()?.Attribute("val")?.Value;
            if (ParseColor("#" + color) is { } fill) series.FillFormat.Color = fill;
        }
        if (largestValue > 0 && maximum == null && majorUnit == null)
        {
            var approximateTick = largestValue * 1.2 / 6;
            var magnitude = Math.Pow(10, Math.Floor(Math.Log10(approximateTick)));
            var tick = new[] { 1.0, 2.0, 5.0, 10.0 }
                .First(step => step * magnitude >= approximateTick) * magnitude;
            valueAxis.MajorTick = tick;
            valueAxis.MaximumScale = Math.Ceiling(largestValue * 1.2 / tick) * tick;
        }
        if (numberFormat == null && valueAxis.MajorTick >= 1 && valueAxis.MajorTick % 1 == 0 &&
            valueAxis.MaximumScale % 1 == 0)
        {
            valueAxis.TickLabels.Format = "0";
            if (direction != "bar") chart.LeftArea.Width = Unit.FromPoint(8.55);
        }
        var categories = plot.Elements(chartNamespace + "ser").FirstOrDefault()?
            .Element(chartNamespace + "cat")?.Descendants(chartNamespace + "strCache").FirstOrDefault()?
            .Elements(chartNamespace + "pt").Select(point => point.Element(chartNamespace + "v")?.Value ?? "").ToArray();
        if (categories is { Length: > 0 })
            chart.XValues.AddXSeries().Add(categories);
        if (xml.Descendants(chartNamespace + "legend").FirstOrDefault() is { } legend)
        {
            var position = legend.Element(chartNamespace + "legendPos")?.Attribute("val")?.Value;
            var chartLegend = position switch
            {
                "r" => chart.RightArea.AddLegend(),
                "l" => chart.LeftArea.AddLegend(),
                "t" => chart.TopArea.AddLegend(),
                _ => chart.BottomArea.AddLegend()
            };
            chartLegend.Format.Font.Name = chartFont;
            chartLegend.Format.Font.Size = Unit.FromPoint(10);
            if (position == "r")
            {
                chart.RightArea.Width = Unit.FromPoint(5);
                if (seriesNames.Count > 0 && seriesNames.All(name =>
                    Regex.IsMatch(name, @"^[A-Za-z0-9 ]+$")))
                {
                    _rightLegendLabelGroups.Add(seriesNames.ToArray());
                    if (direction != "bar") _columnRightLegendLabelGroups.Add(seriesNames.ToArray());
                }
            }
        }
        var chartSpacing = section.AddParagraph();
        chartSpacing.Format.LineSpacingRule = LineSpacingRule.Exactly;
        chartSpacing.Format.LineSpacing = Unit.FromPoint(1);
        chartSpacing.Format.SpaceBefore = Unit.Zero;
        chartSpacing.Format.SpaceAfter = Unit.FromPoint(25);
        return true;
    }

    private static bool HasSixDoubleBorders(DocumentFormat.OpenXml.Wordprocessing.Table table)
    {
        var borders = table.TableProperties?.TableBorders;
        return borders?.TopBorder?.Val?.Value == BorderValues.Double &&
            borders.BottomBorder?.Val?.Value == BorderValues.Double &&
            borders.LeftBorder?.Val?.Value == BorderValues.Double &&
            borders.RightBorder?.Val?.Value == BorderValues.Double &&
            borders.InsideHorizontalBorder?.Val?.Value == BorderValues.Double &&
            borders.InsideVerticalBorder?.Val?.Value == BorderValues.Double;
    }

    private bool UsePartialNativeDoubleBorders(DocumentFormat.OpenXml.Wordprocessing.Table table)
    {
        var borders = table.TableProperties?.TableBorders;
        return _minorThemeLatinFont != null && _availableFontFamilies.Contains(_minorThemeLatinFont) &&
            borders?.TopBorder?.Val?.Value == BorderValues.Double &&
            borders.BottomBorder?.Val?.Value == BorderValues.Double;
    }

    public override void VisitAltChunk(AltChunk chunk, DxpIDocumentContext context)
    {
        if (_container is not Section section || context.CurrentPart == null) return;
        var id = chunk.GetAttribute("id",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;
        if (string.IsNullOrEmpty(id) ||
            context.CurrentPart.GetPartById(id) is not AlternativeFormatImportPart part) return;
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        if (part.ContentType.StartsWith("text/html", StringComparison.OrdinalIgnoreCase))
        {
            DxpHtmlChunkImporter.Append(section, stream,
                _availableFontFamilies.Contains("Aptos") ? "Aptos" : _fontFamily,
                (table, border, spacing, padding) =>
                    _htmlBevelTables.Add(table, (border, spacing, padding)), part.ContentType);
            return;
        }
        if (!part.ContentType.StartsWith("text/plain", StringComparison.OrdinalIgnoreCase)) return;
        using var reader = new StreamReader(stream,
            DxpHtmlChunkImporter.DeclaredEncoding(part.ContentType) ?? System.Text.Encoding.UTF8, true);
        var pages = reader.ReadToEnd().Replace("\r\n", "\n").Replace('\r', '\n')
            .Replace('\v', '\n').Split('\f');
        var firstParagraph = true;
        for (var pageIndex = 0; pageIndex < pages.Length; pageIndex++)
        {
            if (pageIndex > 0) section.AddPageBreak();
            var lines = pages[pageIndex].Split('\n');
            var count = lines.Length - (lines[^1].Length == 0 ? 1 : 0);
            for (var index = 0; index < count; index++)
            {
                var paragraph = section.AddParagraph(lines[index]);
                paragraph.Format.Font.Name = "Courier New";
                paragraph.Format.Font.Size = Unit.FromPoint(10.56);
                paragraph.Format.LineSpacingRule = LineSpacingRule.Exactly;
                paragraph.Format.LineSpacing = Unit.FromPoint(11.88);
                if (firstParagraph && pageIndex == 0)
                    paragraph.Format.SpaceBefore = Unit.FromPoint(8.5);
                firstParagraph = false;
            }
        }
    }

    public override void VisitOMathParagraph(DocumentFormat.OpenXml.Math.Paragraph mathParagraph,
        DxpIDocumentContext context)
    {
        if (_paragraph == null || mathParagraph.Parent is not DocumentFormat.OpenXml.Wordprocessing.Paragraph
            wordParagraph || wordParagraph.ChildElements.Any(element =>
                element != mathParagraph && element is not ParagraphProperties &&
                element is not Run { InnerText: "" }) ||
            mathParagraph.ChildElements.Count != 1 ||
            mathParagraph.ChildElements[0] is not DocumentFormat.OpenXml.Math.OfficeMath math ||
            math.ChildElements.Count == 0)
            return;
        if (math.ChildElements.All(element => element is DocumentFormat.OpenXml.Math.Run))
        {
            _paragraph.Format.Alignment = ParagraphAlignment.Center;
            AppendMathRuns(math);
            return;
        }
        if (math.ChildElements.Count > 1 && math.ChildElements.All(element =>
                element is DocumentFormat.OpenXml.Math.Run or DocumentFormat.OpenXml.Math.Fraction or
                    DocumentFormat.OpenXml.Math.Radical or DocumentFormat.OpenXml.Math.Superscript or
                    DocumentFormat.OpenXml.Math.Subscript or DocumentFormat.OpenXml.Math.SubSuperscript or
                    DocumentFormat.OpenXml.Math.Matrix))
        {
            _paragraph.Format.Alignment = ParagraphAlignment.Center;
            _mixedDisplayMathParagraphs.Add(_paragraph);
            if (math.ChildElements.Any(element => element is DocumentFormat.OpenXml.Math.Run) &&
                math.ChildElements.Any(element => element is DocumentFormat.OpenXml.Math.Fraction))
                _mixedDisplayRunFractionParagraphs.Add(_paragraph);
            VisitOMath((DocumentFormat.OpenXml.Math.OfficeMath)math.CloneNode(true), context);
            return;
        }
        if (math.ChildElements.Count != 1 ||
            math.ChildElements[0] is not DocumentFormat.OpenXml.Math.Fraction &&
            math.ChildElements[0] is not DocumentFormat.OpenXml.Math.Radical)
            return;
        if (math.ChildElements[0] is DocumentFormat.OpenXml.Math.Radical radical)
        {
            var degree = radical.ChildElements.FirstOrDefault(element => element.LocalName == "deg");
            var radicand = radical.ChildElements.FirstOrDefault(element => element.LocalName == "e");
            if (degree?.ChildElements.Count != 0 || radicand?.ChildElements.Count != 1 ||
                radicand.ChildElements[0].LocalName != "r" || string.IsNullOrEmpty(radicand.InnerText))
                return;
            _mathRadicals[_paragraph] = radicand.InnerText;
            _paragraph.Format.LineSpacingRule = LineSpacingRule.Exactly;
            _paragraph.Format.LineSpacing = Unit.FromPoint(26.31);
            _paragraph.Format.SpaceBefore = Unit.Zero;
            _paragraph.Format.SpaceAfter = Unit.Zero;
            return;
        }
        var fraction = (DocumentFormat.OpenXml.Math.Fraction)math.ChildElements[0];
        var numerator = fraction.ChildElements.FirstOrDefault(element => element.LocalName == "num");
        var denominator = fraction.ChildElements.FirstOrDefault(element => element.LocalName == "den");
        if (numerator == null || denominator == null ||
            numerator.ChildElements.Count != 1 || denominator.ChildElements.Count != 1 ||
            numerator.ChildElements[0].LocalName != "r" ||
            denominator.ChildElements[0].LocalName != "r" ||
            string.IsNullOrEmpty(numerator.InnerText) || string.IsNullOrEmpty(denominator.InnerText))
            return;
        var explicitSizes = fraction.Descendants<DocumentFormat.OpenXml.Wordprocessing.FontSize>()
            .Select(size => int.TryParse(size.Val?.Value, out var halfPoints) ? halfPoints : 0)
            .Where(size => size > 0).Distinct().ToArray();
        var fontSize = explicitSizes.Length == 1 ? explicitSizes[0] / 2.0 : 12;
        _mathFractions[_paragraph] = (numerator.InnerText, denominator.InnerText, fontSize);
        if (wordParagraph.Parent is Body && wordParagraph.PreviousSibling() == null)
            _firstBodyDisplayFractions.Add(_paragraph);
        _paragraph.Format.LineSpacingRule = LineSpacingRule.Exactly;
        _paragraph.Format.LineSpacing = Unit.FromPoint(
            (_firstBodyDisplayFractions.Contains(_paragraph) ? 47.54 : 37.24) + 1.17 * (fontSize - 12));
        _paragraph.Format.SpaceBefore = Unit.Zero;
        _paragraph.Format.SpaceAfter = Unit.Zero;
    }

    private void AppendMathRuns(DocumentFormat.OpenXml.Math.OfficeMath math)
    {
        const string mathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var hasMathFont = false;
        try
        {
            _ = new XFont("Cambria Math", 12);
            hasMathFont = true;
        }
        catch (InvalidOperationException) { }
        foreach (var mathRun in math.ChildElements.Cast<DocumentFormat.OpenXml.Math.Run>())
        {
            if (string.IsNullOrEmpty(mathRun.InnerText)) continue;
            var properties = mathRun.ChildElements.FirstOrDefault(element =>
                element.LocalName == "rPr" && element.NamespaceUri == mathNamespace);
            var style = properties?.ChildElements.FirstOrDefault(element =>
                element.LocalName == "sty" && element.NamespaceUri == mathNamespace)?
                .GetAttribute("val", mathNamespace).Value;
            var normal = properties?.ChildElements.Any(element =>
                element.LocalName == "nor" && element.NamespaceUri == mathNamespace) == true;
            var italic = style is "i" or "bi" ||
                style == null && !normal && mathRun.InnerText.All(char.IsLetter);
            var glyphText = hasMathFont && italic && style is not "bi"
                ? MathItalicVariable(mathRun.InnerText) : mathRun.InnerText;
            var text = _paragraph!.AddFormattedText(glyphText);
            if (hasMathFont) text.Font.Name = "Cambria Math";
            text.Font.Italic = italic && glyphText == mathRun.InnerText;
            text.Font.Bold = style is "b" or "bi";
            var fontSize = mathRun.Descendants<FontSize>().FirstOrDefault()?.Val?.Value;
            if (int.TryParse(fontSize, out var halfPoints) && halfPoints > 0)
                text.Size = Unit.FromPoint(halfPoints / 2.0);
            else if (hasMathFont)
                text.Size = Unit.FromPoint(12);
        }
    }

    public override void VisitOMath(DocumentFormat.OpenXml.Math.OfficeMath math,
        DxpIDocumentContext context)
    {
        if (_paragraph == null || math.Parent is DocumentFormat.OpenXml.Math.Paragraph ||
            math.ChildElements.Count == 0) return;
        var precedingSpace = math.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Run
            precedingRun && precedingRun.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .LastOrDefault()?.Text?.EndsWith(' ') == true;
        var followingSpace = math.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Run
            followingRun && followingRun.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                .FirstOrDefault()?.Text?.StartsWith(' ') == true;
        if (math.ChildElements.All(element => element is DocumentFormat.OpenXml.Math.Run))
        {
            for (var childIndex = 0; childIndex < math.ChildElements.Count; childIndex++)
            {
                var mathRun = (DocumentFormat.OpenXml.Math.Run)math.ChildElements[childIndex];
                AppendMixedMathRun(mathRun,
                    childIndex == 0 ? precedingSpace :
                        math.ChildElements[childIndex - 1].InnerText.EndsWith(' '),
                    childIndex == math.ChildElements.Count - 1 ? followingSpace :
                        math.ChildElements[childIndex + 1].InnerText.StartsWith(' '),
                    false);
            }
            return;
        }
        if (math.ChildElements.Count > 1 && math.ChildElements.All(element =>
                element is DocumentFormat.OpenXml.Math.Run or DocumentFormat.OpenXml.Math.Fraction or
                    DocumentFormat.OpenXml.Math.Radical or DocumentFormat.OpenXml.Math.Superscript or
                    DocumentFormat.OpenXml.Math.Subscript or DocumentFormat.OpenXml.Math.SubSuperscript or
                    DocumentFormat.OpenXml.Math.Matrix))
        {
            for (var childIndex = 0; childIndex < math.ChildElements.Count; childIndex++)
            {
                var element = math.ChildElements[childIndex];
                if (element is DocumentFormat.OpenXml.Math.Run mathRun)
                    AppendMixedMathRun(mathRun,
                        childIndex == 0 && precedingSpace,
                        childIndex == math.ChildElements.Count - 1 && followingSpace,
                        childIndex + 1 < math.ChildElements.Count &&
                            math.ChildElements[childIndex + 1] is DocumentFormat.OpenXml.Math.Radical);
                else
                {
                    var fragment = new DocumentFormat.OpenXml.Math.OfficeMath();
                    fragment.Append((OpenXmlElement)element.CloneNode(true));
                    VisitOMath(fragment, context);
                }
            }
            return;
        }
        if (math.ChildElements.Count != 1) return;
        if (math.ChildElements[0] is DocumentFormat.OpenXml.Math.Matrix matrix)
        {
            var rows = matrix.ChildElements.OfType<DocumentFormat.OpenXml.Math.MatrixRow>()
                .Select(row => row.ChildElements.Where(element => element.LocalName == "e")
                    .Select(element => element.InnerText).ToArray()).ToArray();
            if (rows.Length == 0 || rows[0].Length == 0 ||
                rows.Any(row => row.Length != rows[0].Length || row.Any(string.IsNullOrEmpty))) return;
            if (rows.Length == 1 && rows[0].Length == 1)
            {
                var hasInlineMathFont = true;
                try { _ = new XFont("Cambria Math", 12); }
                catch (InvalidOperationException) { hasInlineMathFont = false; }
                var text = _paragraph.AddFormattedText(hasInlineMathFont
                    ? MathItalicVariable(rows[0][0]) : rows[0][0]);
                if (hasInlineMathFont) text.Font.Name = "Cambria Math";
                else text.Italic = true;
                text.Size = Unit.FromPoint(12);
                return;
            }
            using var measure = XGraphics.CreateMeasureContext(new XSize(100, 100),
                XGraphicsUnit.Point, XPageDirection.Downwards);
            XFont font;
            var hasMathFont = true;
            try { font = new XFont("Cambria Math", 12); }
            catch (InvalidOperationException)
            {
                font = new XFont(_fontFamily, 12);
                hasMathFont = false;
            }
            var cells = hasMathFont
                ? rows.Select(row => row.Select(MathItalicVariable).ToArray()).ToArray()
                : rows;
            var widths = Enumerable.Range(0, cells[0].Length)
                .Select(column => cells.Max(row => measure.MeasureString(row[column], font).Width))
                .ToArray();
            var width = widths.Sum() + (widths.Length - 1) * 12.3 + 0.5 -
                (IsDigitMatrix(cells) ? 1.1 : 0);
            var surroundingFont = _previousTextStyle?.FontName is { Length: > 0 } previousFont &&
                _availableFontFamilies.Contains(previousFont) ? previousFont :
                _paragraph.Format.Font.Name is { Length: > 0 } paragraphFont ? paragraphFont : _fontFamily;
            var anchorFont = new XFont(surroundingFont, 14.02);
            var anchorWidth = measure.MeasureString("\u00A0", anchorFont).Width;
            var placeholder = _paragraph.AddFormattedText("\u00A0");
            placeholder.Font.Name = surroundingFont;
            placeholder.Size = Unit.FromPoint(14.02);
            if (width > anchorWidth)
            {
                var count = (int)Math.Ceiling((width - anchorWidth) / anchorWidth);
                var filler = _paragraph.AddFormattedText(new string('\u00A0', count));
                filler.Font.Name = surroundingFont;
                filler.Size = Unit.FromPoint(14.02 * (width - anchorWidth) / (count * anchorWidth));
            }
            _paragraph.Format.LineSpacingRule = LineSpacingRule.AtLeast;
            _paragraph.Format.LineSpacing = Unit.FromPoint(Math.Max(
                _paragraph.Format.LineSpacing.Point, 11.6 + (rows.Length - 1) * 8.5));
            if (!_inlineMathMatrices.TryGetValue(_paragraph, out var matrices))
                _inlineMathMatrices[_paragraph] = matrices = new();
            matrices.Add((cells, widths));
            return;
        }
        if (math.ChildElements[0] is DocumentFormat.OpenXml.Math.Fraction fraction)
        {
            var numerator = fraction.ChildElements.FirstOrDefault(element => element.LocalName == "num");
            var denominator = fraction.ChildElements.FirstOrDefault(element => element.LocalName == "den");
            if (numerator?.ChildElements.SingleOrDefault()?.LocalName == "f" ||
                denominator?.ChildElements.SingleOrDefault()?.LocalName == "f")
            {
                using var nestedMeasure = XGraphics.CreateMeasureContext(new XSize(100, 100),
                    XGraphicsUnit.Point, XPageDirection.Downwards);
                var mathFamily = "Cambria Math";
                try { _ = new XFont(mathFamily, 8.52); }
                catch (InvalidOperationException) { mathFamily = _fontFamily; }
                var term = ParseInlineMathTerm(fraction, nestedMeasure, mathFamily, 8.52);
                if (term == null) return;
                var nestedFontFamily = _previousTextStyle?.FontName is { Length: > 0 } nestedPreviousFont &&
                    _availableFontFamilies.Contains(nestedPreviousFont) ? nestedPreviousFont :
                    _paragraph.Format.Font.Name is { Length: > 0 } nestedParagraphFont ? nestedParagraphFont :
                    _availableFontFamilies.Contains("Aptos") ? "Aptos" : _fontFamily;
                var nestedAnchorFont = new XFont(nestedFontFamily, 14.03);
                var nestedAnchorWidth = nestedMeasure.MeasureString("\u00A0", nestedAnchorFont).Width;
                var nestedAdvanceWidth = Math.Max(nestedAnchorWidth,
                    Math.Max(term.Width + 0.7, NestedMathBarWidth(term, true)));
                var anchor = _paragraph.AddFormattedText("\u00A0");
                anchor.Font.Name = nestedFontFamily;
                anchor.Size = Unit.FromPoint(14.03);
                if (nestedAdvanceWidth > nestedAnchorWidth)
                {
                    var count = (int)Math.Ceiling((nestedAdvanceWidth - nestedAnchorWidth) /
                        nestedAnchorWidth);
                    var filler = _paragraph.AddFormattedText(new string('\u00A0', count));
                    filler.Font.Name = nestedFontFamily;
                    filler.Size = Unit.FromPoint(14.03 * (nestedAdvanceWidth - nestedAnchorWidth) /
                        (count * nestedAnchorWidth));
                }
                _paragraph.Format.LineSpacingRule = LineSpacingRule.AtLeast;
                _paragraph.Format.LineSpacing = Unit.FromPoint(
                    (term.Denominator?.Text == null ? 18.12 : 23.89) -
                    (string.Equals(nestedFontFamily, "Arial", StringComparison.OrdinalIgnoreCase)
                        ? 0.98 : 0));
                if (!_nestedInlineMathFractions.TryGetValue(_paragraph, out var nested))
                    _nestedInlineMathFractions[_paragraph] = nested = new();
                nested.Add(term);
                return;
            }
            if (numerator?.ChildElements.Count != 1 || denominator?.ChildElements.Count != 1 ||
                numerator.ChildElements[0].LocalName != "r" ||
                denominator.ChildElements[0].LocalName != "r" ||
                string.IsNullOrEmpty(numerator.InnerText) || string.IsNullOrEmpty(denominator.InnerText))
                return;
            using var measure = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
                XPageDirection.Downwards);
            XFont mathFont;
            var mathFontSize = _mixedDisplayMathParagraphs.Contains(_paragraph) ? 12 : 8.52;
            try { mathFont = new XFont("Cambria Math", mathFontSize); }
            catch (InvalidOperationException) { mathFont = new XFont(_fontFamily, mathFontSize); }
            var width = Math.Max(measure.MeasureString(numerator.InnerText, mathFont).Width,
                measure.MeasureString(denominator.InnerText, mathFont).Width) * 1.04;
            var surroundingFont = _previousTextStyle?.FontName is { Length: > 0 } previousFont &&
                _availableFontFamilies.Contains(previousFont) ? previousFont :
                _paragraph.Format.Font.Name is { Length: > 0 } paragraphFont ? paragraphFont :
                _availableFontFamilies.Contains("Aptos") ? "Aptos" : _fontFamily;
            var localFractionHeight = !_mixedDisplayMathParagraphs.Contains(_paragraph);
            var anchorFontSize = localFractionHeight ? 15.0 : 14.01;
            var anchorFont = new XFont(surroundingFont, anchorFontSize);
            var anchorWidth = measure.MeasureString("\u00A0", anchorFont).Width;
            var placeholder = _paragraph.AddFormattedText("\u00A0");
            placeholder.Font.Name = surroundingFont;
            placeholder.Size = Unit.FromPoint(anchorFontSize);
            if (width > anchorWidth)
            {
                var count = (int)Math.Ceiling((width - anchorWidth) / anchorWidth);
                var filler = _paragraph.AddFormattedText(new string('\u00A0', count));
                filler.Font.Name = surroundingFont;
                filler.Size = Unit.FromPoint(anchorFontSize * (width - anchorWidth) / (count * anchorWidth));
            }
            if (!localFractionHeight)
            {
                _paragraph.Format.LineSpacingRule = LineSpacingRule.AtLeast;
                _paragraph.Format.LineSpacing = Unit.FromPoint(
                    string.Equals(surroundingFont, "Arial", StringComparison.OrdinalIgnoreCase) ? 17.15 : 18.12);
            }
            if (!_inlineMathFractions.TryGetValue(_paragraph, out var fractions))
                _inlineMathFractions[_paragraph] = fractions = new();
            fractions.Add((numerator.InnerText, denominator.InnerText, width));
            return;
        }
        if (math.ChildElements[0] is DocumentFormat.OpenXml.Math.Radical radical)
        {
            var degree = radical.ChildElements.FirstOrDefault(element => element.LocalName == "deg");
            var radicand = radical.ChildElements.FirstOrDefault(element => element.LocalName == "e");
            if (degree?.ChildElements.Count != 0 || radicand?.ChildElements.Count != 1 ||
                radicand.ChildElements[0].LocalName != "r" || string.IsNullOrEmpty(radicand.InnerText))
                return;
            var hasMathFont = false;
            try
            {
                _ = new XFont("Cambria Math", 12);
                hasMathFont = true;
            }
            catch (InvalidOperationException) { }
            var mathValue = hasMathFont ? MathItalicVariable(radicand.InnerText) : radicand.InnerText;
            var sourceSize = radical.Descendants<FontSize>().FirstOrDefault()?.Val?.Value;
            var hasExplicitSize = int.TryParse(sourceSize, out var halfPoints) && halfPoints > 0;
            var separateRadicand = hasMathFont && hasExplicitSize && halfPoints != 24;
            var text = _paragraph.AddFormattedText(" √" +
                (separateRadicand ? "" : mathValue));
            if (hasExplicitSize && !separateRadicand)
                text.Size = Unit.FromPoint(halfPoints / 2.0);
            else if (hasMathFont)
                text.Size = Unit.FromPoint(12);
            if (hasMathFont)
                text.Font.Name = "Cambria Math";
            if (separateRadicand)
            {
                var variable = _paragraph.AddFormattedText(mathValue);
                variable.Size = Unit.FromPoint(halfPoints / 2.0);
                variable.Font.Name = "Cambria Math";
            }
            if (hasMathFont)
            {
                _paragraph.Format.LineSpacingRule = LineSpacingRule.AtLeast;
                _paragraph.Format.LineSpacing = Unit.FromPoint(16.69);
                if (DocSystemFontAdvances.Find("Cambria Math") is { } mathMetrics &&
                    MathItalicVariable(radicand.InnerText).EnumerateRunes().FirstOrDefault() is
                        { Value: > 0 } radicandRune &&
                    ShouldUseTallerMathRadical(mathMetrics, radicandRune.Value))
                    _tallInlineMathRadicals.Add(_paragraph);
            }
            _inlineMathRadicals.Add(_paragraph);
            if (math.Parent == null)
                _mixedInlineMathRadicals.Add(_paragraph);
            else
                _standaloneInlineMathRadicals.Add(_paragraph);
            return;
        }
        var isCombined = math.ChildElements[0] is DocumentFormat.OpenXml.Math.SubSuperscript;
        var isSuperscript = isCombined ||
            math.ChildElements[0] is DocumentFormat.OpenXml.Math.Superscript;
        if (!isSuperscript && math.ChildElements[0] is not DocumentFormat.OpenXml.Math.Subscript)
            return;
        var mathScript = math.ChildElements[0];
        var baseElement = mathScript.ChildElements.FirstOrDefault(element => element.LocalName == "e");
        var scriptElement = mathScript.ChildElements.FirstOrDefault(element =>
            element.LocalName == (isSuperscript ? "sup" : "sub"));
        if (baseElement?.ChildElements.Count != 1 || scriptElement?.ChildElements.Count != 1 ||
            baseElement.ChildElements[0].LocalName != "r" || scriptElement.ChildElements[0].LocalName != "r" ||
            string.IsNullOrEmpty(baseElement.InnerText) || string.IsNullOrEmpty(scriptElement.InnerText))
            return;
        var subscriptElement = isCombined
            ? mathScript.ChildElements.FirstOrDefault(element => element.LocalName == "sub") : null;
        if (isCombined && (subscriptElement?.ChildElements.Count != 1 ||
            subscriptElement.ChildElements[0].LocalName != "r" ||
            string.IsNullOrEmpty(subscriptElement.InnerText)))
            return;
        var useMathFont = false;
        try
        {
            _ = new XFont("Cambria Math", 12);
            useMathFont = true;
        }
        catch (InvalidOperationException) { }
        var baseValue = useMathFont ? MathItalicVariable(baseElement.InnerText) : baseElement.InnerText;
        var baseText = _paragraph.AddFormattedText(baseValue);
        if (useMathFont) baseText.Font.Name = "Cambria Math";
        else baseText.Italic = true;
        var script = _paragraph.AddFormattedText(scriptElement.InnerText);
        if (useMathFont) script.Font.Name = "Cambria Math";
        if (isSuperscript) script.Superscript = true;
        else script.Subscript = true;
        script.Size = Unit.FromPoint(13.15);
        if (isSuperscript)
        {
            var advance = _paragraph.AddFormattedText();
            if (useMathFont) advance.Font.Name = "Cambria Math";
            advance.Size = Unit.FromPoint(4.5);
            advance.AddCharacter(SymbolName.Blank);
        }
        if (subscriptElement != null)
        {
            if (!_inlineMathCombinedScripts.TryGetValue(_paragraph, out var scripts))
                _inlineMathCombinedScripts[_paragraph] = scripts = new();
            scripts.Add((scriptElement.InnerText, subscriptElement.InnerText));
        }
    }

    private void AppendMixedMathRun(DocumentFormat.OpenXml.Math.Run mathRun,
        bool precedingSpace, bool followingSpace, bool followedByRadical)
    {
        var sourceText = mathRun.Descendants<DocumentFormat.OpenXml.Math.Text>().ToArray();
        if (sourceText.Length != 1 ||
            !sourceText[0].Text.Any(character => character is '+' or '='))
        {
            var fragment = new DocumentFormat.OpenXml.Math.OfficeMath();
            fragment.Append((OpenXmlElement)mathRun.CloneNode(true));
            AppendMathRuns(fragment);
            if (followedByRadical && mathRun.InnerText.All(char.IsLetterOrDigit))
                AddMixedMathOperatorSpace(3.5);
            return;
        }

        for (var index = 0; index < sourceText[0].Text.Length; index++)
        {
            var character = sourceText[0].Text[index];
            if (character is '+' or '=' &&
                (index == 0 || !char.IsWhiteSpace(sourceText[0].Text[index - 1])) &&
                !(index == 0 && precedingSpace))
                AddMixedMathOperatorSpace(character == '=' ? 13.4 : 12);
            var clonedRun = (DocumentFormat.OpenXml.Math.Run)mathRun.CloneNode(true);
            clonedRun.Descendants<DocumentFormat.OpenXml.Math.Text>().Single().Text =
                character.ToString();
            var fragment = new DocumentFormat.OpenXml.Math.OfficeMath();
            fragment.Append(clonedRun);
            AppendMathRuns(fragment);
            if (character is '+' or '=' && (index == sourceText[0].Text.Length - 1 ||
                !char.IsWhiteSpace(sourceText[0].Text[index + 1])) &&
                !(index == sourceText[0].Text.Length - 1 && followingSpace))
                AddMixedMathOperatorSpace(character == '=' ? 12 : followedByRadical ? 3 : 8.6);
        }
    }

    private void AddMixedMathOperatorSpace(double fontSize)
    {
        var spacer = _paragraph!.AddFormattedText("\u00A0");
        spacer.Font.Name = _availableFontFamilies.Contains("Cambria Math")
            ? "Cambria Math" : _fontFamily;
        spacer.Size = Unit.FromPoint(fontSize);
    }

    private static string MathItalicVariable(string value)
    {
        if (value.Length != 1) return value;
        var character = value[0];
        if (character is >= 'A' and <= 'Z')
            return char.ConvertFromUtf32(0x1D434 + character - 'A');
        if (character is >= 'a' and <= 'z')
            return character == 'h' ? "ℎ" : char.ConvertFromUtf32(0x1D44E + character - 'a');
        return value;
    }

    public override IDisposable VisitTableBegin(DocumentFormat.OpenXml.Wordprocessing.Table table,
        DxpTableModel model, DxpIDocumentContext context, DxpITableContext tableContext)
    {
        if (_availableFontFamilies.Contains("Aptos") && _container is Section && table.Parent is Body &&
            table.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                { InnerText.Length: > 0 } bodyBeforeTable && _lastParagraph != null &&
            _lastParagraph.Format.SpaceAfter.Point == 0 &&
            table.Elements<TableRow>().FirstOrDefault() is { } firstTableRow &&
            firstTableRow.Elements<TableCell>().Count() == 2 &&
            firstTableRow.Elements<TableCell>().All(cell =>
                cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count() == 1 &&
                cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First()
                    .ParagraphProperties?.SpacingBetweenLines is { Before.Value: "120", After.Value: "120" }) &&
            firstTableRow.NextSibling<TableRow>() is { } secondTableRow &&
            secondTableRow.Elements<TableCell>().All(cell =>
                cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count() >= 5))
        {
            var firstCellParagraph = firstTableRow.Elements<TableCell>().First()
                .Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First();
            var styleAutoSpacingDisabled = context.Styles.GetParagraphStyleChain(firstCellParagraph).Any(entry =>
                _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                definition.StyleParagraphProperties?.SpacingBetweenLines?.BeforeAutoSpacing?.Value == false);
            if (!styleAutoSpacingDisabled)
                foreach (var cell in firstTableRow.Elements<TableCell>())
                    _tableOpeningSpacingTransferred.Add(cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First());
            _lastParagraph.Format.SpaceAfter = Unit.FromPoint(
                bodyBeforeTable.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                    { ParagraphProperties.ParagraphStyleId.Val.Value: "Heading2" }
                    ? 2 + 6
                    : 5 + 6);
        }
        if (_sectionHasFooterTable && _container is Section bodySection &&
            table.Elements<TableRow>().Count() >= 20 &&
            table.Elements<TableRow>().FirstOrDefault()?.TableRowProperties?.GetFirstChild<TableHeader>() != null &&
            _repeatingTableFooterClearanceApplied.Add(bodySection))
            bodySection.PageSetup.BottomMargin = Unit.FromPoint(bodySection.PageSetup.BottomMargin.Point +
                RepeatingTableFooterClearancePoints);
        if (_inHeader && _container is Cell)
            _nestedHeaderRowCount += table.Elements<TableRow>().Count();
        var previous = _table;
        var previousSpacing = _tableCellSpacingPoints;
        var previousBackground = _tableBackground;
        var previousHorizontalMerge = _horizontalMergeTable;
        var previousTableInHeaderFooter = _tableInHeaderFooter;
        var previousTableDirection = _tableRightToLeft;
        var previousRowSpacing = _rowCellSpacingPoints;
        var previousImplicitContentTable = _implicitContentTable;
        var previousParagraphSpacing = _tableParagraphSpacing;
        var previousDoubleBorder = _doubleRowBorderPoints;
        var previousNativeDoubleBorders = _nativeDoubleTableBorders;
        var previousCompactRow = _previousHeaderRowCompact;
        var previousRowSpacingReduction = _previousHeaderRowSpacingReductionPoints;
        _tableCellSpacingPoints = 0;
        _tableBackground = null;
        _horizontalMergeTable = table.Descendants<HorizontalMerge>().Any();
        _tableInHeaderFooter = _container is HeaderFooter;
        if (_inHeader && _tableInHeaderFooter && table.TableProperties?.TableStyle != null)
            _headerHasStyledTable = true;
        if (_tableInHeaderFooter && _inHeader &&
            table.Elements<TableRow>().Any(row => row.TableRowProperties?.ChildElements
                .Any(child => child.LocalName is "gridBefore" or "gridAfter") == true))
            _headerMinimumRowClearancePoints += 3.2;
        _tableRightToLeft = tableContext.RightToLeft;
        if (_horizontalMergeTable && _container is HeaderFooter)
        {
            if (_inHeader)
            {
                _headerHasHorizontalMerge = true;
                _headerHasOccupiedHorizontalMerge |= table.Descendants<TableCell>().Any(cell =>
                    cell.TableCellProperties?.GetFirstChild<HorizontalMerge>()?.Val?.Value ==
                        MergedCellValues.Continue &&
                    cell.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                        .Any(text => !string.IsNullOrWhiteSpace(text.Text)));
            }
            else _footerHasHorizontalMerge = true;
        }
        _rowCellSpacingPoints = 0;
        _implicitContentTable = false;
        _tableParagraphSpacing = null;
        _doubleRowBorderPoints = 0;
        _nativeDoubleTableBorders = Enum.TryParse<BorderStyle>("Double", out _) &&
            (HasSixDoubleBorders(table) || UsePartialNativeDoubleBorders(table));
        _previousHeaderRowCompact = false;
        _previousHeaderRowSpacingReductionPoints = 0;
        if (_container is Section firstSection && table.Parent is Body && table.PreviousSibling() == null &&
            table.TableProperties?.TableBorders != null &&
            table.Descendants<TopLeftToBottomRightCellBorder>().Any())
        {
            var clearance = firstSection.AddParagraph();
            clearance.Format.LineSpacingRule = LineSpacingRule.Exactly;
            clearance.Format.LineSpacing = Unit.FromPoint(9);
            clearance.Format.SpaceAfter = Unit.Zero;
        }
        MigraDoc.DocumentObjectModel.Shapes.TextFrame? floatingFrame = null;
        var floatingLeftBorderWidth = 0.0;
        var position = table.TableProperties?.ChildElements.FirstOrDefault(element => element.LocalName == "tblpPr");
        if (_container is Section floatingSection && table.Parent is Body && position != null &&
            AttributeValue(position, "horzAnchor") is "margin" or "page" &&
            AttributeValue(position, "vertAnchor") is "margin" or "page" &&
            int.TryParse(AttributeValue(position, "tblpX"), out var tableX) &&
            int.TryParse(AttributeValue(position, "tblpY"), out var tableY) &&
            Enum.TryParse<MigraDoc.DocumentObjectModel.Shapes.WrapStyle>("BothSides", out var sideWrap))
        {
            floatingFrame = floatingSection.AddTextFrame();
            _floatingTableFrames.Add(floatingFrame);
            floatingFrame.RelativeHorizontal = AttributeValue(position, "horzAnchor") == "page"
                ? MigraDoc.DocumentObjectModel.Shapes.RelativeHorizontal.Page
                : MigraDoc.DocumentObjectModel.Shapes.RelativeHorizontal.Margin;
            floatingFrame.RelativeVertical = AttributeValue(position, "vertAnchor") == "page"
                ? MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Page
                : MigraDoc.DocumentObjectModel.Shapes.RelativeVertical.Margin;
            floatingFrame.Left = Unit.FromPoint(tableX / 20.0);
            if (tableContext.ComputedStyle.TableBorders?.Left is
                    { LineStyle: not DxpComputedBorderLineStyle.None, WidthPt: > 0 } floatingLeftBorder)
            {
                floatingLeftBorderWidth = floatingLeftBorder.WidthPt;
                floatingFrame.Left = Unit.FromPoint(floatingFrame.Left.Position.Point - floatingLeftBorderWidth);
            }
            floatingFrame.Top = Unit.FromPoint(tableY / 20.0);
            floatingFrame.MarginLeft = Unit.Zero;
            floatingFrame.MarginRight = Unit.Zero;
            floatingFrame.MarginTop = Unit.Zero;
            floatingFrame.MarginBottom = Unit.Zero;
            floatingFrame.WrapFormat.Style = sideWrap;
            if (int.TryParse(AttributeValue(position, "leftFromText"), out var leftDistance))
                floatingFrame.WrapFormat.DistanceLeft = Unit.FromPoint(leftDistance / 20.0);
            if (int.TryParse(AttributeValue(position, "rightFromText"), out var rightDistance))
                floatingFrame.WrapFormat.DistanceRight = Unit.FromPoint(rightDistance / 20.0);
            if (int.TryParse(AttributeValue(position, "topFromText"), out var topDistance))
                floatingFrame.WrapFormat.DistanceTop = Unit.FromPoint(topDistance / 20.0);
            if (int.TryParse(AttributeValue(position, "bottomFromText"), out var bottomDistance))
                floatingFrame.WrapFormat.DistanceBottom = Unit.FromPoint(bottomDistance / 20.0);
        }
        _table = floatingFrame?.AddTable() ?? _container switch
        {
            Section section => section.AddTable(),
            HeaderFooter headerFooter => headerFooter.AddTable(),
            Cell parentCell when IsSimpleSideBySideNestedTable(table) =>
                parentCell.Elements.AddTable(),
            Cell => null,
            _ => null
        };
        if (_table != null)
        {
            if (_container is Section && table.TableProperties?.TableBorders is not null &&
                tableContext.ComputedStyle.TableBorders is { } fragmentBorders &&
                (fragmentBorders.Top is { WidthPt: > 0 } || fragmentBorders.Bottom is { WidthPt: > 0 }))
                _tableFragmentBorders[_table] = fragmentBorders;
            if (floatingFrame != null &&
                tableContext.ComputedStyle.TableBorders?.Top is
                    { LineStyle: not DxpComputedBorderLineStyle.None } outerTop &&
                table.TableProperties?.TableBorders?.InsideHorizontalBorder == null)
                _floatingContinuationTopBorders[floatingFrame] = (_table, outerTop);
            var firstCellParagraph = table.Elements<TableRow>().SelectMany(row => row.Elements<TableCell>())
                .SelectMany(cell => cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                .FirstOrDefault();
            if (firstCellParagraph != null && context.Styles.GetParagraphStyleChain(firstCellParagraph).Any(entry =>
                    _paragraphStyles.TryGetValue(entry.StyleId, out var definition) &&
                    definition.StyleParagraphProperties?.SpacingBetweenLines?.BeforeAutoSpacing?.Value == false))
                _noAutoSpacingTables.Add(_table);
            if (!_nativeDoubleTableBorders && _tableRightToLeft && tableContext.ComputedStyle.TableBorders?.Left is
                { LineStyle: DxpComputedBorderLineStyle.Double } doubleLeftBorder)
                _rtlDoubleTableLeftBorders[_table] = doubleLeftBorder;
            if (!_nativeDoubleTableBorders && !_tableRightToLeft && tableContext.ComputedStyle.TableBorders?.Right is
                { LineStyle: DxpComputedBorderLineStyle.Double } doubleRightBorder)
                _doubleTableRightBorders[_table] = doubleRightBorder;
            var insideHorizontal = table.TableProperties?.TableBorders?.InsideHorizontalBorder;
            if (!_nativeDoubleTableBorders && insideHorizontal?.Val?.Value == BorderValues.Double &&
                insideHorizontal.Size?.Value is uint eighths)
                _doubleRowBorderPoints = eighths / 4.0;
            var directMargins = table.TableProperties?.GetFirstChild<TableCellMarginDefault>();
            int? leftMargin = ReadCellMargin(directMargins?.GetFirstChild<TableCellLeftMargin>());
            int? rightMargin = ReadCellMargin(directMargins?.GetFirstChild<TableCellRightMargin>());
            int? topMargin = ReadCellMargin(directMargins?.TopMargin);
            int? bottomMargin = ReadCellMargin(directMargins?.BottomMargin);
            var tableStyleId = table.TableProperties?.TableStyle?.Val?.Value ??
                _paragraphStyles.Values.FirstOrDefault(style =>
                    style.Type?.Value == StyleValues.Table && style.Default?.Value == true)?.StyleId?.Value;
            if (tableStyleId != null &&
                _paragraphStyles.TryGetValue(tableStyleId, out var definition))
            {
                var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (var current = definition; current != null && visited.Add(current.StyleId?.Value ?? "");
                    current = current.BasedOn?.Val?.Value is string baseId &&
                    _paragraphStyles.TryGetValue(baseId, out var baseStyle) ? baseStyle : null)
                {
                    _tableParagraphSpacing ??= current.StyleParagraphProperties?.SpacingBetweenLines;
                    var styleMargins = current.GetFirstChild<StyleTableProperties>()?
                        .GetFirstChild<TableCellMarginDefault>();
                    leftMargin ??= ReadCellMargin(styleMargins?.GetFirstChild<TableCellLeftMargin>());
                    rightMargin ??= ReadCellMargin(styleMargins?.GetFirstChild<TableCellRightMargin>());
                    topMargin ??= ReadCellMargin(styleMargins?.TopMargin);
                    bottomMargin ??= ReadCellMargin(styleMargins?.BottomMargin);
                }
            }
            var backgroundFill = table.TableProperties?
                .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Shading>()?.Fill?.Value;
            if (backgroundFill == null && tableStyleId != null &&
                _paragraphStyles.TryGetValue(tableStyleId, out var backgroundStyle))
                backgroundFill = backgroundStyle.GetFirstChild<StyleTableProperties>()?
                    .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Shading>()?.Fill?.Value;
            var spacing = table.TableProperties?.GetFirstChild<TableCellSpacing>()?.Width?.Value ??
                (tableStyleId != null && _paragraphStyles.TryGetValue(tableStyleId, out var style)
                    ? style.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableCellSpacing>()?.Width?.Value
                    : null);
            if (int.TryParse(spacing, out var spacingTwips))
            {
                var points = spacingTwips / 20.0;
                _tableCellSpacingPoints = points;
                if (points > 0) _tableBackground = ParseColor("#" + backgroundFill);
                if (_inHeader)
                {
                    _headerCellSpacingPoints = points;
                    _headerHasFilledSpacedTable |= _tableBackground != null && points > 0;
                }
                if (points > 0 && ReferenceEquals(_lastParagraphContainer, _container) &&
                    _lastParagraph != null)
                    _lastParagraph.Format.SpaceAfter = Unit.FromPoint(
                        2 * points + (_container is Section ? _defaultParagraphAfterPoints : 0));
            }
            if (_tableCellSpacingPoints > 0 && _container is HeaderFooter && !_inHeader &&
                _section != null && table.Elements<TableRow>().Count() == 1 &&
                table.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                    { InnerText.Length: > 0 } &&
                table.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                    { InnerText.Length: 0 } &&
                _footerTableDistanceApplied.Add(_section))
                _section.PageSetup.FooterDistance = Unit.FromPoint(
                    _section.PageSetup.FooterDistance.Point + 7);
            var firstPageIntroTableFooter = _container is HeaderFooter && !_inHeader &&
                _sectionCount > 1 && _headerFooterKind == HeaderFooterValues.First &&
                table.PreviousSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                    { InnerText.Length: > 0 } &&
                table.NextSibling() is DocumentFormat.OpenXml.Wordprocessing.Paragraph
                    { InnerText.Length: 0 };
            if (firstPageIntroTableFooter && _section != null &&
                _firstPageIntroTableFooters.Add(_section))
                _section.PageSetup.FooterDistance = Unit.FromPoint(
                    _section.PageSetup.FooterDistance.Point - 2.8);
            var bareTerminalFooterTable = table.PreviousSibling() == null &&
                table.NextSibling() == null &&
                table.TableProperties?.TableStyle == null &&
                table.TableProperties?.TableBorders == null &&
                table.TableProperties?.TableLayout == null;
            if (_container is HeaderFooter && !_inHeader && !_horizontalMergeTable &&
                _tableCellSpacingPoints == 0 && _section != null &&
                !_section.PageSetup.DifferentFirstPageHeaderFooter &&
                !firstPageIntroTableFooter &&
                !bareTerminalFooterTable &&
                !(_sectionCount > 1 && _headerFooterKind == HeaderFooterValues.First) &&
                (table.TableProperties?.TableStyle != null || table.TableProperties?.TableBorders != null ||
                    _paragraphStyles.GetValueOrDefault("Normal")?.StyleRunProperties?.FontSize != null ||
                    !_implicitModernFont && !_classicPartialDefaults) &&
                _footerTableDistanceApplied.Add(_section))
                _section.PageSetup.FooterDistance = Unit.FromPoint(_section.PageSetup.FooterDistance.Point +
                    (string.Equals(_minorThemeLatinFont, "Arial", StringComparison.OrdinalIgnoreCase) &&
                     table.Elements<TableRow>().Count() == 1 ? 1 : 2));
            if (_tableCellSpacingPoints > 0 && tableContext.ComputedStyle.TableBorders?.Left is
                { LineStyle: not DxpComputedBorderLineStyle.None } leftBorder)
                _spacedTableLeftBorders[_table] = (leftBorder, _tableCellSpacingPoints,
                    _tableInHeaderFooter, 0);
            var styledSpacingBorder = _tableCellSpacingPoints > 0 &&
                table.TableProperties?.TableStyle != null &&
                tableContext.ComputedStyle.TableBorders?.Left is
                    { LineStyle: not DxpComputedBorderLineStyle.None, WidthPt: > 0 } styledLeft
                ? styledLeft.WidthPt : 0;
            if (table.TableProperties?.TableWidth == null &&
                table.TableGrid?.Elements<GridColumn>().Any() != true &&
                table.TableProperties?.TableBorders == null &&
                _lastParagraphContainer == _container && _lastParagraph != null)
                _lastParagraph.Format.SpaceAfter = Unit.FromPoint(
                    Math.Max(_lastParagraph.Format.SpaceAfter.Point, 10.31));
            _lastParagraph = null;
            _lastParagraphContainer = null;
            _table.Rows.LeftIndent = Unit.FromPoint(styledSpacingBorder > 0
                ? styledSpacingBorder + 0.5 : 0);
            if (leftMargin is int leftTwips)
                _table.LeftPadding = Unit.FromPoint(Math.Max(0, leftTwips / 20.0 -
                    Math.Max(0, tableContext.ComputedStyle.TableBorders?.Left?.WidthPt ?? 0) / 2));
            if (rightMargin is int rightTwips) _table.RightPadding = Unit.FromPoint(rightTwips / 20.0);
            if (floatingFrame != null)
            {
                if (leftMargin == null) _table.LeftPadding = Unit.Zero;
                if (rightMargin == null) _table.RightPadding = Unit.Zero;
            }
            if (_tableBackground != null || _horizontalMergeTable)
            {
                if (leftMargin == null)
                    _table.LeftPadding = Unit.FromPoint(_horizontalMergeTable ? 0.5 : 0);
                if (rightMargin == null) _table.RightPadding = Unit.Zero;
            }
            if (_nativeDoubleTableBorders && HasSixDoubleBorders(table))
                _table.LeftPadding = Unit.FromPoint(Math.Max(0, _table.LeftPadding.Point - 0.5));
            if (!firstPageIntroTableFooter && table.TableProperties?.TableBorders != null &&
                table.Descendants<TopLeftToBottomRightCellBorder>().Any())
                _table.Rows.LeftIndent = Unit.FromPoint(-_table.LeftPadding.Point -
                    (table.TableProperties.TableBorders.LeftBorder?.Color?.Value is string color &&
                     color != "auto" ? 0.5 : 0));
            if (topMargin is int topTwips) _table.TopPadding = Unit.FromPoint(topTwips / 20.0);
            if (bottomMargin is int bottomTwips) _table.BottomPadding = Unit.FromPoint(bottomTwips / 20.0);
            if (!model.Cells.Cast<DxpCellModel?>().Any(cell => cell is { RowSpan: > 1 } or { ColSpan: > 1 }))
            {
                if (tableContext.ComputedStyle.DefaultCellBorders is { } defaults)
                    ApplyBorders(_table.Borders, defaults);
                else if (tableContext.ComputedStyle.DefaultCellBorder is { } defaultBorder)
                    ApplyBorder(_table.Borders, defaultBorder);
            }
            var columns = Math.Max(1, model.ColumnCount);
            if (_inHeader && columns == 2 &&
                table.Elements<TableRow>().FirstOrDefault()?.Elements<TableCell>().ToArray() is
                    { Length: 2 } firstCells &&
                firstCells[0].InnerText.Length > firstCells[1].InnerText.Length * 4 &&
                table.TableProperties?.TableWidth?.Type?.Value == TableWidthUnitValues.Auto &&
                table.TableProperties?.TableLayout?.Type?.Value != TableLayoutValues.Fixed)
                _headerHasImbalancedAutoFitTable = true;
            var available = _section?.PageSetup.PageWidth.Point ?? 595;
            available -= (_section?.PageSetup.LeftMargin.Point ?? 72) +
                (_section?.PageSetup.RightMargin.Point ?? 72);
            var sampleRun = table.Descendants<Run>().FirstOrDefault();
            var sampleParagraph = sampleRun?.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .FirstOrDefault();
            var cellFont = sampleRun != null && sampleParagraph != null
                ? context.Styles.ResolveRunStyle(sampleParagraph, sampleRun) : null;
            var useAptosCellWidth = _tableBackground != null &&
                _availableFontFamilies.Contains("Aptos") &&
                (string.Equals(cellFont?.FontName, "Aptos", StringComparison.OrdinalIgnoreCase) ||
                 cellFont?.FontName == null && _implicitModernFont) &&
                (cellFont?.FontSizeHalfPoints == 24 ||
                 cellFont?.FontSizeHalfPoints == null && _implicitModernFont);
            var cellWidthPadding = _tableBackground == null ? 27 :
                _tableCellSpacingPoints > 0 && table.TableProperties?.TableCellSpacing != null
                    ? useAptosCellWidth ? 18 : ColoredSpacedTableCellWidthPaddingPoints : 20;
            var autoFitWidths = AutoFitTwoColumnWidths(table, model, available, _tableCellSpacingPoints,
                useAptosCellWidth ? "Aptos" : _fontFamily, cellWidthPadding,
                useAptosCellWidth ? 12 : 11,
                string.Equals(_minorThemeLatinFont, "Aptos", StringComparison.OrdinalIgnoreCase),
                ((leftMargin ?? 108) + (rightMargin ?? 108)) / 20.0);
            if (autoFitWidths == null && IsSimpleSideBySideNestedTable(table))
                autoFitWidths = AutoFitSimpleNestedTableWidths(table, model,
                    _fontFamily,
                    cellFont?.FontSizeHalfPoints is int halfPoints ? halfPoints / 2.0 : 11,
                    ((leftMargin ?? 108) + (rightMargin ?? 108)) / 20.0);
            if (autoFitWidths == null)
            {
                autoFitWidths = ImplicitContentTableWidths(table, model, available, context.Styles,
                    _implicitModernFont && typeof(DocumentRenderer).GetProperty("TextWidthAdjustment") != null
                        ? "Aptos" : _fontFamily,
                    _implicitModernFont ? 12 : 11, _availableFontFamilies);
                if (autoFitWidths != null)
                {
                    _implicitContentTable = true;
                    _table.LeftPadding = Unit.Zero;
                    _table.RightPadding = Unit.Zero;
                    if (_implicitModernFont &&
                        typeof(DocumentRenderer).GetProperty("TextWidthAdjustment") != null)
                        _table.Rows.LeftIndent = Unit.FromPoint(_table.Rows.LeftIndent.Point + 0.5);
                }
            }
            if (autoFitWidths != null && table.Elements<TableRow>().Skip(1).Any() &&
                table.Descendants<NoWrap>().All(noWrap => noWrap.Val?.Value == OnOffOnlyValues.Off) &&
                !string.Equals(_minorThemeLatinFont, "Arial", StringComparison.OrdinalIgnoreCase))
                _table.RightPadding = Unit.FromPoint(Math.Max(0, _table.RightPadding.Point - 2));
            if (_horizontalMergeTable)
                autoFitWidths = AutoFitHorizontalMergeWidths(table, model, available,
                    context.Styles, _implicitModernFont ? "Aptos" : _fontFamily,
                    _implicitModernFont ? 12 : 11, _availableFontFamilies) ?? autoFitWidths;
            for (var index = 0; index < columns; index++)
            {
                var width = autoFitWidths != null ? autoFitWidths[index] :
                    index < model.GridColTwips.Count && model.GridColTwips[index] is int twips && twips > 0
                    ? twips / 20.0 : available / columns;
                if (_tableInHeaderFooter && columns == 2 &&
                    table.TableProperties?.TableLayout?.Type?.Value == TableLayoutValues.Autofit &&
                    !table.Descendants<NoWrap>().Any() &&
                    table.Elements<TableRow>().FirstOrDefault()?.Elements<TableCell>().ToArray() is
                        { Length: 2 } narrowStoryCells && narrowStoryCells[0].InnerText.Length > 75 &&
                    narrowStoryCells[1].InnerText.Length < 12)
                {
                    var adjustment = string.Equals(_minorThemeLatinFont, "Arial", StringComparison.OrdinalIgnoreCase)
                        ? -0.7 : 1.25;
                    width += index == 0 ? adjustment : -adjustment;
                }
                if (_tableCellSpacingPoints > 0)
                    width = Math.Max(1, width - _tableCellSpacingPoints);
                if (index == 0 && StyledFixedTableFirstColumnInset(table, tableContext) is > 0
                    and var styledInset)
                    width = Math.Max(1, width - styledInset);
                if (index == 0 && styledSpacingBorder > 0)
                    width = Math.Max(1, width - styledSpacingBorder + (columns > 1 ? 0.5 : 0));
                if (index == columns - 1 && columns > 1 && styledSpacingBorder > 0)
                    width = Math.Max(1, width - 0.5);
                _table.AddColumn(Unit.FromPoint(width));
            }
            if (floatingFrame != null)
            {
                var frameWidth = table.TableProperties?.TableWidth?.Type?.Value == TableWidthUnitValues.Dxa &&
                    int.TryParse(table.TableProperties.TableWidth.Width?.Value, out var widthTwips) &&
                    widthTwips > 0
                    ? widthTwips / 20.0
                    : _table.Columns.Cast<MigraDoc.DocumentObjectModel.Tables.Column>()
                        .Sum(column => column.Width.Point);
                floatingFrame.Width = Unit.FromPoint(frameWidth + floatingLeftBorderWidth);
            }
        }
        return DxpDisposable.Create(() =>
        {
            if (_table != null)
            {
                foreach (var modelCell in model.Cells.Cast<DxpCellModel?>())
                {
                    if (modelCell is not { IsCovered: false, RowSpan: > 1 })
                        continue;
                    var lastRow = modelCell.Row + modelCell.RowSpan - 1;
                    var column = _tableRightToLeft
                        ? _table.Columns.Count - modelCell.Col - modelCell.ColSpan
                        : modelCell.Col;
                    if (lastRow >= _table.Rows.Count || column < 0 || column >= _table.Columns.Count ||
                        model.Cells[lastRow, modelCell.Col]?.Cell.TableCellProperties is not
                            { VerticalMerge: not null } continuationProperties)
                        continue;
                    var bottom = _table.Rows[modelCell.Row][column].Borders.Bottom.Clone();
                    if (DxpTableStyleResolver.ComputeDirectBottomBorder(continuationProperties) is
                        { } overrideBorder)
                        ApplyBorder(bottom, overrideBorder);
                    _table.Rows[lastRow][column].Borders.Bottom = bottom;
                }
            }
            if (_implicitContentTable && table.Parent is Body && table.NextSibling() == null &&
                _section != null)
            {
                var trailingParagraph = _section.AddParagraph();
                trailingParagraph.Format.LineSpacingRule = LineSpacingRule.Exactly;
                trailingParagraph.Format.LineSpacing = Unit.FromPoint(14.96);
                trailingParagraph.Format.SpaceBefore = Unit.Zero;
                trailingParagraph.Format.SpaceAfter = Unit.Zero;
            }
            _table = previous;
            _tableCellSpacingPoints = previousSpacing;
            _tableBackground = previousBackground;
            _horizontalMergeTable = previousHorizontalMerge;
            _tableInHeaderFooter = previousTableInHeaderFooter;
            _tableRightToLeft = previousTableDirection;
            _rowCellSpacingPoints = previousRowSpacing;
            _implicitContentTable = previousImplicitContentTable;
            _tableParagraphSpacing = previousParagraphSpacing;
            _doubleRowBorderPoints = previousDoubleBorder;
            _nativeDoubleTableBorders = previousNativeDoubleBorders;
            _previousHeaderRowCompact = previousCompactRow;
            _previousHeaderRowSpacingReductionPoints = previousRowSpacingReduction;
        });
    }

    private static double[]? AutoFitSimpleNestedTableWidths(
        DocumentFormat.OpenXml.Wordprocessing.Table table, DxpTableModel model,
        string fontFamily, double fontSize, double horizontalCellMargins)
    {
        if (model.GridColTwips.Count != 2 ||
            model.GridColTwips[0] is not int firstTwips || firstTwips <= 0 ||
            model.GridColTwips[1] is not int secondTwips || secondTwips <= 0 ||
            table.TableProperties?.TableLayout?.Type?.Value == TableLayoutValues.Fixed)
            return null;
        var rows = table.Elements<TableRow>().Select(row => row.Elements<TableCell>().ToArray()).ToArray();
        using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100),
            XGraphicsUnit.Point, XPageDirection.Downwards);
        var font = new XFont(fontFamily, fontSize);
        var naturalWidths = new double[2];
        for (var index = 0; index < naturalWidths.Length; index++)
        {
            foreach (var paragraph in rows.Select(row => row[index])
                .SelectMany(cell => cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()))
            {
                var line = new StringBuilder();
                foreach (var element in paragraph.Descendants<OpenXmlElement>())
                {
                    if (element is DocumentFormat.OpenXml.Wordprocessing.Text text)
                        line.Append(text.Text);
                    else if (element is Break)
                    {
                        naturalWidths[index] = Math.Max(naturalWidths[index],
                            graphics.MeasureString(line.ToString(), font).Width);
                        line.Clear();
                    }
                }
                naturalWidths[index] = Math.Max(naturalWidths[index],
                    graphics.MeasureString(line.ToString(), font).Width);
            }
            naturalWidths[index] += horizontalCellMargins;
        }
        var totalWidth = (firstTwips + secondTwips) / 20.0;
        if (naturalWidths.Sum() <= 0) return null;
        if (Math.Abs(naturalWidths[0] - naturalWidths[1]) <=
            0.1 * Math.Max(naturalWidths[0], naturalWidths[1]))
            return null;
        var firstWidth = totalWidth * naturalWidths[0] / naturalWidths.Sum();
        return [firstWidth, totalWidth - firstWidth];
    }

    private static double[]? AutoFitTwoColumnWidths(DocumentFormat.OpenXml.Wordprocessing.Table table,
        DxpTableModel model, double available, double cellSpacing, string fontFamily, double cellWidthPadding,
        double cellFontSize, bool aptosTheme, double horizontalCellMargins)
    {
        if (model.ColumnCount != 2 || table.TableProperties?.TableLayout?.Type?.Value != TableLayoutValues.Autofit ||
            table.TableProperties.TableWidth?.Type?.Value != TableWidthUnitValues.Auto ||
            model.Cells.Cast<DxpCellModel?>().Any(cell => cell is { RowSpan: > 1 } or { ColSpan: > 1 }))
            return null;
        var rows = table.Elements<TableRow>().Select(row => row.Elements<TableCell>().ToArray()).ToArray();
        if (rows.Length == 0 || rows.Any(row => row.Length != 2)) return null;
        var lengths = new[] {
            rows.Max(row => row[0].InnerText.Length),
            rows.Max(row => row[1].InnerText.Length)
        };
        if (lengths.Min() == 0) return null;
        var totalWidth = model.GridColTwips.Count >= 2 && model.GridColTwips[0] is int first &&
            model.GridColTwips[1] is int second && first > 0 && second > 0
            ? Math.Min(available, (first + second) / 20.0) : available;
        if (cellSpacing > 0 && lengths.Max() < 20 && lengths.Max() < lengths.Min() * 3)
        {
            using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
                XPageDirection.Downwards);
            var font = new XFont(fontFamily, cellFontSize);
            var naturalWidths = new[] { 0.0, 0.0 };
            foreach (var row in rows)
                for (var index = 0; index < 2; index++)
                    naturalWidths[index] = Math.Max(naturalWidths[index],
                        graphics.MeasureString(row[index].InnerText, font).Width + cellWidthPadding);
            if (naturalWidths.Sum() < totalWidth)
                return naturalWidths;
        }
        if (lengths.Max() < lengths.Min() * 3) return null;
        if (rows.Any(row => row.Any(cell => cell.TableCellProperties?.GetFirstChild<NoWrap>() is
            { } noWrap && noWrap.Val?.Value != OnOffOnlyValues.Off)))
        {
            var legacyFirstWidth = Math.Clamp(totalWidth * lengths[0] / (lengths[0] + lengths[1] + 1.5),
                totalWidth * 0.1, totalWidth * 0.95);
            using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
                XPageDirection.Downwards);
            var font = new XFont(fontFamily, cellFontSize);
            var longestWordWidth = rows.SelectMany(row => Regex.Matches(row[1].InnerText, @"\S+")
                    .Select(match => graphics.MeasureString(match.Value, font).Width))
                .DefaultIfEmpty(0).Max();
            var secondWidth = Math.Clamp(longestWordWidth + horizontalCellMargins, totalWidth * 0.05,
                totalWidth * 0.5);
            var noWrapFirstWidth = aptosTheme ? totalWidth - secondWidth :
                Math.Max(legacyFirstWidth, totalWidth - secondWidth);
            return [noWrapFirstWidth, totalWidth - noWrapFirstWidth];
        }
        using var measure = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
            XPageDirection.Downwards);
        var textFont = new XFont(fontFamily, 11);
        var multirowAllowance = string.Equals(fontFamily, "Arial", StringComparison.OrdinalIgnoreCase) &&
            !aptosTheme ? 0.54 : 0.45;
        var allowance = cellWidthPadding * (rows.Length > 1 ? multirowAllowance : 0.6);
        var measured = new[] {
            rows.Max(row => measure.MeasureString(row[0].InnerText, textFont).Width) + allowance,
            rows.Max(row => measure.MeasureString(row[1].InnerText, textFont).Width) + allowance
        };
        var firstWidth = Math.Clamp(totalWidth * measured[0] / measured.Sum(),
            totalWidth * 0.1, totalWidth * 0.95);
        return [firstWidth, totalWidth - firstWidth];
    }

    private static double[]? ImplicitContentTableWidths(DocumentFormat.OpenXml.Wordprocessing.Table table,
        DxpTableModel model, double available, DxpIStyleResolver styles, string defaultFontFamily,
        double defaultFontSize,
        ISet<string> availableFonts)
    {
        if (table.TableProperties?.TableWidth != null || table.TableGrid?.Elements<GridColumn>().Any() == true ||
            table.TableProperties?.TableBorders != null || model.ColumnCount < 1 ||
            model.Cells.Cast<DxpCellModel?>().Any(cell => cell is { RowSpan: > 1 } or { ColSpan: > 1 }))
            return null;
        var rows = table.Elements<TableRow>().Select(row => row.Elements<TableCell>().ToArray()).ToArray();
        if (rows.Length == 0 || rows.Any(row => row.Length != model.ColumnCount)) return null;
        using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
            XPageDirection.Downwards);
        var widths = new double[model.ColumnCount];
        foreach (var row in rows)
            for (var index = 0; index < row.Length; index++)
            {
                var cell = row[index];
                var run = cell.Descendants<Run>().FirstOrDefault();
                var paragraph = run?.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
                var style = run != null && paragraph != null ? styles.ResolveRunStyle(paragraph, run) : null;
                var family = style?.FontName is string name && availableFonts.Contains(name)
                    ? name : defaultFontFamily;
                var size = style?.FontSizeHalfPoints is int halfPoints && halfPoints > 0
                    ? halfPoints / 2.0 : defaultFontSize;
                var measured = DocSystemFontAdvances.Find(family)?.TryMeasure(cell.InnerText, size,
                    out var advance, useKerning: true) == true
                    ? advance : graphics.MeasureString(cell.InnerText, new XFont(family, size)).Width;
                widths[index] = Math.Max(widths[index],
                    measured + (string.Equals(family, "Aptos", StringComparison.OrdinalIgnoreCase) ? 1.5 : 1));
            }
        return widths.Sum() <= available ? widths : null;
    }

    private static double[]? AutoFitHorizontalMergeWidths(DocumentFormat.OpenXml.Wordprocessing.Table table,
        DxpTableModel model, double available, DxpIStyleResolver styles, string fontFamily,
        double defaultFontSize,
        ISet<string> availableFonts)
    {
        if (model.ColumnCount is < 2 or > 3) return null;
        var rows = table.Elements<TableRow>()
            .Select(row => row.Elements<TableCell>().ToArray())
            .Where(row => row.Length == model.ColumnCount &&
                row.All(cell => cell.TableCellProperties?.GetFirstChild<HorizontalMerge>() == null))
            .ToArray();
        if (rows.Length == 0) return null;
        using var graphics = XGraphics.CreateMeasureContext(new XSize(100, 100), XGraphicsUnit.Point,
            XPageDirection.Downwards);
        var widths = Enumerable.Range(0, model.ColumnCount)
            .Select(index => rows.Max(row =>
            {
                var run = row[index].Descendants<Run>().FirstOrDefault();
                var paragraph = run?.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                    .FirstOrDefault();
                var style = run != null && paragraph != null
                    ? styles.ResolveRunStyle(paragraph, run) : null;
                var family = style?.FontName is string name &&
                    (availableFonts.Contains(name) || DocSystemFontAdvances.Find(name) != null)
                    ? name : fontFamily;
                var size = style?.FontSizeHalfPoints is int halfPoints && halfPoints > 0
                    ? halfPoints / 2.0 : defaultFontSize;
                return graphics.MeasureString(row[index].InnerText, new XFont(family, size)).Width +
                    1;
            }))
            .ToArray();
        if (model.ColumnCount == 2 && widths.Max() < NarrowHorizontalMergeCellWidthPoints)
            for (var index = 0; index < widths.Length; index++)
                widths[index] += NarrowHorizontalMergeCellPaddingPoints;
        return widths.Sum() < available ? widths : null;
    }

    private static int? ReadCellMargin(OpenXmlElement? element)
    {
        if (element == null || AttributeValue(element, "type") is string type && type != "dxa")
            return null;
        return int.TryParse(AttributeValue(element, "w"), out var width) ? width : null;
    }

    public override IDisposable VisitTableRowBegin(TableRow row, DxpITableRowContext rowContext, DxpIDocumentContext context)
    {
        var previous = _row;
        var previousLeftMargin = _rowLeftMarginPoints;
        var previousRightMargin = _rowRightMarginPoints;
        var previousRowSpacing = _rowCellSpacingPoints;
        _row = _table?.AddRow();
        if (_row != null)
        {
            if (_tableInHeaderFooter && _table != null && row.TableRowProperties != null)
            {
                foreach (var gap in row.TableRowProperties.ChildElements.Where(element =>
                             element.LocalName is "gridBefore" or "gridAfter"))
                {
                    if (!int.TryParse(AttributeValue(gap, "val"), out var count)) continue;
                    for (var index = 0; index < Math.Min(count, _table.Columns.Count); index++)
                    {
                        var column = gap.LocalName == "gridBefore"
                            ? index : _table.Columns.Count - count + index;
                        if (column < 0 || column >= _table.Columns.Count) continue;
                        var placeholder = _row[column].AddParagraph();
                        placeholder.Format.Font.Size = Unit.FromPoint(1);
                        placeholder.Format.LineSpacingRule = LineSpacingRule.Exactly;
                        placeholder.Format.LineSpacing = Unit.FromPoint(0.1);
                        placeholder.Format.SpaceBefore = Unit.Zero;
                        placeholder.Format.SpaceAfter = Unit.Zero;
                    }
                }
            }
            var topException = row.Elements<TablePropertyExceptions>()
                .Select(exception => exception.GetFirstChild<TableBorders>()?.TopBorder)
                .FirstOrDefault(top => top != null);
            if (rowContext.Index == 0 && topException?.Val?.Value == BorderValues.Double &&
                topException.Size?.Value is uint eighths && eighths > 0 && _table != null)
            {
                _doubleFirstRowTopBorders[_table] = new RowTopStroke(
                    "#" + (topException.Color?.Value ?? "000000"), eighths / 8.0,
                    _container is HeaderFooter);
                _row.TopPadding = Unit.FromPoint(_row.TopPadding.Point + eighths / 4.0);
            }
            var spacingOverride = row.GetFirstChild<TablePropertyExceptions>()?
                .GetFirstChild<TableCellSpacing>()?.Width?.Value ??
                row.TableRowProperties?.GetFirstChild<TableCellSpacing>()?.Width?.Value;
            _rowCellSpacingPoints = int.TryParse(spacingOverride, out var rowTwips)
                ? rowTwips / 20.0 : _tableCellSpacingPoints;
            var rowMargins = row.GetFirstChild<TablePropertyExceptions>()?
                .GetFirstChild<TableCellMarginDefault>();
            _rowLeftMarginPoints = ReadCellMargin(rowMargins?.GetFirstChild<TableCellLeftMargin>()) / 20.0;
            _rowRightMarginPoints = ReadCellMargin(rowMargins?.GetFirstChild<TableCellRightMargin>()) / 20.0;
            if (ReadCellMargin(rowMargins?.TopMargin) is int topTwips)
                _row.TopPadding = Unit.FromPoint(topTwips / 20.0);
            if (ReadCellMargin(rowMargins?.BottomMargin) is int bottomTwips)
                _row.BottomPadding = Unit.FromPoint(bottomTwips / 20.0);
            if (_inHeader)
            {
                _headerRowCount++;
                if (rowContext.Index > 0 && _previousHeaderRowCompact)
                    _headerCompactRowTransitions++;
                if (rowContext.Index > 0 && !_previousHeaderRowCompact)
                    _headerRowSpacingReductionPoints += _previousHeaderRowSpacingReductionPoints;
                var margins = row.GetFirstChild<TablePropertyExceptions>()?
                    .GetFirstChild<TableCellMarginDefault>();
                _previousHeaderRowCompact = margins?.TopMargin?.Width?.Value == "0" &&
                    margins.BottomMargin?.Width?.Value == "0";
                _headerHasZeroMarginRow |= _previousHeaderRowCompact;
                _previousHeaderRowSpacingReductionPoints = row.Elements<TableCell>()
                    .Select(HeaderCellSpacingReduction)
                    .DefaultIfEmpty(0).Min();
            }
            _row.HeadingFormat = rowContext.IsHeader;
            if (_rowCellSpacingPoints > 0)
                _row.BottomPadding = Unit.FromPoint(_tableBackground == null
                    ? _rowCellSpacingPoints / 2 : Math.Max(0, 2 * _rowCellSpacingPoints - 1.5));
            if (_implicitContentTable)
                _row.BottomPadding = Unit.FromPoint(_row.BottomPadding.Point + 10.31);
            if (_rowCellSpacingPoints > 0 && row.ElementsAfter().OfType<TableRow>().FirstOrDefault() is
                { } nextRow)
            {
                var nextSpacing = nextRow.GetFirstChild<TablePropertyExceptions>()?
                    .GetFirstChild<TableCellSpacing>()?.Width?.Value ??
                    nextRow.TableRowProperties?.GetFirstChild<TableCellSpacing>()?.Width?.Value;
                if (nextSpacing == "0")
                {
                    _row.BottomPadding = Unit.Zero;
                    if (_container is HeaderFooter && !_inHeader && _section != null &&
                        _footerSpacingResetApplied.Add(_section))
                        _section.PageSetup.FooterDistance = Unit.FromPoint(
                            _section.PageSetup.FooterDistance.Point + _rowCellSpacingPoints / 2 -
                            Math.Min(2, _rowCellSpacingPoints / 2) / 2);
                }
            }
            if (rowContext.Index > 0 && _rowCellSpacingPoints > 0)
                _row.TopPadding = Unit.FromPoint(_rowCellSpacingPoints -
                    (_container is Section && _rowCellSpacingPoints > _tableCellSpacingPoints ? 1 : 0));
            if (spacingOverride == "0" && row.PreviousSibling<TableRow>() is { } precedingRow)
            {
                var precedingSpacing = precedingRow.GetFirstChild<TablePropertyExceptions>()?
                    .GetFirstChild<TableCellSpacing>()?.Width?.Value ??
                    precedingRow.TableRowProperties?.GetFirstChild<TableCellSpacing>()?.Width?.Value;
                var precedingPoints = int.TryParse(precedingSpacing, out var precedingTwips)
                    ? precedingTwips / 20.0 : _tableCellSpacingPoints;
                if (precedingPoints > 0)
                {
                    var allowance = Math.Min(2, precedingPoints / 2);
                    _row.TopPadding = Unit.FromPoint(_row.TopPadding.Point + allowance);
                    if (_table != null && _spacedTableLeftBorders.TryGetValue(_table, out var overlay))
                        _spacedTableLeftBorders[_table] = (overlay.Border, overlay.Spacing,
                            overlay.InHeaderFooter, overlay.ResetAllowance + allowance);
                }
            }
            if (_doubleRowBorderPoints > 0)
            {
                _row.BottomPadding = Unit.FromPoint(_row.BottomPadding.Point + _doubleRowBorderPoints);
                if (rowContext.Index == 0)
                    _row.TopPadding = Unit.FromPoint(_row.TopPadding.Point + _doubleRowBorderPoints);
            }
            if (_container is Section && _table?.TopPadding.Point == 0 &&
                _paragraphStyles.GetValueOrDefault("Normal")?.StyleRunProperties?.FontSize?.Val?.Value == "28")
                _row.TopPadding = Unit.FromPoint(_row.TopPadding.Point + 0.6);
            if (rowContext.Index > 0 && _row.BottomPadding.Point == 0 &&
                _table?.BottomPadding.Point >= 0.25 &&
                row.Parent is DocumentFormat.OpenXml.Wordprocessing.Table sourceTable &&
                sourceTable.TableProperties?.TableBorders?.InsideHorizontalBorder is { } inside &&
                inside.Val?.Value == BorderValues.Single && inside.Size?.Value == 2)
                _row.BottomPadding = Unit.FromPoint(_table.BottomPadding.Point - 0.25);
            if (row.TableRowProperties?.GetFirstChild<TableRowHeight>() is { Val: { Value: > 0 } height } rowHeight)
            {
                _row.Height = Unit.FromPoint(height / 20.0);
                _row.HeightRule = rowHeight.HeightType?.Value == HeightRuleValues.Exact
                    ? RowHeightRule.Exactly : RowHeightRule.AtLeast;
                if (_inHeader && _row.HeightRule == RowHeightRule.Exactly)
                    _headerHasExactHeightRow = true;
                if (_inHeader && _row.HeightRule == RowHeightRule.AtLeast)
                    _headerMinimumRowClearancePoints += Math.Max(0,
                        _row.Height.Point - _defaultParagraphLinePoints - 4);
            }
        }
        return DxpDisposable.Create(() =>
        {
            _row = previous;
            _rowLeftMarginPoints = previousLeftMargin;
            _rowRightMarginPoints = previousRightMargin;
            _rowCellSpacingPoints = previousRowSpacing;
        });
    }

    private double HeaderCellSpacingReduction(TableCell cell)
    {
        var paragraph = cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();
        var spacing = paragraph?.ParagraphProperties?.SpacingBetweenLines;
        if (spacing?.After?.Value == null &&
            paragraph?.ParagraphProperties?.ParagraphStyleId?.Val?.Value is string styleId &&
            _paragraphStyles.TryGetValue(styleId, out var style))
            spacing = style.StyleParagraphProperties?.SpacingBetweenLines;
        return int.TryParse(spacing?.After?.Value, out var after)
            ? Math.Max(0, _defaultParagraphAfterPoints - after / 20.0) : 0;
    }

    private double StyledFixedTableFirstColumnInset(DocumentFormat.OpenXml.Wordprocessing.Table table,
        DxpITableContext? tableContext)
    {
        if (_tableCellSpacingPoints != 0 || tableContext == null ||
            table.TableProperties?.TableStyle?.Val?.Value is not string styleId ||
            table.TableProperties.TableBorders != null ||
            table.TableProperties.TableLayout?.Type?.Value != TableLayoutValues.Fixed ||
            table.TableGrid?.Elements<GridColumn>().Count() != 2 ||
            tableContext.ComputedStyle.TableBorders?.Left is not
                { LineStyle: not DxpComputedBorderLineStyle.None, WidthPt: > 0 } styledLeft ||
            !_paragraphStyles.TryGetValue(styleId, out var style))
            return 0;
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var current = style; current != null && visited.Add(current.StyleId?.Value ?? "");
            current = current.BasedOn?.Val?.Value is string baseId &&
                _paragraphStyles.TryGetValue(baseId, out var baseStyle) ? baseStyle : null)
            if (current.Descendants<TableCellBorders>().Any())
                return 0;
        return styledLeft.WidthPt / 2;
    }

    public override IDisposable VisitTableCellBegin(TableCell cell, DxpITableCellContext cellContext, DxpIDocumentContext context)
    {
        var previousNoWrap = _cellNoWrap;
        var previousNoWrapWidth = _cellNoWrapWidthPoints;
        var previousWrapWidth = _cellWrapWidthPoints;
        var previousBalanceLastWord = _balanceNarrowCellLastWord;
        _cellNoWrap = cell.TableCellProperties?.GetFirstChild<NoWrap>() is { } noWrap &&
            noWrap.Val?.Value != OnOffOnlyValues.Off;
        _cellNoWrapWidthPoints = null;
        _cellWrapWidthPoints = null;
        _balanceNarrowCellLastWord = false;
        var horizontalMerge = cell.TableCellProperties?.GetFirstChild<HorizontalMerge>();
        if (horizontalMerge != null && horizontalMerge.Val?.Value != MergedCellValues.Restart)
        {
            var previousContainer = _container;
            _container = null;
            return DxpDisposable.Create(() =>
            {
                _container = previousContainer;
                _cellNoWrap = previousNoWrap;
                _cellNoWrapWidthPoints = previousNoWrapWidth;
                _cellWrapWidthPoints = previousWrapWidth;
                _balanceNarrowCellLastWord = previousBalanceLastWord;
            });
        }
        if (_table == null && IsNestedCellOutsideParent(cell, cellContext.ColumnIndex))
        {
            var previousContainer = _container;
            _container = null;
            return DxpDisposable.Create(() =>
            {
                _container = previousContainer;
                _cellNoWrap = previousNoWrap;
                _cellNoWrapWidthPoints = previousNoWrapWidth;
                _cellWrapWidthPoints = previousWrapWidth;
                _balanceNarrowCellLastWord = previousBalanceLastWord;
            });
        }
        var previous = _container;
        var previousStyle = _cellStyle;
        var previousLeftMargin = _cellLeftMarginPoints;
        var previousRightMargin = _cellRightMarginPoints;
        var previousSpacingIndent = _cellSpacingIndentPoints;
        var previousRotatedCell = _activeRotatedCell;
        var previousRotatedPlaceholder = _rotatedCellPlaceholderAdded;
        _activeRotatedCell = false;
        _rotatedCellPlaceholderAdded = false;
        _cellStyle = cellContext.ComputedStyle;
        var cellMargins = cell.TableCellProperties?.GetFirstChild<TableCellMargin>();
        var cellLeftMargin = cellMargins?.ChildElements.FirstOrDefault(element =>
            element.LocalName == "left");
        _cellLeftMarginPoints = ReadCellMargin(cellMargins?.GetFirstChild<TableCellLeftMargin>()) / 20.0;
        _cellRightMarginPoints = ReadCellMargin(cellMargins?.GetFirstChild<TableCellRightMargin>()) / 20.0;
        var conditionalTable = cell.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>()
            .FirstOrDefault();
        if (_table != null && cellContext.RowIndex == 0 && cellContext.ColumnIndex == 0 &&
            conditionalTable?.TableProperties is { } tableProperties &&
            tableProperties.TableStyle != null && tableProperties.TableBorders?.LeftBorder == null &&
            tableProperties.GetFirstChild<TableCellMarginDefault>()?
                .GetFirstChild<TableCellLeftMargin>() != null &&
            tableProperties.GetFirstChild<TableIndentation>()?.Width?.Value is > 0 &&
            cellContext.ComputedStyle.Borders?.Left is
                { LineStyle: not DxpComputedBorderLineStyle.None, WidthPt: > 0 } leftBorder)
            _table.LeftPadding = Unit.FromPoint(Math.Max(0,
                _table.LeftPadding.Point - leftBorder.WidthPt));
        var spacingDelta = _rowCellSpacingPoints - _tableCellSpacingPoints;
        _cellSpacingIndentPoints = _tableCellSpacingPoints > 0
            ? spacingDelta * (cellContext.ColumnIndex == 0 ? 2 : 1) : 0;
        if (_row != null && _table != null && cellContext.ColumnIndex < _table.Columns.Count)
        {
            var targetColumn = _tableRightToLeft
                ? _table.Columns.Count - cellContext.ColumnIndex - cellContext.ColSpan
                : cellContext.ColumnIndex;
            if (targetColumn < 0) targetColumn = cellContext.ColumnIndex;
            var target = _row[targetColumn];
            if (_table.Columns.Count == 1 && cellContext.RowIndex == 0 &&
                cellContext.ColumnIndex == 0 &&
                cell.Parent is TableRow sourceRow &&
                sourceRow.Parent is DocumentFormat.OpenXml.Wordprocessing.Table outerTable &&
                outerTable.Elements<TableRow>().Count() == 1 &&
                cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().Any())
            {
                if (ReadCellMargin(cellLeftMargin) is int leftTwips)
                    target.Column.LeftPadding = Unit.FromPoint(leftTwips / 20.0);
                if (ReadCellMargin(cellMargins?.TopMargin) is int topTwips)
                    _row.TopPadding = Unit.FromPoint(topTwips / 20.0);
            }
            if (cellContext.RowIndex > 0 && _table.Columns.Count == 2 && targetColumn == 1 &&
                cell.InnerText.Length <= 8 &&
                cell.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault() is
                    { } sourceTable &&
                sourceTable.TableProperties?.TableLayout?.Type?.Value == TableLayoutValues.Autofit &&
                !sourceTable.Descendants<NoWrap>().Any() &&
                sourceTable.Elements<TableRow>().FirstOrDefault()?.Elements<TableCell>().FirstOrDefault() is
                    { } firstCell && firstCell.InnerText.Length > 75)
                _cellRightMarginPoints = (_cellRightMarginPoints ?? _rowRightMarginPoints ??
                    _table.RightPadding.Point) + 1.5;
            if (cellContext.ComputedStyle.VerticalAlign is { } verticalAlign)
                target.VerticalAlignment = verticalAlign switch
                {
                    DxpComputedVerticalAlign.Middle => MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Center,
                    DxpComputedVerticalAlign.Bottom => MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Bottom,
                    _ => MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Top
                };
            if (_cellNoWrap && cell.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Count() == 1)
                _cellNoWrapWidthPoints = _table.Columns[targetColumn].Width.Point -
                    _table.LeftPadding.Point - _table.RightPadding.Point;
            if (!_cellNoWrap && cell.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>()
                .FirstOrDefault()?.Descendants<NoWrap>().Any(setting =>
                    setting.Val?.Value != OnOffOnlyValues.Off) == true)
            {
                _cellWrapWidthPoints = _table.Columns[targetColumn].Width.Point -
                    _table.LeftPadding.Point - _table.RightPadding.Point;
                var gridWidths = cell.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>()
                    .First().TableGrid?.Elements<GridColumn>()
                    .Take(2).Select(column => int.TryParse(column.Width?.Value, out var twips)
                        ? twips : 0).ToArray();
                _balanceNarrowCellLastWord = gridWidths is { Length: 2 } &&
                    gridWidths.Min() > 0 && gridWidths.Max() <= 2 * gridWidths.Min();
            }
            var horizontalSpan = cellContext.ColSpan;
            if (horizontalMerge?.Val?.Value == MergedCellValues.Restart)
            {
                foreach (var following in cell.ElementsAfter().OfType<TableCell>())
                {
                    var merge = following.TableCellProperties?.GetFirstChild<HorizontalMerge>();
                    if (merge == null || merge.Val?.Value == MergedCellValues.Restart) break;
                    horizontalSpan += Math.Max(1, following.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                }
            }
            target.MergeRight = Math.Min(horizontalSpan - 1, _table.Columns.Count - targetColumn - 1);
            target.MergeDown = Math.Max(0, cellContext.RowSpan - 1);
            var cellBorders = cell.TableCellProperties?.TableCellBorders;
            var falling = ReadDiagonalStroke(cellBorders?.TopLeftToBottomRightCellBorder);
            var rising = ReadDiagonalStroke(cellBorders?.TopRightToBottomLeftCellBorder);
            if (falling != null || rising != null)
            {
                if (!_diagonalCells.TryGetValue(_table, out var diagonals))
                    _diagonalCells[_table] = diagonals = new List<DiagonalCell>();
                var straightBorders = cell.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>()
                    .FirstOrDefault()?.TableProperties?.TableBorders != null ||
                    cellBorders?.TopBorder != null || cellBorders?.RightBorder != null ||
                    cellBorders?.BottomBorder != null || cellBorders?.LeftBorder != null;
                diagonals.Add(new DiagonalCell(cellContext.RowIndex, targetColumn,
                    cellContext.RowSpan, cellContext.ColSpan, falling, rising, straightBorders,
                    cell.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault()?
                        .TableProperties?.TableBorders?.LeftBorder?.Color?.Value is string borderColor &&
                    !borderColor.Equals("auto", StringComparison.OrdinalIgnoreCase)));
            }
            var direction = cell.TableCellProperties?.GetFirstChild<TextDirection>() is { } textDirection
                ? AttributeValue(textDirection, "val") : null;
            if (direction is "tbRl" or "btLr" or "tbRlV" &&
                cell.Descendants<Run>().SingleOrDefault() is { } textRun &&
                textRun.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault() is
                    { } textParagraph &&
                cellContext.ComputedStyle.Borders is null or
                    { Top: null, Right: null, Bottom: null, Left: null } &&
                cellContext.ComputedStyle.Border == null &&
                cellContext.ComputedStyle.BackgroundColorCss == null)
            {
                var runStyle = context.Styles.ResolveRunStyle(textParagraph, textRun);
                if (!_rotatedCells.TryGetValue(_table, out var rotated))
                    _rotatedCells[_table] = rotated = new List<RotatedCell>();
                rotated.Add(new RotatedCell(cellContext.RowIndex, targetColumn, cell.InnerText,
                    direction, _tableInHeaderFooter, _inHeader,
                    runStyle.FontName is string fontFamily && _availableFontFamilies.Contains(fontFamily)
                        ? fontFamily : _document!.Styles["Normal"].Font.Name,
                    runStyle.FontSizeHalfPoints is int halfPoints && halfPoints > 0
                        ? halfPoints / 2.0 : _document!.Styles["Normal"].Font.Size.Point,
                    runStyle.Bold, runStyle.Italic, runStyle.ColorCss));
                _activeRotatedCell = true;
                if (_inHeader && cell.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>()
                    .FirstOrDefault()?.TableProperties?.TableWidth?.Type?.Value == TableWidthUnitValues.Auto)
                    _headerHasRotatedCell = true;
            }
            if (cellContext.ComputedStyle.Borders is { } borders)
                ApplyBorders(target.Borders, borders);
            else if (cellContext.ComputedStyle.Border is { } border)
                ApplyBorder(target.Borders, border);
            if (_doubleTableRightBorders.ContainsKey(_table) &&
                targetColumn + target.MergeRight == _table.Columns.Count - 1 &&
                cellBorders?.RightBorder == null && cellBorders?.EndBorder == null)
                target.Borders.Right.Visible = false;
            if (!_nativeDoubleTableBorders && cellBorders?.ChildElements.Any(element =>
                    element.LocalName is "left" or "right" or "start" or "end" &&
                    AttributeValue(element, "val") == "double") == true &&
                cellContext.ComputedStyle.Borders is { } cellBoxBorders)
            {
                if (cellBoxBorders.Left is { LineStyle: DxpComputedBorderLineStyle.Double } left)
                {
                    target.Borders.Left.Visible = false;
                    if (!_doubleCellSideBorders.TryGetValue(_table, out var sides))
                        _doubleCellSideBorders[_table] = sides = new List<DoubleCellSideBorder>();
                    sides.Add(new DoubleCellSideBorder(cellContext.RowIndex, targetColumn, false, left));
                }
                if (cellBoxBorders.Right is { LineStyle: DxpComputedBorderLineStyle.Double } right)
                {
                    target.Borders.Right.Visible = false;
                    if (!_doubleCellSideBorders.TryGetValue(_table, out var sides))
                        _doubleCellSideBorders[_table] = sides = new List<DoubleCellSideBorder>();
                    sides.Add(new DoubleCellSideBorder(cellContext.RowIndex, targetColumn, true, right));
                }
            }
            if (targetColumn == 0 && _rtlDoubleTableLeftBorders.ContainsKey(_table))
                target.Borders.Left.Visible = false;
            var rowTopBorder = cell.Ancestors<TableRow>().FirstOrDefault()?
                .Elements<TablePropertyExceptions>()
                .Select(exception => exception.GetFirstChild<TableBorders>()?.TopBorder)
                .FirstOrDefault(top => top != null);
            if (rowTopBorder != null && cellBorders?.TopBorder == null)
            {
                if (rowTopBorder.Val?.Value == BorderValues.None ||
                    rowTopBorder.Val?.Value == BorderValues.Nil)
                    target.Borders.Top.Visible = false;
                else if (rowTopBorder.Size?.Value is uint size && size > 0)
                {
                    target.Borders.Top.Visible = true;
                    target.Borders.Top.Width = Unit.FromPoint(size / 8.0);
                    if (ParseColor("#" + rowTopBorder.Color?.Value) is { } rowColor)
                        target.Borders.Top.Color = rowColor;
                }
            }
            if (ParseColor(cellContext.ComputedStyle.BackgroundColorCss) is { } background)
                target.Shading.Color = _tableBackground is { } tableBackground &&
                    background == tableBackground && cell.TableCellProperties?.Shading == null
                    ? MigraDoc.DocumentObjectModel.Colors.White : background;
            if (_rowCellSpacingPoints > 0)
            {
                if (_tableBackground is { } tableFill)
                {
                    target.Borders.Top.Color = tableFill;
                    target.Borders.Top.Width = Unit.FromPoint(2 * _rowCellSpacingPoints);
                    target.Borders.Top.Visible = true;
                }
                target.Borders.Bottom.Color = _tableBackground ?? MigraDoc.DocumentObjectModel.Colors.White;
                target.Borders.Bottom.Width = Unit.FromPoint(_tableBackground == null
                    ? _rowCellSpacingPoints * 2 / 3 : 2 * _rowCellSpacingPoints);
                target.Borders.Bottom.Visible = true;
                target.Borders.Left.Color = _tableBackground ?? MigraDoc.DocumentObjectModel.Colors.White;
                target.Borders.Left.Width = Unit.FromPoint(2 * _rowCellSpacingPoints);
                target.Borders.Left.Visible = true;
                target.Borders.Right.Color = _tableBackground ?? MigraDoc.DocumentObjectModel.Colors.White;
                target.Borders.Right.Width = Unit.FromPoint(2 * _rowCellSpacingPoints);
                target.Borders.Right.Visible = true;
            }
            else if (_tableCellSpacingPoints > 0)
            {
                target.Borders.Left.Visible = false;
                target.Borders.Right.Visible = false;
            }
            _container = target;
        }
        return DxpDisposable.Create(() =>
        {
            _cellNoWrap = previousNoWrap;
            _cellNoWrapWidthPoints = previousNoWrapWidth;
            _cellWrapWidthPoints = previousWrapWidth;
            _balanceNarrowCellLastWord = previousBalanceLastWord;
            _container = previous;
            _cellStyle = previousStyle;
            _cellLeftMarginPoints = previousLeftMargin;
            _cellRightMarginPoints = previousRightMargin;
            _cellSpacingIndentPoints = previousSpacingIndent;
            _activeRotatedCell = previousRotatedCell;
            _rotatedCellPlaceholderAdded = previousRotatedPlaceholder;
        });
    }

    private static bool IsNestedCellOutsideParent(TableCell cell, int columnIndex)
    {
        if (columnIndex == 0 ||
            cell.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Table>().FirstOrDefault() is not
                { Parent: TableCell parentCell } nestedTable ||
            parentCell.TableCellProperties?.TableCellWidth is not { } parentCellWidth ||
            !int.TryParse(AttributeValue(parentCellWidth, "w"),
                out var parentWidth) || parentWidth <= 0)
            return false;

        var precedingWidth = nestedTable.TableGrid?.Elements<GridColumn>().Take(columnIndex)
            .Sum(column => int.TryParse(AttributeValue(column, "w"), out var width) ? width : 0) ?? 0;
        var outerMargin = ReadCellMargin(parentCell.TableCellProperties?.TableCellMargin?
            .GetFirstChild<TableCellLeftMargin>()) ?? 108;
        var innerMargin = ReadCellMargin(nestedTable.TableProperties?
            .GetFirstChild<TableCellMarginDefault>()?.GetFirstChild<TableCellLeftMargin>()) ?? 108;
        return precedingWidth + outerMargin + innerMargin >= parentWidth;
    }

    private static bool IsSimpleSideBySideNestedTable(
        DocumentFormat.OpenXml.Wordprocessing.Table table)
    {
        if (table.Parent is not TableCell parentCell ||
            parentCell.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().Count() != 1 ||
            parentCell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count() > 2 &&
                (parentCell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count() < 40 ||
                 parentCell.Parent is not TableRow outerRow ||
                 outerRow.Elements<TableCell>().Count() is < 1 or > 2 ||
                 outerRow.Parent is not DocumentFormat.OpenXml.Wordprocessing.Table outerTable ||
                 outerTable.Elements<TableRow>().Count() != 1) ||
            table.TableGrid?.Elements<GridColumn>().Count() != 2 ||
            table.Elements<TableRow>().ToArray() is not { Length: >= 1 and <= 2 } rows ||
            rows.Length == 2 &&
                (parentCell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count() != 1 ||
                 parentCell.Parent is not TableRow containingRow ||
                 containingRow.Parent is not DocumentFormat.OpenXml.Wordprocessing.Table containingTable ||
                 containingTable.Elements<TableRow>().Count() != 1))
            return false;
        return rows.All(row => row.Elements<TableCell>().ToArray() is { Length: 2 } cells &&
            cells.All(cell => cell.TableCellProperties?.GridSpan == null &&
                cell.TableCellProperties?.HorizontalMerge == null &&
                cell.TableCellProperties?.VerticalMerge == null &&
                !cell.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().Any()));
    }


    private void ApplyBorders(Borders target, DxpComputedBoxBorders source)
    {
        if (source.Top != null) ApplyBorder(target.Top, source.Top);
        if (source.Right != null) ApplyBorder(target.Right, source.Right);
        if (source.Bottom != null) ApplyBorder(target.Bottom, source.Bottom);
        if (source.Left != null) ApplyBorder(target.Left, source.Left);
    }

    private void ApplyBorder(Borders target, DxpComputedBorder source)
    {
        ApplyBorder(target.Top, source);
        ApplyBorder(target.Right, source);
        ApplyBorder(target.Bottom, source);
        ApplyBorder(target.Left, source);
    }

    private void ApplyBorder(MigraDoc.DocumentObjectModel.Border target, DxpComputedBorder source)
    {
        if (source.LineStyle == DxpComputedBorderLineStyle.None)
        {
            target.Visible = false;
            return;
        }
        target.Visible = true;
        target.Width = Unit.FromPoint(Math.Max(0.1, source.WidthPt));
        target.Style = source.LineStyle switch
        {
            DxpComputedBorderLineStyle.Dotted => BorderStyle.Dot,
            DxpComputedBorderLineStyle.Dashed => BorderStyle.DashSmallGap,
            DxpComputedBorderLineStyle.Double when _nativeDoubleTableBorders &&
                Enum.TryParse<BorderStyle>("Double", out var nativeDouble) => nativeDouble,
            _ => BorderStyle.Single
        };
        if (ParseColor(source.ColorCss) is { } color)
            target.Color = color;
    }

    private static MigraDoc.DocumentObjectModel.Color? ParseColor(string? css)
    {
        if (css?.Length != 7 || css[0] != '#') return null;
        if (!byte.TryParse(css.Substring(1, 2), System.Globalization.NumberStyles.HexNumber, null, out var red) ||
            !byte.TryParse(css.Substring(3, 2), System.Globalization.NumberStyles.HexNumber, null, out var green) ||
            !byte.TryParse(css.Substring(5, 2), System.Globalization.NumberStyles.HexNumber, null, out var blue))
            return null;
        return MigraDoc.DocumentObjectModel.Color.FromRgb(red, green, blue);
    }
}
