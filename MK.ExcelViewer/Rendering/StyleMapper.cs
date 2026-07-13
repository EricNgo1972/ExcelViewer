using ClosedXML.Excel;
using MK.ExcelViewer.Rendering.Color;
using MK.ExcelViewer.Rendering.Model;

namespace MK.ExcelViewer.Rendering;

/// <summary>
/// IXLStyle → the parser-agnostic CellStyle, interned into the workbook's StyleTable.
///
/// Mapping a style is ~20 property reads, which is the dominant cost on a large sheet. XLStyle has
/// value equality (it delegates to an interned XLStyleValue), so caching on the source style means
/// identically-styled cells map once. If that ever regresses the only cost is speed — correctness
/// is unaffected.
/// </summary>
internal sealed class StyleMapper(StyleTable table, ThemePalette theme)
{
    private readonly Dictionary<IXLStyle, int> _cache = [];

    internal int Map(IXLStyle style)
    {
        if (_cache.TryGetValue(style, out var cached)) return cached;

        var font = table.InternFont(new FontStyle
        {
            Family = style.Font.FontName ?? "Calibri",
            SizePt = style.Font.FontSize,
            Bold = style.Font.Bold,
            Italic = style.Font.Italic,
            Strike = style.Font.Strikethrough,
            Underline = style.Font.Underline switch
            {
                XLFontUnderlineValues.None => UnderlineKind.None,
                XLFontUnderlineValues.Double or XLFontUnderlineValues.DoubleAccounting => UnderlineKind.Double,
                _ => UnderlineKind.Single,   // accounting underlines degrade to a plain one
            },
            ColorHex = ColorResolver.ToHex(style.Font.FontColor, theme),
        });

        var id = table.Intern(new CellStyle
        {
            FontId = font,

            // PatternType None means the cell has NO fill — not a white one. Emitting white here
            // would paint over the alternating-row look of every themed table.
            BackgroundHex = style.Fill.PatternType == XLFillPatternValues.None
                ? null
                : ColorResolver.ToHex(style.Fill.BackgroundColor, theme),

            Top = Edge(style.Border.TopBorder, style.Border.TopBorderColor),
            Right = Edge(style.Border.RightBorder, style.Border.RightBorderColor),
            Bottom = Edge(style.Border.BottomBorder, style.Border.BottomBorderColor),
            Left = Edge(style.Border.LeftBorder, style.Border.LeftBorderColor),

            HAlign = style.Alignment.Horizontal switch
            {
                XLAlignmentHorizontalValues.Left => HAlign.Left,
                XLAlignmentHorizontalValues.Center => HAlign.Center,
                XLAlignmentHorizontalValues.Right => HAlign.Right,
                XLAlignmentHorizontalValues.Fill => HAlign.Fill,
                XLAlignmentHorizontalValues.Justify => HAlign.Justify,
                XLAlignmentHorizontalValues.CenterContinuous => HAlign.CenterContinuous,
                XLAlignmentHorizontalValues.Distributed => HAlign.Distributed,
                _ => HAlign.General,
            },
            VAlign = style.Alignment.Vertical switch
            {
                XLAlignmentVerticalValues.Top => VAlign.Top,
                XLAlignmentVerticalValues.Center => VAlign.Center,
                XLAlignmentVerticalValues.Justify => VAlign.Justify,
                XLAlignmentVerticalValues.Distributed => VAlign.Distributed,
                _ => VAlign.Bottom,
            },
            WrapText = style.Alignment.WrapText,
            TextRotation = NormalizeRotation(style.Alignment.TextRotation),
            Indent = style.Alignment.Indent,
        });

        _cache[style] = id;
        return id;
    }

    private BorderEdge Edge(XLBorderStyleValues border, XLColor color)
    {
        var kind = border switch
        {
            XLBorderStyleValues.None => BorderKind.None,
            XLBorderStyleValues.Hair => BorderKind.Hair,
            XLBorderStyleValues.Thin => BorderKind.Thin,
            XLBorderStyleValues.Dotted => BorderKind.Dotted,
            XLBorderStyleValues.Dashed => BorderKind.Dashed,
            XLBorderStyleValues.DashDot => BorderKind.DashDot,
            XLBorderStyleValues.DashDotDot => BorderKind.DashDotDot,
            XLBorderStyleValues.Medium => BorderKind.Medium,
            XLBorderStyleValues.MediumDashed => BorderKind.MediumDashed,
            XLBorderStyleValues.MediumDashDot => BorderKind.MediumDashDot,
            XLBorderStyleValues.MediumDashDotDot => BorderKind.MediumDashDotDot,
            XLBorderStyleValues.SlantDashDot => BorderKind.SlantDashDot,
            XLBorderStyleValues.Thick => BorderKind.Thick,
            XLBorderStyleValues.Double => BorderKind.Double,
            _ => BorderKind.None,
        };

        return kind == BorderKind.None ? BorderEdge.None : new BorderEdge(kind, ColorResolver.ToHex(color, theme));
    }

    /// <summary>
    /// Excel stores rotation as 0–90 counter-clockwise, then 91–180 meaning CLOCKWISE by (v − 90),
    /// and 255 for vertically stacked text. Normalise to signed CCW degrees so the CSS layer has
    /// one convention to deal with; miss this and every rotated header tilts the wrong way.
    /// </summary>
    internal static int NormalizeRotation(int raw) => raw switch
    {
        255 => 255,                 // stacked — a sentinel, not an angle
        <= 90 => raw,
        <= 180 => -(raw - 90),
        _ => 0,
    };
}
