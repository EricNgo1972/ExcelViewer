namespace MK.ExcelViewer.Rendering.Model;

public enum HAlign { General, Left, Center, Right, Fill, Justify, CenterContinuous, Distributed }
public enum VAlign { Top, Center, Bottom, Justify, Distributed }
public enum UnderlineKind { None, Single, Double }

public enum BorderKind
{
    None, Hair, Thin, Dotted, Dashed, DashDot, DashDotDot,
    Medium, MediumDashed, MediumDashDot, MediumDashDotDot, SlantDashDot,
    Thick, Double,
}

/// <summary>A null ColorHex on a non-None edge means Excel's "automatic" — the builder draws it black.</summary>
public readonly record struct BorderEdge(BorderKind Kind, string? ColorHex)
{
    public static readonly BorderEdge None = new(BorderKind.None, null);
}

public sealed record FontStyle
{
    public required string Family { get; init; }
    public required double SizePt { get; init; }
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public UnderlineKind Underline { get; init; }
    public bool Strike { get; init; }
    public string? ColorHex { get; init; }
}

public sealed record CellStyle
{
    public int FontId { get; init; }
    public string? BackgroundHex { get; init; }        // null = no fill
    public BorderEdge Top { get; init; }
    public BorderEdge Right { get; init; }
    public BorderEdge Bottom { get; init; }
    public BorderEdge Left { get; init; }
    public HAlign HAlign { get; init; }
    public VAlign VAlign { get; init; } = VAlign.Bottom;   // Excel's default is Bottom, not Top
    public bool WrapText { get; init; }

    /// <summary>Signed degrees, counter-clockwise-positive (Excel's convention). 255 = stacked.</summary>
    public int TextRotation { get; init; }
    public int Indent { get; init; }
}

/// <summary>
/// Interns styles so cells can carry an int instead of a copy. A heavily formatted workbook has
/// only a few hundred distinct styles, so the emitted HTML gets class="s42" (11 bytes) where an
/// inline style= would be ~200. That ratio is what makes a 100k-cell sheet renderable at all.
///
/// This only works because CellStyle and FontStyle are records of SCALARS. Add a collection member
/// to either and C# record equality silently degrades to reference equality, every cell interns a
/// fresh style, and the dedup — and the size budget — quietly collapses.
/// </summary>
public sealed class StyleTable
{
    private readonly Dictionary<CellStyle, int> _styleIndex = new();
    private readonly List<CellStyle> _styles = [];
    private readonly Dictionary<FontStyle, int> _fontIndex = new();
    private readonly List<FontStyle> _fonts = [];

    public IReadOnlyList<CellStyle> Styles => _styles;
    public IReadOnlyList<FontStyle> Fonts => _fonts;

    public StyleTable()
    {
        InternFont(new FontStyle { Family = "Calibri", SizePt = 11 });   // font 0 = default
        Intern(new CellStyle { FontId = 0 });                            // style 0 = default
    }

    public int Intern(CellStyle style)
    {
        if (_styleIndex.TryGetValue(style, out var id)) return id;
        id = _styles.Count;
        _styles.Add(style);
        _styleIndex[style] = id;
        return id;
    }

    public int InternFont(FontStyle font)
    {
        if (_fontIndex.TryGetValue(font, out var id)) return id;
        id = _fonts.Count;
        _fonts.Add(font);
        _fontIndex[font] = id;
        return id;
    }
}
