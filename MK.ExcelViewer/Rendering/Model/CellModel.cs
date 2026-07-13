namespace MK.ExcelViewer.Rendering.Model;

public enum CellValueKind { Blank, Text, Number, Date, Boolean, Error }

public sealed record CellModel
{
    public required int Row { get; init; }        // 1-based, absolute
    public required int Column { get; init; }     // 1-based, absolute

    /// <summary>
    /// Display-ready: the number format is already applied, and a formula cell carries its CACHED
    /// RESULT. A formula string never reaches this field.
    /// </summary>
    public required string Text { get; init; }

    /// <summary>Index into <see cref="StyleTable.Styles"/>.</summary>
    public required int StyleId { get; init; }

    /// <summary>
    /// Exists for exactly one reason: CSS has no equivalent of Excel's "General" alignment, which
    /// right-aligns numbers and left-aligns text. The builder needs the type to resolve it.
    /// </summary>
    public CellValueKind Kind { get; init; }

    public int RowSpan { get; init; } = 1;
    public int ColSpan { get; init; } = 1;

    /// <summary>
    /// Swallowed by another cell's merge. The builder must emit NOTHING for it — not even an empty
    /// &lt;td&gt;, which would shove the rest of the row one column to the right.
    /// </summary>
    public bool IsMergeCovered { get; init; }

    public string? Href { get; init; }
    public string? Comment { get; init; }

    /// <summary>
    /// Set when the number format carries a colour section — e.g. "[Red]-#,##0.00" paints negative
    /// numbers red. ClosedXML's formatter applies the format but drops its colour, so without this
    /// every red-negative accounting sheet renders in black.
    /// </summary>
    public string? TextColorOverride { get; init; }
}
