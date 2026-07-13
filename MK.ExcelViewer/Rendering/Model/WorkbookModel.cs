namespace MK.ExcelViewer.Rendering.Model;

public enum SheetVisibility { Visible, Hidden, VeryHidden }

/// <summary>
/// A workbook with every value and colour already RESOLVED — no ClosedXML types survive into this
/// model, so the HTML builder can never accidentally re-enter the parser or trip over an XLColor.
/// </summary>
public sealed record WorkbookModel
{
    public required string FileName { get; init; }
    public required IReadOnlyList<SheetModel> Sheets { get; init; }

    /// <summary>Shared across ALL sheets — one CSS block per workbook, not one per sheet.</summary>
    public required StyleTable Styles { get; init; }

    /// <summary>Things we could not render faithfully (charts, images, conditional formats).</summary>
    public required IReadOnlyList<string> Warnings { get; init; }
}

public sealed record SheetModel
{
    public required string Name { get; init; }
    public required int Index { get; init; }
    public required SheetVisibility Visibility { get; init; }

    // Used range, 1-based inclusive.
    public required int FirstRow { get; init; }
    public required int LastRow { get; init; }
    public required int FirstColumn { get; init; }
    public required int LastColumn { get; init; }

    /// <summary>Dense over FirstColumn..LastColumn.</summary>
    public required IReadOnlyList<ColumnModel> Columns { get; init; }

    /// <summary>Sparse — only rows that exist. The builder fills gaps with the default height.</summary>
    public required IReadOnlyList<RowModel> Rows { get; init; }

    public int FrozenRows { get; init; }
    public int FrozenColumns { get; init; }
    public bool ShowGridLines { get; init; } = true;
    public double DefaultRowHeightPx { get; init; } = 20;
    public string? TabColorHex { get; init; }

    /// <summary>True when the sheet was larger than the render cap and we cut it off.</summary>
    public bool RowsTruncated { get; init; }

    /// <summary>Pre-truncation row count, for the "showing rows 1–N of M" banner.</summary>
    public int TotalRowCount { get; init; }
}

/// <summary>
/// DefaultStyleId is not decoration: Excel stores whole-column fills, so a shaded column with no
/// cell content in it has styling but no cells. Ignore it and the column renders white.
/// </summary>
public sealed record ColumnModel(int Index, double WidthPx, bool IsHidden, int DefaultStyleId);

public sealed record RowModel(
    int Index,
    double HeightPx,
    bool IsHidden,
    int DefaultStyleId,
    IReadOnlyList<CellModel> Cells);   // sorted ascending by Column, sparse
