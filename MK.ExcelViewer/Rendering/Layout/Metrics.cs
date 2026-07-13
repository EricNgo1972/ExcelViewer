namespace MK.ExcelViewer.Rendering.Layout;

/// <summary>
/// Excel's units → CSS pixels at 96 DPI. These constants are CALIBRATED, not derived: the check
/// that keeps them honest is that Excel's default column (8.43 characters) must come out at
/// exactly 64px and its default row (15pt) at exactly 20px. MetricsTests pins both.
/// </summary>
internal static class Metrics
{
    /// <summary>Max digit width of Calibri 11 at 96 DPI, in px.</summary>
    internal const double MaxDigitWidth = 7.0;

    /// <summary>Excel's per-column cell padding (two 2px gutters plus the 1px gridline), in px.</summary>
    internal const double ColumnPadding = 5.0;

    internal const double DefaultColumnWidthChars = 8.43;
    internal const double DefaultRowHeightPoints = 15.0;

    /// <summary>Points → px. 15pt → 20px.</summary>
    internal static double RowPx(double points) => points * 96.0 / 72.0;

    /// <summary>
    /// Excel-UI character count → px. ClosedXML's IXLColumn.Width is already the number the Excel
    /// UI shows (it has subtracted the 0.71 storage offset on load), so: 8.43 → round(59.01) + 5 = 64px.
    /// </summary>
    internal static double ColumnPx(double chars) =>
        Math.Round(chars * MaxDigitWidth, MidpointRounding.AwayFromZero) + ColumnPadding;

    /// <summary>Indent levels → px. Roughly one max-digit-width plus a gutter per level.</summary>
    internal static double IndentPx(int indent) => Math.Max(0, indent) * 9.0;
}
