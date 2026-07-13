using System.Net;
using System.Text;
using MK.ExcelViewer.Rendering.Model;

namespace MK.ExcelViewer.Rendering.Html;

/// <summary>
/// SheetModel → one HTML &lt;table&gt; as a single string.
///
/// A string, not a Razor component tree: a Razor-authored &lt;td&gt; compiles to roughly five
/// RenderTreeFrames which the Blazor Server circuit then RETAINS in server memory for as long as
/// the connection lives, and re-pushes on every reconnect. At a hundred thousand cells that's half
/// a million frames per viewer. As a MarkupString it is exactly one.
/// </summary>
internal static class SheetHtmlBuilder
{
    internal static string Build(SheetModel sheet, StyleTable styles)
    {
        // ~70 bytes/cell with class-based styles.
        var html = new StringBuilder(Math.Min(8 * 1024 * 1024, 4096 + sheet.Rows.Sum(r => r.Cells.Count) * 72));

        var visibleColumns = sheet.Columns.Where(c => !c.IsHidden).ToList();

        // Sticky offsets have to be cumulative pixel positions, so they depend on the widths of the
        // frozen columns that precede them.
        var stickyLeft = new Dictionary<int, double>();
        var offset = 0.0;
        for (var i = 0; i < Math.Min(sheet.FrozenColumns, visibleColumns.Count); i++)
        {
            stickyLeft[visibleColumns[i].Index] = offset;
            offset += visibleColumns[i].WidthPx;
        }

        var rowLookup = sheet.Rows.ToDictionary(r => r.Index);
        var stickyTop = new Dictionary<int, double>();
        offset = 0.0;
        for (var r = sheet.FirstRow; r < sheet.FirstRow + sheet.FrozenRows && r <= sheet.LastRow; r++)
        {
            if (rowLookup.TryGetValue(r, out var row) && row.IsHidden) continue;
            stickyTop[r] = offset;
            offset += rowLookup.TryGetValue(r, out var m) ? m.HeightPx : sheet.DefaultRowHeightPx;
        }

        html.Append("<table class=\"").Append(sheet.ShowGridLines ? "" : "nogrid").Append("\">");

        html.Append("<colgroup>");
        foreach (var column in visibleColumns)
            html.Append("<col style=\"width:").Append(CssMap.Px(column.WidthPx)).Append("\">");
        html.Append("</colgroup><tbody>");

        for (var r = sheet.FirstRow; r <= sheet.LastRow; r++)
        {
            var row = rowLookup.GetValueOrDefault(r);
            if (row is { IsHidden: true }) continue;

            var height = row?.HeightPx ?? sheet.DefaultRowHeightPx;
            html.Append("<tr style=\"height:").Append(CssMap.Px(height)).Append("\">");

            var cells = row?.Cells.ToDictionary(c => c.Column);

            foreach (var column in visibleColumns)
            {
                var cell = cells?.GetValueOrDefault(column.Index);

                // A merge-covered cell gets NO element at all. An empty <td> here would push every
                // following cell in the row one column to the right.
                if (cell is { IsMergeCovered: true }) continue;

                AppendCell(html, styles, row, cell, column, r, stickyTop, stickyLeft);
            }

            html.Append("</tr>");
        }

        html.Append("</tbody></table>");
        return html.ToString();
    }

    private static void AppendCell(
        StringBuilder html,
        StyleTable styles,
        RowModel? row,
        CellModel? cell,
        ColumnModel column,
        int rowIndex,
        Dictionary<int, double> stickyTop,
        Dictionary<int, double> stickyLeft)
    {
        // An absent cell still needs a style: Excel stores whole-row and whole-column formatting, so
        // a shaded empty cell has a colour even though it holds no content. Row formatting wins over
        // column formatting, matching Excel.
        var styleId = cell?.StyleId
            ?? (row?.DefaultStyleId is int rowStyle and not 0 ? rowStyle : column.DefaultStyleId);

        html.Append("<td class=\"s").Append(styleId);

        if (cell is not null && cell.Kind != CellValueKind.Blank)
            html.Append(' ').Append(AlignClass(cell));

        if (cell?.Comment is not null) html.Append(" cmt");
        html.Append('"');

        var isStickyRow = stickyTop.TryGetValue(rowIndex, out var top);
        var isStickyCol = stickyLeft.TryGetValue(column.Index, out var left);

        if (isStickyRow || isStickyCol)
        {
            html.Append(" style=\"position:sticky;");
            if (isStickyRow) html.Append("top:").Append(CssMap.Px(top)).Append(';');
            if (isStickyCol) html.Append("left:").Append(CssMap.Px(left)).Append(';');
            html.Append("z-index:").Append(isStickyRow && isStickyCol ? 3 : 2).Append(';');

            // A sticky cell with no fill of its own is TRANSPARENT, so the rows scrolling underneath
            // show straight through it — the classic sticky-header bug. But force white ONLY when
            // the cell really has no fill: this is an inline style, so it beats the style class, and
            // painting it unconditionally silently erases the fill on every frozen header (a white
            // title bar with white text on it disappears entirely).
            if (styles.Styles[styleId].BackgroundHex is null)
                html.Append("background:#FFFFFF;");

            if (cell?.TextColorOverride is not null)
                html.Append("color:").Append(cell.TextColorOverride).Append(';');

            html.Append('"');
        }
        else if (cell?.TextColorOverride is not null)
        {
            // The [Red] section of a number format — per-cell, so it stays out of the style table.
            html.Append(" style=\"color:").Append(cell.TextColorOverride).Append('"');
        }

        if (cell is { RowSpan: > 1 }) html.Append(" rowspan=\"").Append(cell.RowSpan).Append('"');
        if (cell is { ColSpan: > 1 }) html.Append(" colspan=\"").Append(cell.ColSpan).Append('"');

        if (cell?.Comment is not null)
            html.Append(" title=\"").Append(WebUtility.HtmlEncode(cell.Comment)).Append('"');

        html.Append('>');

        if (cell is not null && cell.Text.Length > 0)
        {
            // The XSS boundary. A cell containing "<script>" is an entirely normal thing for someone
            // to upload, and this output is injected as raw markup.
            var text = WebUtility.HtmlEncode(cell.Text);

            // Rotation needs an inline-block wrapper — a transform on the <td> itself would move the
            // cell, not its text. Rare enough that an inline style costs nothing.
            var rotation = RotationCss(styles.Styles[styleId].TextRotation);
            if (rotation is not null) html.Append("<span style=\"").Append(rotation).Append("\">");

            if (cell.Href is not null)
                html.Append("<a href=\"").Append(WebUtility.HtmlEncode(cell.Href))
                    .Append("\" target=\"_blank\" rel=\"noopener noreferrer\">").Append(text).Append("</a>");
            else
                html.Append(text);

            if (rotation is not null) html.Append("</span>");
        }

        html.Append("</td>");
    }

    /// <summary>
    /// Excel measures rotation counter-clockwise-positive; CSS rotates clockwise-positive. Negate,
    /// or every rotated header tilts the wrong way. 255 is the sentinel for vertically stacked text,
    /// which isn't a rotation at all.
    /// </summary>
    private static string? RotationCss(int rotation) => rotation switch
    {
        0 => null,
        255 => "writing-mode:vertical-rl;text-orientation:upright;",
        _ => $"display:inline-block;transform:rotate({-rotation}deg);",
    };

    /// <summary>
    /// The resolved "General" alignment, as a class. It is emitted on every non-blank cell, but the
    /// per-style .sN rule comes later in the stylesheet at equal specificity — so an explicit
    /// alignment in the workbook still wins, and .al/.ac/.ar only take effect where Excel would
    /// have applied General.
    /// </summary>
    private static string AlignClass(CellModel cell) => CssMap.TextAlign(HAlign.General, cell.Kind) switch
    {
        "right" => "ar",
        "center" => "ac",
        _ => "al",
    };
}
