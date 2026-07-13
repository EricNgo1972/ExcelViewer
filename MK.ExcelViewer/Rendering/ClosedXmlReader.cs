using ClosedXML.Excel;
using MK.ExcelViewer.Rendering.Color;
using MK.ExcelViewer.Rendering.Layout;
using MK.ExcelViewer.Rendering.Model;

namespace MK.ExcelViewer.Rendering;

/// <summary>
/// Caps that bound the emitted HTML — and therefore the browser's DOM, which is what actually gives
/// out first, since every cell becomes a real &lt;td&gt;. Also the backstop against a decompression
/// bomb: a 40 KB zip can declare a ten-million-cell sheet.
/// Defaults here mirror ExcelViewerOptions; the running app always passes its configured values.
/// </summary>
public sealed record ReadOptions
{
    public int MaxRowsPerSheet { get; init; } = 50_000;
    public int MaxColumnsPerSheet { get; init; } = 1_024;
    public int MaxCellsPerSheet { get; init; } = 200_000;

    public bool IncludeHiddenSheets { get; init; }
}

/// <summary>Thrown when the bytes are not a workbook we can read. Becomes a 400 at the API boundary.</summary>
public sealed class WorkbookReadException(string message, Exception? inner = null)
    : Exception(message, inner);

/// <summary>
/// .xlsx → WorkbookModel. Stateless and thread-safe: every parser object is a local.
/// </summary>
public sealed class ClosedXmlReader
{
    public WorkbookModel Read(byte[] bytes, string fileName, ReadOptions options, CancellationToken ct = default)
    {
        // Read the theme from the raw zip rather than trusting the library to surface it — this is
        // the ground truth for the colours in THIS file.
        var theme = ThemePalette.FromWorkbookBytes(bytes);

        XLWorkbook workbook;
        try
        {
            // RecalculateAllFormulas: false — we display cached results and never re-evaluate.
            workbook = new XLWorkbook(new MemoryStream(bytes, writable: false),
                new LoadOptions { RecalculateAllFormulas = false });
        }
        catch (Exception ex)
        {
            throw new WorkbookReadException("The workbook could not be opened — it may be corrupt.", ex);
        }

        using (workbook)
        {
            var styles = new StyleTable();
            var mapper = new StyleMapper(styles, theme);
            var warnings = new List<string>();
            var sheets = new List<SheetModel>();

            foreach (var worksheet in workbook.Worksheets)
            {
                ct.ThrowIfCancellationRequested();

                var visibility = worksheet.Visibility switch
                {
                    XLWorksheetVisibility.Hidden => SheetVisibility.Hidden,
                    XLWorksheetVisibility.VeryHidden => SheetVisibility.VeryHidden,
                    _ => SheetVisibility.Visible,
                };

                if (visibility != SheetVisibility.Visible && !options.IncludeHiddenSheets) continue;

                sheets.Add(ReadSheet(worksheet, sheets.Count, visibility, mapper, theme, options, warnings, ct));
            }

            if (sheets.Count == 0)
                throw new WorkbookReadException("The workbook contains no visible worksheets.");

            return new WorkbookModel
            {
                FileName = fileName,
                Sheets = sheets,
                Styles = styles,
                Warnings = warnings,
            };
        }
    }

    private static SheetModel ReadSheet(
        IXLWorksheet sheet,
        int index,
        SheetVisibility visibility,
        StyleMapper mapper,
        ThemePalette theme,
        ReadOptions options,
        List<string> warnings,
        CancellationToken ct)
    {
        // XLCellsUsedOptions.All, not the default: the default is contents-only, so a shaded but
        // EMPTY block — a highlighted total row waiting to be filled, say — falls outside the used
        // range and its formatting silently disappears.
        var used = sheet.RangeUsed(XLCellsUsedOptions.All);

        if (used is null)   // genuinely empty sheet
        {
            return new SheetModel
            {
                Name = sheet.Name, Index = index, Visibility = visibility,
                FirstRow = 1, LastRow = 1, FirstColumn = 1, LastColumn = 1,
                Columns = [new ColumnModel(1, Metrics.ColumnPx(sheet.ColumnWidth), false, 0)],
                Rows = [],
                ShowGridLines = sheet.ShowGridLines,
                DefaultRowHeightPx = Metrics.RowPx(sheet.RowHeight),
                TotalRowCount = 0,
            };
        }

        var firstRow = used.RangeAddress.FirstAddress.RowNumber;
        var lastRow = used.RangeAddress.LastAddress.RowNumber;
        var firstCol = used.RangeAddress.FirstAddress.ColumnNumber;
        var lastCol = used.RangeAddress.LastAddress.ColumnNumber;

        // Bound the sheet: columns first, then rows against the cell budget, so a wide sheet can't
        // blow the cap by being short.
        if (lastCol - firstCol + 1 > options.MaxColumnsPerSheet)
            lastCol = firstCol + options.MaxColumnsPerSheet - 1;

        var columnCount = lastCol - firstCol + 1;
        var totalRows = lastRow - firstRow + 1;
        var rowBudget = Math.Min(options.MaxRowsPerSheet, Math.Max(1, options.MaxCellsPerSheet / Math.Max(1, columnCount)));

        var truncated = totalRows > rowBudget;
        if (truncated) lastRow = firstRow + rowBudget - 1;

        // ── Merges ───────────────────────────────────────────────────────────────────────────
        // Anchors carry the span; every other cell in the region is marked covered so the builder
        // emits nothing at all for it. Covered cells are materialised explicitly, because a covered
        // cell is usually BLANK and so would never appear in CellsUsed — and a blank we don't know
        // is covered would get an empty <td> and shove the rest of the row sideways.
        var spans = new Dictionary<(int Row, int Col), (int RowSpan, int ColSpan)>();
        var covered = new HashSet<(int Row, int Col)>();

        foreach (var merge in sheet.MergedRanges)
        {
            var mr1 = merge.RangeAddress.FirstAddress.RowNumber;
            var mc1 = merge.RangeAddress.FirstAddress.ColumnNumber;
            var mr2 = Math.Min(merge.RangeAddress.LastAddress.RowNumber, lastRow);
            var mc2 = Math.Min(merge.RangeAddress.LastAddress.ColumnNumber, lastCol);

            if (mr1 > lastRow || mc1 > lastCol || mr2 < mr1 || mc2 < mc1) continue;

            spans[(mr1, mc1)] = (mr2 - mr1 + 1, mc2 - mc1 + 1);

            for (var r = mr1; r <= mr2; r++)
                for (var c = mc1; c <= mc2; c++)
                    if (r != mr1 || c != mc1)
                        covered.Add((r, c));
        }

        // ── Columns ──────────────────────────────────────────────────────────────────────────
        var columns = new List<ColumnModel>(columnCount);
        for (var c = firstCol; c <= lastCol; c++)
        {
            var column = sheet.Column(c);
            columns.Add(new ColumnModel(
                c,
                Metrics.ColumnPx(column.Width),
                column.IsHidden,
                mapper.Map(column.Style)));
        }

        // ── Cells ────────────────────────────────────────────────────────────────────────────
        var byRow = new Dictionary<int, List<CellModel>>();

        foreach (var cell in used.CellsUsed(XLCellsUsedOptions.All))
        {
            ct.ThrowIfCancellationRequested();

            var r = cell.Address.RowNumber;
            var c = cell.Address.ColumnNumber;
            if (r > lastRow || c > lastCol) continue;

            if (!byRow.TryGetValue(r, out var list)) byRow[r] = list = [];

            if (covered.Contains((r, c)))
            {
                list.Add(new CellModel { Row = r, Column = c, Text = "", StyleId = 0, IsMergeCovered = true });
                continue;
            }

            var (text, kind, colorOverride) = ValueFormatter.Format(cell);
            var (rowSpan, colSpan) = spans.TryGetValue((r, c), out var s) ? s : (1, 1);

            list.Add(new CellModel
            {
                Row = r,
                Column = c,
                Text = text,
                Kind = kind,
                StyleId = mapper.Map(cell.Style),
                RowSpan = rowSpan,
                ColSpan = colSpan,
                TextColorOverride = colorOverride,
                Href = HyperlinkOf(cell),
                Comment = cell.HasComment ? cell.GetComment().Text : null,
            });
        }

        // A merge anchor over an empty cell — an empty merged title bar — never shows up in
        // CellsUsed, so materialise it or the merged region collapses.
        foreach (var ((r, c), (rowSpan, colSpan)) in spans)
        {
            if (r < firstRow || r > lastRow) continue;
            if (!byRow.TryGetValue(r, out var list)) byRow[r] = list = [];
            if (list.Any(x => x.Column == c)) continue;

            list.Add(new CellModel
            {
                Row = r, Column = c, Text = "", Kind = CellValueKind.Blank,
                StyleId = mapper.Map(sheet.Cell(r, c).Style),
                RowSpan = rowSpan, ColSpan = colSpan,
            });
        }

        // Every covered cell must be known to the builder, even the ones with no content at all.
        foreach (var (r, c) in covered)
        {
            if (r < firstRow || r > lastRow) continue;
            if (!byRow.TryGetValue(r, out var list)) byRow[r] = list = [];
            if (list.Any(x => x.Column == c)) continue;

            list.Add(new CellModel { Row = r, Column = c, Text = "", StyleId = 0, IsMergeCovered = true });
        }

        // ── Rows ─────────────────────────────────────────────────────────────────────────────
        var rows = new List<RowModel>(byRow.Count);
        foreach (var (r, cells) in byRow.OrderBy(kv => kv.Key))
        {
            var row = sheet.Row(r);
            cells.Sort((a, b) => a.Column.CompareTo(b.Column));

            rows.Add(new RowModel(
                r,
                Metrics.RowPx(row.Height),
                row.IsHidden,
                mapper.Map(row.Style),
                cells));
        }

        CollectWarnings(sheet, warnings);

        // A SPLIT pane is not a frozen one and must not become sticky.
        var view = sheet.SheetView;
        var frozenRows = view.SplitRow;
        var frozenCols = view.SplitColumn;

        return new SheetModel
        {
            Name = sheet.Name,
            Index = index,
            Visibility = visibility,
            FirstRow = firstRow,
            LastRow = lastRow,
            FirstColumn = firstCol,
            LastColumn = lastCol,
            Columns = columns,
            Rows = rows,
            FrozenRows = Math.Max(0, frozenRows),
            FrozenColumns = Math.Max(0, frozenCols),
            ShowGridLines = sheet.ShowGridLines,
            DefaultRowHeightPx = Metrics.RowPx(sheet.RowHeight),
            TabColorHex = ColorResolver.ToHex(sheet.TabColor, theme),
            RowsTruncated = truncated,
            TotalRowCount = totalRows,
        };
    }

    /// <summary>
    /// Excel will happily store a "javascript:" hyperlink, and we inject the result as raw markup —
    /// so the scheme allowlist is load-bearing, not hygiene. Anything else keeps its text and loses
    /// its link.
    /// </summary>
    private static string? HyperlinkOf(IXLCell cell)
    {
        if (!cell.HasHyperlink) return null;

        var link = cell.GetHyperlink();
        if (!link.IsExternal) return null;              // internal jumps have nowhere to go in a static page

        var target = link.ExternalAddress?.ToString();
        if (string.IsNullOrWhiteSpace(target)) return null;

        return Uri.TryCreate(target, UriKind.Absolute, out var uri)
            && uri.Scheme is "http" or "https" or "mailto"
            ? target
            : null;
    }

    /// <summary>
    /// Be honest about what we dropped. A silently missing chart reads as a rendering bug; a stated
    /// one reads as a known limit.
    /// </summary>
    private static void CollectWarnings(IXLWorksheet sheet, List<string> warnings)
    {
        var pictures = sheet.Pictures.Count;
        if (pictures > 0)
            warnings.Add($"“{sheet.Name}” contains {pictures} image{(pictures == 1 ? "" : "s")}, which are not displayed.");

        var conditional = sheet.ConditionalFormats.Count();
        if (conditional > 0)
            warnings.Add($"“{sheet.Name}” uses conditional formatting, which is not applied here — Excel stores the rules, not the resulting colours.");
    }
}
