using MK.ExcelViewer.Rendering.Html;
using MK.ExcelViewer.Rendering.Model;

namespace MK.ExcelViewer.Rendering;

/// <summary>A workbook, rendered: sheet tabs, one HTML string per sheet, one CSS block for the lot.</summary>
public sealed record RenderedWorkbook
{
    public required string FileName { get; init; }
    public required IReadOnlyList<RenderedSheet> Sheets { get; init; }
    public required string Css { get; init; }
    public required string ScopeId { get; init; }
    public required IReadOnlyList<string> Warnings { get; init; }
}

public sealed record RenderedSheet
{
    public required string Name { get; init; }
    public required string Html { get; init; }
    public required string? TabColorHex { get; init; }
    public required bool RowsTruncated { get; init; }
    public required int ShownRowCount { get; init; }
    public required int TotalRowCount { get; init; }
}

/// <summary>bytes → renderable HTML. Throws <see cref="WorkbookReadException"/> on anything unreadable.</summary>
public sealed class WorkbookRenderer(ClosedXmlReader reader)
{
    public RenderedWorkbook Render(byte[] bytes, string fileName, ReadOptions? options = null, CancellationToken ct = default)
    {
        var model = reader.Read(bytes, fileName, options ?? new ReadOptions(), ct);

        // The CSS is scoped to this id so a workbook's generated rules can't reach the app's chrome.
        var scopeId = "xv-" + Guid.NewGuid().ToString("n")[..8];

        return new RenderedWorkbook
        {
            FileName = model.FileName,
            ScopeId = scopeId,
            Css = StyleSheetBuilder.Build(model, scopeId),
            Warnings = model.Warnings,
            Sheets = model.Sheets.Select(sheet => new RenderedSheet
            {
                Name = sheet.Name,
                Html = SheetHtmlBuilder.Build(sheet, model.Styles),
                TabColorHex = sheet.TabColorHex,
                RowsTruncated = sheet.RowsTruncated,
                ShownRowCount = sheet.LastRow - sheet.FirstRow + 1,
                TotalRowCount = sheet.TotalRowCount,
            }).ToList(),
        };
    }
}
