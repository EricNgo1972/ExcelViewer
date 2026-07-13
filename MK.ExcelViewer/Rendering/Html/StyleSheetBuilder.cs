using System.Text;
using MK.ExcelViewer.Rendering.Layout;
using MK.ExcelViewer.Rendering.Model;

namespace MK.ExcelViewer.Rendering.Html;

/// <summary>
/// StyleTable → one CSS block per workbook. Every rule is scoped under the viewport's id so a
/// workbook's styles can never leak out into the app's own chrome.
/// </summary>
internal static class StyleSheetBuilder
{
    internal static string Build(WorkbookModel workbook, string scopeId)
    {
        var css = new StringBuilder(16 * 1024);
        var scope = "#" + scopeId;

        // The grid colour is a variable so that ShowGridLines=false is a one-token change per sheet
        // rather than a different rule for every style.
        css.Append(scope).Append("{--xv-grid:#D4D4D4;}\n");
        css.Append(scope).Append(" .nogrid{--xv-grid:transparent;}\n");

        css.Append(scope).Append(" table{border-collapse:collapse;table-layout:fixed;")
           .Append("font-family:").Append(FontStacks.For("Calibri")).Append(";font-size:11pt;}\n");

        // white-space:pre — Excel does not collapse runs of spaces, and neither should we.
        css.Append(scope).Append(" td{padding:0 2px;overflow:hidden;white-space:pre;vertical-align:bottom;")
           .Append("border:1px solid var(--xv-grid);}\n");

        // Resolved "General" alignment. Kept out of the style classes so that one Excel style
        // doesn't fork into three model styles and wreck the dedup.
        css.Append(scope).Append(" .al{text-align:left;}").Append(scope).Append(" .ac{text-align:center;}")
           .Append(scope).Append(" .ar{text-align:right;}\n");

        css.Append(scope).Append(" .cmt{position:relative;}\n");
        css.Append(scope).Append(" .cmt::after{content:'';position:absolute;top:0;right:0;border-top:5px solid #E33;border-left:5px solid transparent;}\n");

        for (var i = 0; i < workbook.Styles.Styles.Count; i++)
            AppendStyle(css, scope, i, workbook.Styles.Styles[i], workbook.Styles.Fonts);

        return css.ToString();
    }

    private static void AppendStyle(StringBuilder css, string scope, int id, CellStyle style, IReadOnlyList<FontStyle> fonts)
    {
        var font = fonts[style.FontId];

        css.Append(scope).Append(" .s").Append(id).Append('{');

        css.Append("font-family:").Append(FontStacks.For(font.Family)).Append(';');
        css.Append("font-size:").Append(CssMap.Pt(font.SizePt)).Append(';');
        if (font.Bold) css.Append("font-weight:700;");
        if (font.Italic) css.Append("font-style:italic;");
        if (font.ColorHex is not null) css.Append("color:").Append(font.ColorHex).Append(';');

        // Underline and strikethrough are the same CSS property — emitted separately, the second
        // silently overwrites the first.
        var decorations = (font.Underline, font.Strike) switch
        {
            (UnderlineKind.None, false) => null,
            (UnderlineKind.None, true) => "line-through",
            (UnderlineKind.Double, true) => "underline line-through double",
            (UnderlineKind.Double, false) => "underline double",
            (_, true) => "underline line-through",
            (_, false) => "underline",
        };
        if (decorations is not null) css.Append("text-decoration:").Append(decorations).Append(';');

        if (style.BackgroundHex is not null) css.Append("background-color:").Append(style.BackgroundHex).Append(';');

        // Excel draws a gridline only where a cell has neither an explicit border nor a fill. A
        // filled cell with no border shows no gridline — which is exactly why a filled block looks
        // like a solid slab in Excel and like graph paper in a naive converter.
        AppendEdge(css, "border-top", style.Top, style.BackgroundHex);
        AppendEdge(css, "border-right", style.Right, style.BackgroundHex);
        AppendEdge(css, "border-bottom", style.Bottom, style.BackgroundHex);
        AppendEdge(css, "border-left", style.Left, style.BackgroundHex);

        if (style.HAlign != HAlign.General)
            css.Append("text-align:").Append(CssMap.TextAlign(style.HAlign, CellValueKind.Text)).Append(';');

        css.Append("vertical-align:").Append(CssMap.VerticalAlign(style.VAlign)).Append(';');

        if (style.WrapText) css.Append("white-space:pre-wrap;overflow-wrap:break-word;");

        if (style.Indent > 0)
        {
            var side = style.HAlign == HAlign.Right ? "padding-right" : "padding-left";
            css.Append(side).Append(':').Append(CssMap.Px(Metrics.IndentPx(style.Indent))).Append(';');
        }

        css.Append("}\n");
    }

    private static void AppendEdge(StringBuilder css, string property, BorderEdge edge, string? background)
    {
        var border = CssMap.Border(edge);

        if (border is not null) css.Append(property).Append(':').Append(border).Append(';');
        else if (background is not null) css.Append(property).Append(":none;");
        // else: leave the default `1px solid var(--xv-grid)` from the td rule — that's the gridline.
    }
}
