using System.Globalization;
using MK.ExcelViewer.Rendering.Model;

namespace MK.ExcelViewer.Rendering.Html;

/// <summary>Model enums → CSS fragments.</summary>
internal static class CssMap
{
    internal static string Px(double value) => value.ToString("0.##", CultureInfo.InvariantCulture) + "px";
    internal static string Pt(double value) => value.ToString("0.##", CultureInfo.InvariantCulture) + "pt";

    /// <summary>
    /// CSS has no dash-dot, so those collapse to a dash at the right weight. Excel's three weights
    /// are 1/2/3px. A non-None edge with no colour is Excel's "automatic" — black.
    /// </summary>
    internal static string? Border(BorderEdge edge)
    {
        if (edge.Kind == BorderKind.None) return null;

        var (width, style) = edge.Kind switch
        {
            BorderKind.Hair or BorderKind.Thin => ("1px", "solid"),
            BorderKind.Dotted => ("1px", "dotted"),
            BorderKind.Dashed or BorderKind.DashDot or BorderKind.DashDotDot => ("1px", "dashed"),
            BorderKind.Medium => ("2px", "solid"),
            BorderKind.MediumDashed or BorderKind.MediumDashDot
                or BorderKind.MediumDashDotDot or BorderKind.SlantDashDot => ("2px", "dashed"),
            BorderKind.Thick => ("3px", "solid"),
            BorderKind.Double => ("3px", "double"),
            _ => ("1px", "solid"),
        };

        return $"{width} {style} {edge.ColorHex ?? "#000000"}";
    }

    /// <summary>
    /// Excel's "General" has no CSS equivalent: it right-aligns numbers, centres booleans and
    /// errors, and left-aligns everything else. Resolving it needs the cell's value type, which is
    /// the only reason CellModel carries Kind at all.
    /// </summary>
    internal static string TextAlign(HAlign align, CellValueKind kind) => align switch
    {
        HAlign.Left or HAlign.Fill => "left",
        HAlign.Center or HAlign.CenterContinuous => "center",
        HAlign.Right => "right",
        HAlign.Justify or HAlign.Distributed => "justify",
        _ => kind switch
        {
            CellValueKind.Number or CellValueKind.Date => "right",
            CellValueKind.Boolean or CellValueKind.Error => "center",
            _ => "left",
        },
    };

    internal static string VerticalAlign(VAlign align) => align switch
    {
        VAlign.Top => "top",
        VAlign.Bottom => "bottom",
        _ => "middle",
    };
}
